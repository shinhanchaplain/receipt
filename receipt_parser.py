import os
import io
import json
import re
from typing import List, Dict, Any, Optional, Callable
from PIL import Image
import fitz  # PyMuPDF

# SDK 안전한 임포트
legacy_genai = None
try:
    import google.generativeai as legacy_genai
except ImportError:
    pass

NEW_GENAI = False
try:
    from google import genai
    from google.genai import types
    NEW_GENAI = True
except ImportError:
    pass

from sample_data import SAMPLE_RECEIPTS

def extract_images_from_pdf(pdf_bytes: bytes, dpi: int = 220) -> List[Image.Image]:
    """PDF 바이트에서 각 페이지를 초고해상도(220 DPI) PIL 이미지 리스트로 렌더링하여 작은 글씨 및 필기체 인식률 극대화"""
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        images.append(img)
    return images

def calculate_attendees(total_quantity: int, total_amount: int) -> str:
    """
    금액과 구매한 개수를 기반으로 대상인원을 산출 (확인 불가능한 경우 공란 처리 지원)
    """
    if total_quantity <= 0 and total_amount <= 0:
        return ""
    
    if total_quantity <= 0:
        if total_amount > 0:
            est = max(1, round(total_amount / 12000))
            return f"{est}명"
        return ""
    
    avg_per_item = total_amount / total_quantity if total_quantity > 0 else 0
    
    if 1 <= total_quantity <= 15:
        if total_quantity == 1 and total_amount >= 30000:
            est = max(2, round(total_amount / 15000))
            return f"{est}명"
        if avg_per_item > 25000:
            est = max(total_quantity, round(total_amount / 15000))
            return f"{est}명"
        return f"{total_quantity}명"
    else:
        est = max(1, round(total_amount / 12000))
        return f"{est}명"

SYSTEM_INSTRUCTION = """
당신은 대한민국 공공기관 및 대학교/기업의 법인카드 및 업무추진비 영수증 전문 최고 수준의 비전 AI입니다.
스캔된 이미지에는 1장 이상의 영수증(신용카드 전표, 백화점/마트 영수증, POS 영수증, 간이영수증 등) 또는 상단 증빙서 첨부지가 포함되어 있습니다.

이미지를 세밀하게 관찰하고 다음 원칙을 반드시 준수하여 JSON으로 추출하세요:

1. [영수증 완벽 분리]
   - 한 페이지에 2장 이상의 영수증이 좌/우 또는 상/하로 부착되어 있다면, 각 영수증의 외곽 테두리와 경계선을 정확히 파악하여 1건씩 누락 없이 개별 분리 추출하세요. (보통 좌측 영수증 -> 우측 영수증 순서)

2. [수기(손글씨) 메모 최우선 정밀 식별]
   - 사용목적/집행목적(purpose): 영수증 상단, 하단 여백, 또는 영수증 본체 위에 볼펜/사인펜/연필로 기재된 손글씨 메모(예: '김순 학과장', '교원, 학과장', '예비대', '보건소명', '회의명', '간담회' 등)를 세밀하게 읽어내어 그대로 입력하세요.
   - 단, 손글씨 메모가 전혀 없는 영수증은 절대로 지어내지 말고 반드시 빈 문자열("")(공란)으로 설정하세요.

3. [가맹점 상호명 정확 추출]
   - 사업자등록번호 옆의 대표자 개인 성명이 아니라, 실제 매장명/상호명(예: '문수반', '롯데백화점(미아점)', '에센브르제과', '신토불이갈비', '하이커피', '럭키마트', '뚜레쥬르 본점', '삼송부대&양푼김치탕' 등)을 정확히 추출하세요.

4. [일시 및 금액, 카드번호 정밀 추출]
   - 승인일자(date): 영수증에 기재된 결제일자 (YYYYMMDD 숫자 8자리 형식, 예: '20261008')
   - 승인시각(time): 결제시각 (HH:MM:SS 또는 HH:MM 형식, 예: '12:04:00')
   - 최종 결제금액(amount): 할인이 적용된 경우 최종 결제/승인 합계금액 (순수 숫자 정수, 예: 102200)
   - 카드번호 끝4자리(card_last4): 영수증 또는 첨부지에 기재된 카드번호 끝 4자리 (예: '4806', '0480' 등). 확인 불가 시 빈 문자열("")
   - 사용자명(user_name): 실물에 명확히 없으면 빈 문자열("")

반드시 아래와 같은 JSON 구조로만 응답하세요:
{
  "receipts": [
    {
      "receipt_sub_index": 1,
      "date": "20261008",
      "time": "12:04:00",
      "store_name": "롯데백화점(미아점)",
      "purpose": "",
      "items": [
        {"name": "상품", "quantity": 1, "price": 48000},
        {"name": "케이크류", "quantity": 1, "price": 29000}
      ],
      "total_quantity": 2,
      "amount": 102200,
      "card_last4": "0480",
      "user_name": ""
    }
  ]
}
"""

def _extract_json_from_text(raw_text: str) -> dict:
    """LLM 응답 텍스트에서 안전하게 JSON 파싱"""
    clean_text = raw_text.strip()
    if clean_text.startswith("```json"):
        clean_text = clean_text[7:]
    if clean_text.startswith("```"):
        clean_text = clean_text[3:]
    if clean_text.endswith("```"):
        clean_text = clean_text[:-3]
    clean_text = clean_text.strip()

    try:
        return json.loads(clean_text)
    except Exception:
        match = re.search(r'(\{[\s\S]*\})', clean_text)
        if match:
            return json.loads(match.group(1))
        raise ValueError(f"AI 응답에서 JSON 형식을 추출할 수 없습니다: {clean_text[:200]}")

def parse_receipt_page_with_gemini(image: Image.Image, api_key: str, page_num: int = 1) -> List[Dict[str, Any]]:
    """Gemini Vision API를 호출하여 영수증 분리 및 필드 추출"""
    prompt = (
        "이 고해상도 영수증 이미지에 포함된 모든 영수증을 누락 없이 좌측/우측 순서대로 분리 추출해줘. "
        "특히 영수증 상단이나 여백에 적힌 손글씨/볼펜 메모(교원명, 학과장명, 목적 등)를 정밀하게 읽어 purpose에 넣고, "
        "상호명(가맹점명), 결제일자(YYYYMMDD), 결제시각, 최종 결제금액, 카드번호 끝4자리를 정확하게 추출하여 지정된 JSON으로 응답해줘."
    )
    
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='JPEG', quality=95)
    img_bytes = img_byte_arr.getvalue()

    raw_text = None
    errors = []

    # 1. New Google GenAI SDK 시도 (최신 권장 gemini-3.8-flash 모델 적용)
    if NEW_GENAI:
        try:
            client = genai.Client(api_key=api_key)
            models_to_try = ['gemini-3.8-flash', 'gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash']
            for m_name in models_to_try:
                try:
                    response = client.models.generate_content(
                        model=m_name,
                        contents=[
                            types.Part.from_bytes(data=img_bytes, mime_type='image/jpeg'),
                            prompt
                        ],
                        config=types.GenerateContentConfig(
                            system_instruction=SYSTEM_INSTRUCTION,
                            response_mime_type="application/json",
                            temperature=0.1
                        )
                    )
                    if response and response.text:
                        raw_text = response.text
                        break
                except Exception as ex:
                    errors.append(f"{m_name}: {ex}")
                    continue
        except Exception as e:
            errors.append(f"genai.Client: {e}")

    # 2. Legacy google.generativeai SDK fallback
    if not raw_text and legacy_genai is not None:
        try:
            legacy_genai.configure(api_key=api_key)
            for m_name in ['gemini-3.8-flash', 'gemini-2.5-flash', 'gemini-1.5-flash', 'gemini-1.5-pro']:
                try:
                    model = legacy_genai.GenerativeModel(
                        model_name=m_name,
                        system_instruction=SYSTEM_INSTRUCTION,
                        generation_config={"response_mime_type": "application/json", "temperature": 0.1}
                    )
                    response = model.generate_content([prompt, image])
                    if response and response.text:
                        raw_text = response.text
                        break
                except Exception as ex:
                    errors.append(f"legacy {m_name}: {ex}")
                    continue
        except Exception as e:
            errors.append(f"legacy config: {e}")

    if not raw_text:
        err_msg = " | ".join(errors) if errors else "알 수 없는 연결 오류"
        raise RuntimeError(f"Gemini API 호출에 실패했습니다 ({err_msg}). API 키와 네트워크 상태를 확인해주세요.")

    data = _extract_json_from_text(raw_text)
    receipts_list = data.get("receipts", [])

    # 대상인원 및 부가정보 계산
    for idx, r in enumerate(receipts_list, start=1):
        r["page_index"] = page_num
        r["receipt_sub_index"] = idx
        
        # 공란 정리
        if not r.get("purpose"):
            r["purpose"] = ""
        if not r.get("user_name"):
            r["user_name"] = ""
            
        t_qty = r.get("total_quantity", 0)
        if not t_qty and "items" in r and r["items"]:
            t_qty = sum(item.get("quantity", 1) for item in r["items"])
            r["total_quantity"] = t_qty
        
        t_amt = r.get("amount", 0)
        r["attendees"] = calculate_attendees(int(t_qty or 0), int(t_amt or 0))

    return receipts_list

def parse_receipt_document(
    file_bytes: bytes,
    file_name: str,
    api_key: Optional[str] = None,
    progress_callback: Optional[Callable[[int, int, str], None]] = None
) -> List[Dict[str, Any]]:
    """
    PDF 또는 이미지 파일을 받아 전체 영수증 리스트를 반환합니다.
    - 실제 파일 분석 시 반드시 유효한 api_key가 필요합니다.
    """
    if not api_key or not str(api_key).strip():
        raise ValueError("Gemini API 키가 입력되지 않았습니다. 좌측 사이드바에 API 키를 입력하시거나 Streamlit Secrets에 GEMINI_API_KEY를 등록해 주세요.")

    is_pdf = file_name.lower().endswith(".pdf")
    if is_pdf:
        images = extract_images_from_pdf(file_bytes)
    else:
        images = [Image.open(io.BytesIO(file_bytes))]

    total_pages = len(images)
    if total_pages == 0:
        raise ValueError("파일에서 분석할 수 있는 이미지를 찾지 못했습니다.")

    all_receipts = []
    current_seq = 1

    for page_idx, img in enumerate(images, start=1):
        if progress_callback:
            progress_callback(page_idx, total_pages, f"{page_idx}/{total_pages} 페이지 영수증 및 수기 메모 분석 중...")

        page_receipts = parse_receipt_page_with_gemini(img, api_key=api_key.strip(), page_num=page_idx)
        for r in page_receipts:
            r["seq"] = current_seq
            current_seq += 1
            all_receipts.append(r)

    if not all_receipts:
        raise ValueError("영수증 이미지에서 내역을 인식하지 못했습니다. 이미지가 선명한지 확인해 주세요.")

    return all_receipts
