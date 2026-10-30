import os
import io
import json
import re
from typing import List, Dict, Any, Optional, Callable
from PIL import Image
import fitz  # PyMuPDF

# 신규 google-genai SDK 및 기존 SDK 모두 지원
try:
    from google import genai
    from google.genai import types
    NEW_GENAI = True
except ImportError:
    NEW_GENAI = False
    try:
        import google.generativeai as legacy_genai
    except ImportError:
        legacy_genai = None

from sample_data import SAMPLE_RECEIPTS

def extract_images_from_pdf(pdf_bytes: bytes, dpi: int = 150) -> List[Image.Image]:
    """PDF 바이트에서 각 페이지를 고해상도 PIL 이미지 리스트로 렌더링"""
    images = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for page in doc:
        pix = page.get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        images.append(img)
    return images

def calculate_attendees(total_quantity: int, total_amount: int) -> str:
    """
    금액과 구매한 개수를 기반으로 대상인원을 합리적으로 자동 산출
    - 음료, 베이커리, 다과류는 기본적으로 구매 개수가 인원수와 밀접하게 연관됨.
    - 금액 대비 개수 비율을 검토하여 1인당 10,000원~15,000원 수준을 보정 기준으로 활용.
    """
    if total_quantity <= 0:
        if total_amount > 0:
            est = max(1, round(total_amount / 12000))
            return f"{est}명"
        return "1명"
    
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
당신은 한국의 법인카드 영수증 및 증빙서 전문 분석 AI입니다.
스캔된 이미지에는 1장 또는 2장 이상의 영수증과 상단에 증빙서 첨부지가 포함되어 있을 수 있습니다.

이미지를 세밀하게 관찰하고 다음 사항을 엄격히 준수하여 JSON으로 추출하세요:
1. [공란 원칙] 영수증이나 이미지에서 명확하게 확인되지 않는 정보(특히 사용자명 user_name, 소속부서 department 등)는 절대로 임의로 추정하지 말고 반드시 빈 문자열("")(공란)으로 설정하세요.
2. 영수증이 2장 이상 나란히 붙어있다면 좌측 영수증부터 우측 영수증 순으로 각각 분리하여 추출하세요.
3. [중요] 사용목적 (purpose): 영수증 상단이나 여백에 볼펜/매직으로 직접 '손글씨(수기)'로 적힌 메모 내용(예: '이정복 보건소', '영등포 보건소', '동작구 보건소', '양주시 보건소', 병원명, 회의명 등)을 반드시 정확히 읽어내어 입력하세요. 손글씨가 없으면 빈 문자열("")로 두세요.
4. 가맹점명(store_name), 사용일자(date, YYYYMMDD), 사용시각(time, HH:MM:SS), 품목 리스트(items), 총 구매 수량(total_quantity), 최종 결제금액(amount), 카드번호 끝4자리(card_last4)를 추출하세요.

반드시 다음 JSON 형식으로만 응답하세요:
{
  "receipts": [
    {
      "receipt_sub_index": 1,
      "date": "20261001",
      "time": "12:50:09",
      "store_name": "개포행복 파리바게뜨",
      "purpose": "이정복 보건소",
      "items": [
        {"name": "호두파운드1호", "quantity": 2, "price": 31800},
        {"name": "말차)명가카스테라", "quantity": 1, "price": 15900}
      ],
      "total_quantity": 3,
      "amount": 47700,
      "card_last4": "9716",
      "user_name": "",
      "department": "",
      "notes": "손글씨 메모 인식됨"
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
        # 혹시 앞뒤로 부연 설명이 섞인 경우 정규식으로 JSON 추출
        match = re.search(r'(\{[\s\S]*\})', clean_text)
        if match:
            return json.loads(match.group(1))
        raise ValueError(f"AI 응답에서 JSON 형식을 추출할 수 없습니다: {clean_text[:200]}")

def parse_receipt_page_with_gemini(image: Image.Image, api_key: str, page_num: int = 1) -> List[Dict[str, Any]]:
    """Gemini Vision API를 호출하여 영수증 및 수기 메모 추출 (여러 모델 및 SDK fallback)"""
    prompt = "이 이미지에서 영수증들을 분리 분석하고, 특히 영수증 위에 손글씨(수기)로 적힌 메모(보건소명/병원명 등)를 빠짐없이 추출하여 지정된 JSON으로 응답해줘."
    
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='JPEG', quality=95)
    img_bytes = img_byte_arr.getvalue()

    raw_text = None
    last_error = None

    # 시도할 모델 순서
    models_to_try = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash']

    # 1. New Google GenAI SDK 우선 시도
    if NEW_GENAI:
        try:
            client = genai.Client(api_key=api_key)
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
                    last_error = ex
                    continue
        except Exception as e:
            last_error = e

    # 2. Legacy google.generativeai SDK fallback
    if not raw_text and legacy_genai is not None:
        try:
            legacy_genai.configure(api_key=api_key)
            for m_name in ['gemini-1.5-flash', 'gemini-1.5-pro']:
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
                    last_error = ex
                    continue
        except Exception as e:
            last_error = e

    if not raw_text:
        raise RuntimeError(f"Gemini API 호출에 실패했습니다: {last_error}")

    data = _extract_json_from_text(raw_text)
    receipts_list = data.get("receipts", [])

    # 대상인원 및 부가정보 계산
    for idx, r in enumerate(receipts_list, start=1):
        r["page_index"] = page_num
        r["receipt_sub_index"] = idx
        t_qty = r.get("total_quantity")
        if not t_qty and "items" in r:
            t_qty = sum(item.get("quantity", 1) for item in r["items"])
            r["total_quantity"] = t_qty
        
        t_amt = r.get("amount", 0)
        r["attendees"] = calculate_attendees(int(t_qty or 1), int(t_amt or 0))

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
    - API 키가 없거나 분석 실패 시 명확한 예외를 발생시켜 사용자에게 알립니다.
    """
    if not api_key or not str(api_key).strip():
        raise ValueError("Gemini API 키가 입력되지 않았습니다. 사이드바에 API 키를 입력하시거나 Streamlit Secrets에 GEMINI_API_KEY를 등록해 주세요.")

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
