import os
import io
import datetime
from copy import copy
import openpyxl
from openpyxl.styles import Alignment, Font, Border, PatternFill

# 유연한 헤더 매핑을 위한 키워드 사전 (소문자, 공백제거 비교)
HEADER_KEYWORDS = {
    "seq": ["연번", "순번", "no", "번호", "순서", "일련번호"],
    "date": ["사용일자", "사용일시", "일자", "날짜", "승인일자", "승인일시", "결제일자"],
    "time": ["사용시간", "사용시각", "시간", "시각", "승인시간"],
    "store_name": ["사용장소", "가맹점", "가맹점명", "상호", "상호명", "장소", "구입처", "거래처", "물품명"],
    "purpose": ["사용목적", "목적", "사용내역", "적요", "지출목적", "구체적명기", "내용", "비고"],
    "attendees": ["대상인원", "인원", "수량", "인원수", "참석자", "인원(수량)", "참석인원"],
    "card_last4": ["카드번호", "카드", "카드끝", "끝4자리", "카드번호(끝4자리)", "카드끝자리"],
    "amount": ["금액", "결제금액", "금 액", "지출액", "사용금액", "합계금액", "승인금액"],
    "user_name": ["사용자", "담당자", "소지자", "사용자명", "성명", "작성자"]
}

def clean_str(val):
    if val is None:
        return ""
    return str(val).replace("\n", "").replace(" ", "").replace("_", "").lower()

def format_time_value(time_str):
    if isinstance(time_str, datetime.time):
        return time_str
    if not time_str:
        return ""
    try:
        parts = [int(p) for p in str(time_str).strip().split(':')]
        if len(parts) == 3:
            return datetime.time(parts[0], parts[1], parts[2])
        elif len(parts) == 2:
            return datetime.time(parts[0], parts[1], 0)
    except Exception:
        pass
    return str(time_str)

def format_date_value(date_str):
    if not date_str:
        return ""
    clean = str(date_str).replace("-", "").replace(".", "").replace("/", "").strip()
    try:
        return int(clean)
    except Exception:
        return clean

def find_target_sheet(wb):
    """법인카드 관련 시트 탐색 (없으면 활성 시트 반환)"""
    candidate_keywords = ["법인카드", "사용내역", "지출", "경비", "카드"]
    for sheetname in wb.sheetnames:
        clean_name = clean_str(sheetname)
        if any(kw in clean_name for kw in candidate_keywords):
            return wb[sheetname]
    return wb.active

def detect_column_mapping_and_start_row(ws):
    """
    엑셀 시트를 상단부터 스캔하여 헤더 행과 각 컬럼(열)의 위치를 유연하게 탐지합니다.
    양식이 조금씩 달라도 키워드를 찾아 매핑합니다.
    """
    header_row_idx = None
    column_map = {}  # e.g., {'seq': 1, 'date': 2, 'time': 3, 'store_name': 4, ...}
    max_scan_row = min(20, ws.max_row)

    for r in range(1, max_scan_row + 1):
        row_vals = {}
        for c in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=r, column=c).value
            if cell_val is not None:
                row_vals[c] = clean_str(cell_val)
        
        # 헤더 키워드 매칭 개수 확인
        matched_in_this_row = {}
        for c, text in row_vals.items():
            for field, keywords in HEADER_KEYWORDS.items():
                if field not in matched_in_this_row:
                    if any(kw in text for kw in keywords):
                        matched_in_this_row[field] = c
        
        # 최소 3개 이상 필드가 매칭되면 헤더 행으로 판정
        if len(matched_in_this_row) >= 3:
            header_row_idx = r
            column_map = matched_in_this_row
            break

    # 만약 헤더를 탐지하지 못했을 경우의 기본 폴백 (기존 표준 서식: 5행 헤더, 6행 데이터)
    if not header_row_idx:
        header_row_idx = 5
        column_map = {
            "seq": 1, "date": 2, "time": 3, "store_name": 4,
            "purpose": 5, "attendees": 6, "card_last4": 7,
            "amount": 8, "user_name": 9
        }
    else:
        # B열이 '사용일시'이고 C열이 비어있는 병합 서식 등의 경우: 시간 컬럼이 없으면 날짜 컬럼 옆을 시간으로 보완
        if "date" in column_map and "time" not in column_map:
            date_col = column_map["date"]
            # date_col 다음 열이 비어있거나 매핑되지 않은 경우 시간 열로 할당
            if date_col + 1 not in column_map.values():
                column_map["time"] = date_col + 1

    start_data_row = header_row_idx + 1

    # 소계/합계 행 위치 탐지
    sum_row_idx = None
    for r in range(start_data_row, min(start_data_row + 50, ws.max_row + 1)):
        for c in range(1, 5):
            val = clean_str(ws.cell(row=r, column=c).value)
            if "소계" in val or "합계" in val or "total" in val:
                sum_row_idx = r
                break
        if sum_row_idx:
            break

    return header_row_idx, start_data_row, column_map, sum_row_idx

def fill_receipts_to_excel(template_source=None, receipts: list = None, output_path: str = None, **kwargs) -> bytes:
    """
    법인카드 사용내역서 엑셀 템플릿에 영수증 데이터를 유연하게 매핑하여 주입합니다.
    - template_source 또는 template_path 키워드 인자 모두 완벽 호환
    - 양식이 조금씩 달라도 헤더 키워드를 탐지하여 정확한 열에 찾아 넣습니다.
    - 서식, 테두리, 수식(=SUM) 보존
    """
    # 키워드 인자 호환성 처리
    source = template_source if template_source is not None else kwargs.get("template_path")
    if receipts is None and "receipts" in kwargs:
        receipts = kwargs["receipts"]
    if receipts is None:
        receipts = []

    if isinstance(source, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(source))
    elif isinstance(source, io.BytesIO):
        wb = openpyxl.load_workbook(source)
    elif isinstance(source, str):
        if not os.path.exists(source):
            raise FileNotFoundError(f"엑셀 템플릿 파일을 찾을 수 없습니다: {source}")
        wb = openpyxl.load_workbook(source)
    elif source is not None:
        wb = openpyxl.load_workbook(source)
    else:
        # fallback 기본 템플릿
        default_tpl = "법인카드 사용내역서.xlsx"
        if os.path.exists(default_tpl):
            wb = openpyxl.load_workbook(default_tpl)
        else:
            raise ValueError("엑셀 템플릿이 지정되지 않았습니다.")

    ws = find_target_sheet(wb)

    # 스마트 헤더 및 컬럼 매핑 탐색
    header_row, start_row, col_map, sum_row = detect_column_mapping_and_start_row(ws)

    # 데이터 입력 가능 행 수 계산
    if sum_row and sum_row > start_row:
        max_available_rows = sum_row - start_row
    else:
        max_available_rows = 11

    # 기준 행 스타일 복사
    sample_row = start_row
    cell_styles = {}
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=sample_row, column=c)
        cell_styles[c] = {
            "font": copy(cell.font),
            "border": copy(cell.border),
            "fill": copy(cell.fill),
            "alignment": copy(cell.alignment),
            "number_format": cell.number_format
        }

    num_receipts = len(receipts)
    total_slots = max(max_available_rows, num_receipts)

    for i in range(total_slots):
        current_row = start_row + i

        # 만약 합계 행 직전 슬롯을 초과하면 행 삽입
        if sum_row and current_row >= sum_row:
            ws.insert_rows(current_row)
            sum_row += 1

        if i < num_receipts:
            item = receipts[i]
            
            # 필드 값 준비
            values_by_field = {
                "seq": item.get("seq", i + 1),
                "date": format_date_value(item.get("date", "")),
                "time": format_time_value(item.get("time", "")),
                "store_name": str(item.get("store_name", "")).strip(),
                "purpose": str(item.get("purpose", "")).strip(),
                "attendees": str(item.get("attendees", "")).strip(),
                "user_name": str(item.get("user_name", "")).strip() or None
            }
            
            # 카드번호 끝자리
            card_raw = str(item.get("card_last4", "")).strip()
            try:
                values_by_field["card_last4"] = int(card_raw)
            except ValueError:
                values_by_field["card_last4"] = card_raw or None

            # 금액
            try:
                values_by_field["amount"] = int(str(item.get("amount", 0)).replace(",", ""))
            except ValueError:
                values_by_field["amount"] = 0

            # 탐지된 열에 맞게 값 채우기
            for field, col_idx in col_map.items():
                val = values_by_field.get(field)
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = val

                # 스타일 적용
                if col_idx in cell_styles:
                    st = cell_styles[col_idx]
                    cell.font = copy(st["font"])
                    cell.border = copy(st["border"])
                    cell.fill = copy(st["fill"])
                    cell.alignment = copy(st["alignment"])
                    if st["number_format"]:
                        cell.number_format = st["number_format"]
        else:
            # 남은 빈 슬롯 정리 (연번만 남기고 비움)
            for field, col_idx in col_map.items():
                cell = ws.cell(row=current_row, column=col_idx)
                if field == "seq":
                    cell.value = i + 1
                else:
                    cell.value = None
                if col_idx in cell_styles:
                    st = cell_styles[col_idx]
                    cell.font = copy(st["font"])
                    cell.border = copy(st["border"])
                    cell.fill = copy(st["fill"])
                    cell.alignment = copy(st["alignment"])
                    if st["number_format"]:
                        cell.number_format = st["number_format"]

    # 만약 데이터가 늘어나서 행이 삽입된 경우 소계 수식 업데이트
    if sum_row and "amount" in col_map:
        amt_col_letter = openpyxl.utils.get_column_letter(col_map["amount"])
        sum_cell = ws.cell(row=sum_row, column=col_map["amount"])
        if sum_cell.value and str(sum_cell.value).startswith("="):
            sum_cell.value = f"=SUM({amt_col_letter}{start_row}:{amt_col_letter}{sum_row - 1})"

    # 증빙1 시트 처리 (있을 경우만 안전하게 처리)
    if "증빙1" in wb.sheetnames and num_receipts > 0:
        try:
            ws_ev = wb["증빙1"]
            first_item = receipts[0]
            card = first_item.get("card_last4", "9716")
            user = first_item.get("user_name", "").strip()
            dept = first_item.get("department", "").strip()
            total_sum = sum(int(str(r.get("amount", 0)).replace(",", "")) for r in receipts)

            if ws_ev["D3"].value is not None:
                try:
                    ws_ev["D3"] = int(card)
                except Exception:
                    ws_ev["D3"] = card
            if user and dept:
                ws_ev["G3"] = f"{dept}\n({user})"
            elif user:
                ws_ev["G3"] = user
            else:
                ws_ev["G3"] = ""
            ws_ev["I3"] = total_sum
        except Exception:
            pass

    output_stream = io.BytesIO()
    wb.save(output_stream)
    excel_bytes = output_stream.getvalue()

    if output_path:
        with open(output_path, "wb") as f:
            f.write(excel_bytes)

    return excel_bytes


if __name__ == "__main__":
    from sample_data import SAMPLE_RECEIPTS
    wb_test = openpyxl.load_workbook("법인카드 사용내역서.xlsx")
    ws_test = find_target_sheet(wb_test)
    h_row, s_row, c_map, sum_r = detect_column_mapping_and_start_row(ws_test)
    print("스마트 헤더 탐지 결과:")
    print(f"헤더 행: {h_row}, 데이터 시작 행: {s_row}, 합계 행: {sum_r}")
    print("컬럼 매핑 결과:", c_map)
    out = fill_receipts_to_excel("법인카드 사용내역서.xlsx", SAMPLE_RECEIPTS, "스마트매핑_테스트결과.xlsx")
    print(f"생성 완료: {len(out)} bytes")
