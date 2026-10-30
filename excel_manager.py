import os
import io
import datetime
from copy import copy
import openpyxl
from openpyxl.styles import Alignment, Font, Border, PatternFill

# 유연한 헤더 매핑 키워드 사전 (모든 특수문자/공백/줄바꿈 제거 후 소문자로 비교)
HEADER_KEYWORDS = {
    "seq": [
        "연번", "순번", "no", "번호", "순서", "일련번호"
    ],
    "date": [
        "집행일자", "집행일시", "사용일자", "사용일시", "지출일자", "지출일시",
        "일자", "날짜", "승인일자", "승인일시", "결제일자"
    ],
    "time": [
        "집행시간", "집행시각", "사용시간", "사용시각", "지출시간", "시간", "시각", "승인시간"
    ],
    "store_name": [
        "집행장소", "사용장소", "가맹점", "가맹점명", "상호", "상호명",
        "장소", "구입처", "거래처", "물품명", "구입점", "지출처"
    ],
    "purpose": [
        "집행목적", "사용목적", "지출목적", "목적", "집행내역", "사용내역", "지출내용", "집행내용",
        "적요", "구체적명기", "구체적", "세부내용", "사업명", "품명", "내용", "비고"
    ],
    "attendees": [
        "집행대상", "대상인원", "참석자", "인원", "수량", "인원수",
        "인원(수량)", "참석인원", "대상자", "집행인원"
    ],
    "card_last4": [
        "카드번호", "카드", "카드끝", "끝4자리", "카드번호(끝4자리)", "카드끝자리"
    ],
    "amount": [
        "집행금액", "지출금액", "금액", "결제금액", "금 액", "지출액", "사용금액", "합계금액", "승인금액"
    ],
    "user_name": [
        "집행자", "사용자", "담당자", "소지자", "사용자명", "성명", "작성자", "카드사용자"
    ]
}

def clean_str(val):
    """문자열에서 괄호, 공백, 언더바, 줄바꿈 등을 모두 제거하여 정규화"""
    if val is None:
        return ""
    s = str(val).lower()
    for ch in [" ", "\n", "\r", "\t", "_", "(", ")", "[", "]", "{", "}", "-", ":"]:
        s = s.replace(ch, "")
    return s

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
    """
    문자열 날짜를 datetime.date 객체로 변환하여 엑셀에서 '################' 오류가 나지 않도록 보장
    """
    if not date_str:
        return None
    if isinstance(date_str, datetime.date):
        return date_str
    if isinstance(date_str, datetime.datetime):
        return date_str.date()

    clean = str(date_str).replace("-", "").replace(".", "").replace("/", "").strip()
    if len(clean) == 8 and clean.isdigit():
        try:
            y = int(clean[:4])
            m = int(clean[4:6])
            d = int(clean[6:8])
            return datetime.date(y, m, d)
        except Exception:
            pass
    return str(date_str).strip() or None

def find_target_sheet(wb):
    """업무추진비 / 법인카드 / 지출 관련 시트 자동 탐색 (없으면 첫 번째 활성 시트 반환)"""
    candidate_keywords = ["업무추진비", "법인카드", "사용내역", "지출", "경비", "카드", "내역", "sheet1"]
    for sheetname in wb.sheetnames:
        clean_name = clean_str(sheetname)
        if any(kw in clean_name for kw in candidate_keywords):
            return wb[sheetname]
    return wb.active

def detect_column_mapping_and_start_row(ws):
    """
    엑셀 시트를 스캔하여 헤더 행과 각 컬럼(열)의 위치를 유연하게 탐지합니다.
    - 양식이 조금씩 달라도 키워드를 찾아 매핑합니다.
    """
    header_row_idx = None
    column_map = {}
    max_scan_row = min(25, ws.max_row)

    for r in range(1, max_scan_row + 1):
        row_vals = {}
        for c in range(1, ws.max_column + 1):
            cell_val = ws.cell(row=r, column=c).value
            if cell_val is not None:
                row_vals[c] = clean_str(cell_val)
        
        matched_in_this_row = {}
        for c, text in row_vals.items():
            for field, keywords in HEADER_KEYWORDS.items():
                if field not in matched_in_this_row:
                    if any(clean_str(kw) in text for kw in keywords):
                        matched_in_this_row[field] = c
        
        # 최소 2개 이상 필드가 매칭되면 헤더 행으로 판정
        if len(matched_in_this_row) >= 2:
            header_row_idx = r
            column_map = matched_in_this_row
            break

    # 만약 헤더를 탐지하지 못했을 경우의 기본 폴백
    if not header_row_idx:
        header_row_idx = 5
        column_map = {
            "seq": 1, "date": 2, "time": 3, "store_name": 4,
            "purpose": 5, "attendees": 6, "card_last4": 7,
            "amount": 8, "user_name": 9
        }
    else:
        # 일시 컬럼 옆에 빈 칸이 있으면 시간 열로 할당
        if "date" in column_map and "time" not in column_map:
            date_col = column_map["date"]
            if date_col + 1 not in column_map.values() and date_col + 1 <= ws.max_column:
                column_map["time"] = date_col + 1

    start_data_row = header_row_idx + 1

    # 소계/합계 행 위치 탐색
    sum_row_idx = None
    for r in range(start_data_row, min(start_data_row + 60, ws.max_row + 1)):
        for c in range(1, 6):
            val = clean_str(ws.cell(row=r, column=c).value)
            if any(k in val for k in ["소계", "합계", "총계", "total", "sum"]):
                sum_row_idx = r
                break
        if sum_row_idx:
            break

    return header_row_idx, start_data_row, column_map, sum_row_idx

def fill_receipts_to_excel(template_source=None, receipts: list = None, output_path: str = None, **kwargs) -> bytes:
    """
    엑셀 템플릿에 영수증 데이터를 안전하게 주입합니다.
    - 확인되지 않거나 구분할 수 없는 필드는 100% 공란(None)으로 유지
    - 기존 서식, 글꼴, 셀 테두리 및 수식(=SUM) 보존
    """
    source = template_source if template_source is not None else kwargs.get("template_path")
    if receipts is None and "receipts" in kwargs:
        receipts = kwargs["receipts"]
    if receipts is None:
        receipts = []

    if isinstance(source, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(source))
    elif isinstance(source, io.BytesIO):
        wb = openpyxl.load_workbook(source)
    elif isinstance(source, openpyxl.Workbook):
        wb = source
    elif isinstance(source, str):
        if not os.path.exists(source):
            raise FileNotFoundError(f"엑셀 템플릿 파일을 찾을 수 없습니다: {source}")
        wb = openpyxl.load_workbook(source)
    elif source is not None:
        try:
            wb = openpyxl.load_workbook(source)
        except Exception:
            wb = source
    else:
        default_tpl = "법인카드 사용내역서.xlsx"
        if os.path.exists(default_tpl):
            wb = openpyxl.load_workbook(default_tpl)
        else:
            raise ValueError("엑셀 템플릿이 지정되지 않았습니다.")

    ws = find_target_sheet(wb)

    # 스마트 헤더 탐지
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

        # 합계 행을 침범하지 않도록 필요 시 동적 행 삽입
        if sum_row and current_row >= sum_row:
            ws.insert_rows(current_row)
            sum_row += 1

        if i < num_receipts:
            item = receipts[i]
            
            # 필드 값 준비 (확인되지 않는 값은 철저히 None 공란 처리)
            p_val = str(item.get("purpose", "")).strip()
            u_val = str(item.get("user_name", "")).strip()
            att_val = str(item.get("attendees", "")).strip()
            store_val = str(item.get("store_name", "")).strip()
            
            values_by_field = {
                "seq": item.get("seq", i + 1),
                "date": format_date_value(item.get("date", "")) or None,
                "time": format_time_value(item.get("time", "")) or None,
                "store_name": store_val or None,
                "purpose": p_val or None,       # 확인 안 되면 공란
                "attendees": att_val or None,   # 확인 안 되면 공란
                "user_name": u_val or None      # 확인 안 되면 공란
            }
            
            # 카드번호 끝자리
            card_raw = str(item.get("card_last4", "")).strip()
            if card_raw:
                try:
                    values_by_field["card_last4"] = int(card_raw)
                except ValueError:
                    values_by_field["card_last4"] = card_raw
            else:
                values_by_field["card_last4"] = None

            # 금액
            amt_raw = str(item.get("amount", "")).replace(",", "").strip()
            if amt_raw:
                try:
                    values_by_field["amount"] = int(amt_raw)
                except ValueError:
                    values_by_field["amount"] = None
            else:
                values_by_field["amount"] = None

            # 매핑된 열에만 안전하게 값 삽입 (알 수 없는 열은 원본 보존)
            for field, col_idx in col_map.items():
                val = values_by_field.get(field)
                cell = ws.cell(row=current_row, column=col_idx)
                cell.value = val

                if col_idx in cell_styles:
                    st = cell_styles[col_idx]
                    cell.font = copy(st["font"])
                    cell.border = copy(st["border"])
                    cell.fill = copy(st["fill"])
                    cell.alignment = copy(st["alignment"])
                    if st["number_format"]:
                        cell.number_format = st["number_format"]

                # 날짜 필드인데 서식이 General이거나 없으면 yyyy-mm-dd로 보정
                if field == "date" and isinstance(val, (datetime.date, datetime.datetime)):
                    if not cell.number_format or str(cell.number_format).lower() == "general":
                        cell.number_format = "yyyy-mm-dd"
        else:
            # 남은 빈 행 정리
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

                # 날짜 필드인데 서식이 General이거나 없으면 yyyy-mm-dd로 보정
                if field == "date" and isinstance(val, (datetime.date, datetime.datetime)):
                    if not cell.number_format or str(cell.number_format).lower() == "general":
                        cell.number_format = "yyyy-mm-dd"

    # 소계/합계 수식 업데이트 (행이 늘어난 경우)
    if sum_row and "amount" in col_map:
        amt_col_letter = openpyxl.utils.get_column_letter(col_map["amount"])
        sum_cell = ws.cell(row=sum_row, column=col_map["amount"])
        if sum_cell.value and str(sum_cell.value).startswith("="):
            sum_cell.value = f"=SUM({amt_col_letter}{start_row}:{amt_col_letter}{sum_row - 1})"

    # 날짜/시간 열 너비 자동 보정 (너비 부족으로 인한 '################' 방지)
    try:
        if "date" in col_map:
            d_letter = openpyxl.utils.get_column_letter(col_map["date"])
            curr_w = ws.column_dimensions[d_letter].width
            ws.column_dimensions[d_letter].width = max(float(curr_w or 8.5), 14.0)
        if "time" in col_map:
            t_letter = openpyxl.utils.get_column_letter(col_map["time"])
            curr_tw = ws.column_dimensions[t_letter].width
            ws.column_dimensions[t_letter].width = max(float(curr_tw or 8.5), 10.0)
    except Exception:
        pass

    # 증빙1 시트 부가 처리 (안전하게 시도하되 실패해도 메인은 정상 출력)
    if "증빙1" in wb.sheetnames and num_receipts > 0:
        try:
            ws_ev = wb["증빙1"]
            first_item = receipts[0]
            card = first_item.get("card_last4", "")
            user = first_item.get("user_name", "").strip()
            total_sum = sum(int(str(r.get("amount", 0) or 0).replace(",", "")) for r in receipts)

            if ws_ev["D3"].value is not None and card:
                try:
                    ws_ev["D3"] = int(card)
                except Exception:
                    ws_ev["D3"] = card
            if user:
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
