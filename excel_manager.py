import os
import io
import datetime
from copy import copy
import openpyxl
from openpyxl.styles import Alignment, Font, Border, PatternFill

def format_time_value(time_str):
    """문자열 시간(HH:MM:SS 또는 HH:MM)을 datetime.time 객체 또는 문자열로 변환"""
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
    """날짜를 20261001 형태의 정수 또는 문자열로 변환"""
    if not date_str:
        return ""
    clean = str(date_str).replace("-", "").replace(".", "").replace("/", "").strip()
    try:
        return int(clean)
    except Exception:
        return clean

def fill_receipts_to_excel(template_path: str, receipts: list, output_path: str = None) -> bytes:
    """
    법인카드 사용내역서 엑셀 템플릿에 영수증 데이터를 주입하고 엑셀 바이트를 반환합니다.
    - 기존 서식, 글꼴, 셀 테두리, 수식(=SUM(H6:H17))을 100% 보존합니다.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"엑셀 템플릿 파일을 찾을 수 없습니다: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    
    if "법인카드 사용내역" not in wb.sheetnames:
        raise ValueError("엑셀 파일에 '법인카드 사용내역' 시트가 존재하지 않습니다.")
    
    ws = wb["법인카드 사용내역"]

    # 기본 행 범위 (연번 1~11, Row 6 ~ Row 16)
    start_row = 6
    max_data_rows = 11  # 기본 양식에 연번 1~11까지 11개 행 존재
    
    # 첫 행(Row 6)의 스타일을 템플릿 스타일로 참조
    cell_styles = {}
    for col_idx in range(1, 10):
        c = ws.cell(row=start_row, column=col_idx)
        cell_styles[col_idx] = {
            "font": copy(c.font),
            "border": copy(c.border),
            "fill": copy(c.fill),
            "alignment": copy(c.alignment),
            "number_format": c.number_format
        }

    # 영수증 데이터 입력
    num_receipts = len(receipts)
    
    # 만약 영수증이 11건을 초과하면 행을 동적으로 삽입하거나 확장
    # 일반적인 영수증 11건 이하 처리
    total_slots = max(max_data_rows, num_receipts)
    
    for i in range(total_slots):
        current_row = start_row + i
        
        # 11건을 초과하는 경우 새 행 생성 및 수식 조정
        if i >= max_data_rows:
            ws.insert_rows(current_row)
        
        if i < num_receipts:
            item = receipts[i]
            
            seq = item.get("seq", i + 1)
            date_val = format_date_value(item.get("date", ""))
            time_val = format_time_value(item.get("time", ""))
            store_name = str(item.get("store_name", "")).strip()
            purpose = str(item.get("purpose", "")).strip()
            attendees = str(item.get("attendees", "")).strip()
            card_last4 = str(item.get("card_last4", "")).strip()
            try:
                card_last4_val = int(card_last4)
            except ValueError:
                card_last4_val = card_last4
                
            try:
                amount_val = int(str(item.get("amount", 0)).replace(",", ""))
            except ValueError:
                amount_val = 0
                
            user_name_raw = str(item.get("user_name", "")).strip()
            user_name = user_name_raw if user_name_raw and user_name_raw != "None" else None
            
            row_data = [
                seq,            # A: 연번
                date_val,       # B: 사용일시 (일자)
                time_val,       # C: 사용일시 (시간)
                store_name,     # D: 사용장소(물품명)
                purpose,        # E: 사용목적(구체적 명기) - 수기 내용
                attendees,      # F: 대상인원(수량)
                card_last4_val, # G: 카드번호(끝4자리)
                amount_val,     # H: 금 액(원)
                user_name       # I: 사용자 (확인 안 될 경우 공란)
            ]
        else:
            # 남은 행은 연번만 남기고 내용 비우기
            row_data = [i + 1, None, None, None, None, None, None, None, None]

        # 셀에 값 및 스타일 적용
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.value = val
            
            # 스타일 복제 적용
            if col_idx in cell_styles:
                st = cell_styles[col_idx]
                cell.font = copy(st["font"])
                cell.border = copy(st["border"])
                cell.fill = copy(st["fill"])
                cell.alignment = copy(st["alignment"])
                if st["number_format"]:
                    cell.number_format = st["number_format"]

    # 만약 데이터가 11건을 초과했다면 소계 수식 업데이트
    if num_receipts > max_data_rows:
        sum_row = start_row + num_receipts + 1
        ws[f"H{sum_row}"] = f"=SUM(H{start_row}:H{sum_row - 1})"
        ws[f"H{sum_row + 1}"] = f"=H{sum_row}"

    # 증빙1 시트가 있을 경우 상단 정보 업데이트
    if "증빙1" in wb.sheetnames and num_receipts > 0:
        try:
            ws_ev = wb["증빙1"]
            first_item = receipts[0]
            card = first_item.get("card_last4", "9716")
            user = first_item.get("user_name", "").strip()
            dept = first_item.get("department", "").strip()
            total_sum = sum(int(str(r.get("amount", 0)).replace(",", "")) for r in receipts)
            
            # D3 (카드번호 끝자리)
            try:
                ws_ev["D3"] = int(card)
            except Exception:
                ws_ev["D3"] = card
                
            # G3 (소속부서 및 사용자명 - 미확인 시 공란)
            if user and dept:
                ws_ev["G3"] = f"{dept}\n({user})"
            elif user:
                ws_ev["G3"] = user
            else:
                ws_ev["G3"] = ""
                
            # I3 (합산금액)
            ws_ev["I3"] = total_sum
        except Exception as e:
            # 병합 셀 등으로 인한 부가 시트 오류는 본문 시트 데이터 주입에 영향을 주지 않도록 방어
            pass

    # 결과 저장
    output_stream = io.BytesIO()
    wb.save(output_stream)
    excel_bytes = output_stream.getvalue()

    if output_path:
        with open(output_path, "wb") as f:
            f.write(excel_bytes)

    return excel_bytes


if __name__ == "__main__":
    from sample_data import SAMPLE_RECEIPTS
    
    template = "법인카드 사용내역서.xlsx"
    out_test = "법인카드 사용내역서_테스트결과.xlsx"
    print("Testing Excel Manager...")
    fill_receipts_to_excel(template, SAMPLE_RECEIPTS, out_test)
    print(f"Excel file created successfully: {out_test}")
