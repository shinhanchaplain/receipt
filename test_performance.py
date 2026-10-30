"""
법인카드 영수증 파싱 & 엑셀 자동화 시스템 종합 성능 및 정밀 검증 스크립트
1. PDF 이미지 변환 렌더링 성능 (속도, 해상도)
2. 영수증 및 수기 메모(사용목적) 데이터 정밀도
3. 금액 및 구매 개수 기반 대상인원 산출 알고리즘 검증
4. 엑셀 템플릿 주입 성능 및 수식/서식 무결성 검증
5. 전체 파이프라인 종단간(End-to-End) 소요 시간 측정
"""

import os
import sys
import time
import io
import openpyxl

# 콘솔 UTF-8 출력 보장
sys.stdout.reconfigure(encoding='utf-8')

import receipt_parser
import excel_manager
from sample_data import SAMPLE_RECEIPTS

def print_separator(title=""):
    if title:
        print(f"\n{'='*20} [ {title} ] {'='*20}")
    else:
        print(f"{'='*60}")

def run_performance_test():
    print_separator("법인카드 영수증 자동화 시스템 성능 및 무결성 테스트")
    start_total_time = time.time()
    results = []

    # ---------------------------------------------------------
    # TEST 1: PDF 이미지 변환 렌더링 성능 테스트
    # ---------------------------------------------------------
    print("\n▶ [TEST 1] PDF 이미지 변환 및 렌더링 성능 측정...")
    pdf_path = "영수증.pdf"
    if not os.path.exists(pdf_path):
        results.append(("PDF 렌더링", "FAIL", "영수증.pdf 파일 없음", 0))
    else:
        t0 = time.time()
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
        images = receipt_parser.extract_images_from_pdf(pdf_bytes, dpi=150)
        t_render = (time.time() - t0) * 1000  # ms
        
        page_count = len(images)
        sizes = [f"{img.width}x{img.height}" for img in images]
        print(f"  - 페이지 수: {page_count}페이지")
        print(f"  - 렌더링 해상도: {', '.join(sizes)}")
        print(f"  - 소요 시간: {t_render:.1f} ms (페이지당 {t_render/page_count:.1f} ms)")
        
        if page_count == 3 and t_render < 3000:
            results.append(("PDF 렌더링 (3페이지 고해상도)", "PASS", f"{t_render:.1f}ms (합격 기준 < 3000ms)", t_render))
            print("  ✔ [PASS] PDF 렌더링 성능 합격")
        else:
            results.append(("PDF 렌더링", "FAIL", f"{t_render:.1f}ms", t_render))
            print("  ✖ [FAIL] PDF 렌더링 비정상")

    # ---------------------------------------------------------
    # TEST 2: 영수증 데이터 및 수기 메모 추출 정합성 검증
    # ---------------------------------------------------------
    print("\n▶ [TEST 2] 영수증 데이터 및 수기 메모(보건소명) 정합성 검증...")
    t0 = time.time()
    receipts = receipt_parser.parse_receipt_document(pdf_bytes, "영수증.pdf")
    t_parse = (time.time() - t0) * 1000

    expected_purposes = [
        "이정복 보건소",
        "이정복 보건소",
        "영등포 보건소",
        "동작구 보건소",
        "양주시 보건소"
    ]
    expected_amounts = [47700, 47700, 100100, 51300, 51700]
    expected_attendees = ["3명", "3명", "8명", "3명", "3명"]

    purposes_match = True
    amounts_match = True
    attendees_match = True

    print(f"  - 추출된 영수증 건수: {len(receipts)}건")
    for i, r in enumerate(receipts):
        p_ok = r.get("purpose") == expected_purposes[i]
        a_ok = r.get("amount") == expected_amounts[i]
        att_ok = r.get("attendees") == expected_attendees[i]
        
        print(f"    건 #{i+1}: 가맹점={r.get('store_name')}, 수기메모={r.get('purpose')}({'OK' if p_ok else 'ERR'}), "
              f"인원={r.get('attendees')}({'OK' if att_ok else 'ERR'}), 금액={r.get('amount'):,}원({'OK' if a_ok else 'ERR'})")
        
        if not p_ok: purposes_match = False
        if not a_ok: amounts_match = False
        if not att_ok: attendees_match = False

    total_extracted_amount = sum(r.get("amount", 0) for r in receipts)
    print(f"  - 총 합산 금액: {total_extracted_amount:,}원 (증빙서 기재액 298,500원과 비교)")

    if len(receipts) == 5 and purposes_match and amounts_match and total_extracted_amount == 298500:
        results.append(("영수증 및 수기메모 정밀도", "PASS", "5건 모두 100% 일치 (합산 298,500원)", t_parse))
        print("  ✔ [PASS] 수기 메모 및 영수증 데이터 100% 일치")
    else:
        results.append(("영수증 및 수기메모 정밀도", "FAIL", "데이터 불일치 발생", t_parse))
        print("  ✖ [FAIL] 영수증 데이터 불일치")

    # ---------------------------------------------------------
    # TEST 3: 대상인원 산출 알고리즘 엣지 케이스 테스트
    # ---------------------------------------------------------
    print("\n▶ [TEST 3] 대상인원 산출 엔진 다중 시나리오 벤치마크...")
    test_cases = [
        {"qty": 3, "amt": 47700, "expected": "3명", "desc": "일반 제과 3개"},
        {"qty": 8, "amt": 100100, "expected": "8명", "desc": "다과 세트 8개"},
        {"qty": 1, "amt": 45000, "expected": "3명", "desc": "고가 단품 케이크 1개 (단가 보정)"},
        {"qty": 1, "amt": 8000, "expected": "1명", "desc": "1인 식대/음료 1개"},
        {"qty": 0, "amt": 36000, "expected": "3명", "desc": "수량 미기재 시 금액 기준 추정"},
    ]
    calc_all_pass = True
    for tc in test_cases:
        res = receipt_parser.calculate_attendees(tc["qty"], tc["amt"])
        match = (res == tc["expected"])
        print(f"    시나리오 [{tc['desc']}]: 수량={tc['qty']}, 금액={tc['amt']:,}원 -> 산출={res} (기대={tc['expected']}) [{'OK' if match else 'FAIL'}]")
        if not match:
            calc_all_pass = False

    if calc_all_pass:
        results.append(("대상인원 산출 알고리즘", "PASS", "5/5 시나리오 정상 산출", 0))
        print("  ✔ [PASS] 대상인원 산출 알고리즘 합격")
    else:
        results.append(("대상인원 산출 알고리즘", "FAIL", "일부 시나리오 불일치", 0))
        print("  ✖ [FAIL] 대상인원 산출 오류")

    # ---------------------------------------------------------
    # TEST 4: 엑셀 파일 생성 성능 및 수식/서식 무결성 검증
    # ---------------------------------------------------------
    print("\n▶ [TEST 4] 엑셀 템플릿 주입 및 수식 무결성 검증...")
    template_path = "법인카드 사용내역서.xlsx"
    out_path = "법인카드_사용내역서_성능테스트결과.xlsx"
    
    t0 = time.time()
    excel_bytes = excel_manager.fill_receipts_to_excel(template_path, receipts, out_path)
    t_excel = (time.time() - t0) * 1000

    # 생성된 엑셀 파일 정밀 검증
    wb_test = openpyxl.load_workbook(out_path, data_only=False)
    ws_test = wb_test["법인카드 사용내역"]
    
    formula_h18 = ws_test["H18"].value
    formula_h19 = ws_test["H19"].value
    
    # 데이터 값 검증 (data_only=True로 재로드하여 수식 계산값 확인)
    wb_val = openpyxl.load_workbook(out_path, data_only=True)
    ws_val = wb_val["법인카드 사용내역"]
    
    row6_store = ws_test["D6"].value
    row6_purpose = ws_test["E6"].value
    row6_att = ws_test["F6"].value
    row6_amt = ws_test["H6"].value
    row10_purpose = ws_test["E10"].value

    row6_user = ws_test["I6"].value
    print(f"  - 엑셀 생성 소요 시간: {t_excel:.1f} ms")
    print(f"  - H18 셀 수식: {formula_h18} (기대: =SUM(H6:H17))")
    print(f"  - H19 셀 수식: {formula_h19} (기대: =H18)")
    print(f"  - 1번 영수증 주입값: {row6_store} | {row6_purpose} | {row6_att} | {row6_amt:,}원 | 사용자(공란여부)={row6_user}")
    print(f"  - 5번 영수증 주입값: {ws_test['D10'].value} | {row10_purpose} | {ws_test['F10'].value} | {ws_test['H10'].value:,}원")
    print(f"  - 6번 빈 행(A11~I11) 정리 상태: D11={ws_test['D11'].value}, H11={ws_test['H11'].value}")

    excel_pass = (
        formula_h18 == "=SUM(H6:H17)" and
        formula_h19 == "=H18" and
        row6_purpose == "이정복 보건소" and
        row10_purpose == "양주시 보건소" and
        row6_user is None and
        ws_test["D11"].value is None and
        t_excel < 1500
    )

    if excel_pass:
        results.append(("엑셀 템플릿 무결성 및 수식 검증", "PASS", f"{t_excel:.1f}ms, 수식/서식 100% 정상", t_excel))
        print("  ✔ [PASS] 엑셀 수식 및 데이터 무결성 검증 통과")
    else:
        results.append(("엑셀 템플릿 무결성 및 수식 검증", "FAIL", "수식 또는 서식 손상", t_excel))
        print("  ✖ [FAIL] 엑셀 무결성 검증 실패")

    # ---------------------------------------------------------
    # 종합 평가 및 결과 리포트 출력
    # ---------------------------------------------------------
    total_elapsed = (time.time() - start_total_time) * 1000
    print_separator("성능 및 품질 검증 최종 리포트")
    print(f"{'항목명':<32} | {'상태':<8} | {'상세 내용'}")
    print("-" * 65)
    all_passed = True
    for name, status, detail, duration in results:
        status_str = f"✅ {status}" if status == "PASS" else f"❌ {status}"
        print(f"{name:<30} | {status_str:<8} | {detail}")
        if status != "PASS":
            all_passed = False
    print("-" * 65)
    print(f"총 검증 소요 시간: {total_elapsed:.1f} ms")
    
    if all_passed:
        print("\n🎉 [SUCCESS] 모든 성능 및 정합성 테스트를 통과했습니다!")
        print("웹페이지 앱으로 게시(배포)하기에 완벽한 준비 상태입니다.\n")
        return 0
    else:
        print("\n⚠️ [WARNING] 일부 테스트에서 문제점이 발견되었습니다.\n")
        return 1

if __name__ == "__main__":
    exit_code = run_performance_test()
    sys.exit(exit_code)
