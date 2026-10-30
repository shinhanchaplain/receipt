import os
import io
import datetime
import pandas as pd
import streamlit as st
from PIL import Image

import receipt_parser
import excel_manager
from sample_data import SAMPLE_RECEIPTS

# 1. 페이지 설정
st.set_page_config(
    page_title="법인카드 영수증 파싱 & 엑셀 자동화",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# 커스텀 CSS 스타일링
st.markdown("""
<style>
    .main-title {
        font-size: 2.1rem;
        font-weight: 800;
        color: #1E293B;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .step-box {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
    }
    .step-header {
        font-size: 1.25rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-value {
        font-size: 1.7rem;
        font-weight: 800;
        color: #0F172A;
    }
    .metric-label {
        font-size: 0.88rem;
        color: #64748B;
        margin-bottom: 4px;
    }
</style>
""", unsafe_allow_html=True)

DEFAULT_EXCEL_TEMPLATE = "법인카드 사용내역서.xlsx"
DEFAULT_RECEIPT_PDF = "영수증.pdf"

# 세션 상태 초기화
if "receipt_data" not in st.session_state:
    st.session_state.receipt_data = [dict(r) for r in SAMPLE_RECEIPTS]

if "pdf_images" not in st.session_state:
    st.session_state.pdf_images = []
    if os.path.exists(DEFAULT_RECEIPT_PDF):
        with open(DEFAULT_RECEIPT_PDF, "rb") as f:
            st.session_state.pdf_images = receipt_parser.extract_images_from_pdf(f.read())

if "excel_template_bytes" not in st.session_state:
    if os.path.exists(DEFAULT_EXCEL_TEMPLATE):
        with open(DEFAULT_EXCEL_TEMPLATE, "rb") as f:
            st.session_state.excel_template_bytes = f.read()
    else:
        st.session_state.excel_template_bytes = None

# 상단 헤더
st.markdown('<div class="main-title">🧾 법인카드 영수증 자동 정리 & 엑셀 다운로드</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">엑셀 양식 파일과 스캔 영수증을 업로드하고 <b>[실행]</b>을 누르면, '
    'AI가 가맹점/일시/금액과 <b>수기 메모(보건소명/병원명 등 사용목적)</b>를 인식하여 엑셀에 자동으로 정리해줍니다.</div>',
    unsafe_allow_html=True
)

# 사이드바 (API 키 및 기본 설정)
with st.sidebar:
    st.header("⚙️ 환경 설정")
    api_key_input = st.text_input(
        "Gemini API Key (선택)",
        type="password",
        value=os.environ.get("GEMINI_API_KEY", ""),
        help="새로운 영수증 스캔본을 실시간 비전 AI로 분석할 때 사용합니다. 미입력 시 내장 샘플 데이터셋으로 시연됩니다."
    )
    if api_key_input:
        os.environ["GEMINI_API_KEY"] = api_key_input
        st.success("API Key 설정 완료")
    
    st.markdown("---")
    st.caption("※ 사용자 등 영수증에서 확인되지 않는 항목은 기본적으로 **공란**으로 처리됩니다.")

# ==============================================================================
# [STEP 1] 파일 업로드 및 실행 영역
# ==============================================================================
st.markdown("""
<div class="step-header">
    <span>📂 [1단계] 엑셀 양식 및 영수증 스캔 파일 업로드</span>
</div>
""", unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2)

with col_up1:
    uploaded_excel = st.file_uploader(
        "1. 정리해서 넣을 엑셀 파일 (.xlsx)",
        type=["xlsx"],
        help="법인카드 사용내역서 양식 엑셀 파일을 업로드하세요. (미업로드 시 기본 '법인카드 사용내역서.xlsx' 사용)"
    )
    if uploaded_excel:
        st.session_state.excel_template_bytes = uploaded_excel.read()
        st.caption(f"✔ 업로드된 엑셀: `{uploaded_excel.name}` 적용됨")
    elif os.path.exists(DEFAULT_EXCEL_TEMPLATE):
        st.caption(f"ℹ 기본 엑셀 양식 적용 중: `{DEFAULT_EXCEL_TEMPLATE}`")

with col_up2:
    uploaded_receipt = st.file_uploader(
        "2. 스캔한 영수증 파일 (PDF 또는 이미지)",
        type=["pdf", "png", "jpg", "jpeg"],
        help="스캔된 영수증 PDF 또는 이미지 파일을 업로드하세요. (미업로드 시 기본 '영수증.pdf' 사용)"
    )
    if uploaded_receipt:
        st.caption(f"✔ 업로드된 영수증: `{uploaded_receipt.name}` 준비됨")
    elif os.path.exists(DEFAULT_RECEIPT_PDF):
        st.caption(f"ℹ 기본 스캔 파일 적용 중: `{DEFAULT_RECEIPT_PDF}` (3페이지)")

# 🚀 실행 버튼
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
with col_btn2:
    run_btn = st.button("🚀 영수증 분석 & 엑셀 자동 정리 실행", type="primary", use_container_width=True)

if run_btn:
    with st.spinner("영수증 이미지 및 손글씨(수기) 메모를 분석하여 엑셀 내역을 정리 중입니다..."):
        # 1. 영수증 바이트 확보
        if uploaded_receipt:
            receipt_bytes = uploaded_receipt.getvalue()
            receipt_filename = uploaded_receipt.name
        else:
            with open(DEFAULT_RECEIPT_PDF, "rb") as f:
                receipt_bytes = f.read()
            receipt_filename = DEFAULT_RECEIPT_PDF
        
        # 2. 이미지 렌더링
        if receipt_filename.lower().endswith(".pdf"):
            st.session_state.pdf_images = receipt_parser.extract_images_from_pdf(receipt_bytes)
        else:
            st.session_state.pdf_images = [Image.open(io.BytesIO(receipt_bytes))]

        # 3. 파싱 실행 (확인 안 되는 자료는 공란 세팅)
        parsed_results = receipt_parser.parse_receipt_document(
            file_bytes=receipt_bytes,
            file_name=receipt_filename,
            api_key=api_key_input or os.environ.get("GEMINI_API_KEY")
        )
        st.session_state.receipt_data = parsed_results
        st.success(f"🎉 총 {len(parsed_results)}건의 영수증 내역 정리가 완료되었습니다! 아래에서 확인하세요.")

st.markdown("---")

# ==============================================================================
# [STEP 2] 결과 검토 및 실시간 수정 영역
# ==============================================================================
st.markdown("""
<div class="step-header">
    <span>✏️ [2단계] 영수증 실물 확인 및 정리된 내역 검토 / 수정</span>
</div>
""", unsafe_allow_html=True)

data_list = st.session_state.receipt_data
total_count = len(data_list)
total_amount = sum(int(str(r.get("amount", 0)).replace(",", "")) for r in data_list)
avg_amount = total_amount // total_count if total_count > 0 else 0
handwritten_count = sum(1 for r in data_list if r.get("purpose"))

# KPI 카드
col_m1, col_m2, col_m3, col_m4 = st.columns(4)
with col_m1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">총 영수증 건수</div>
        <div class="metric-value">{total_count}건</div>
    </div>
    """, unsafe_allow_html=True)
with col_m2:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">총 지출 금액</div>
        <div class="metric-value">{total_amount:,}원</div>
    </div>
    """, unsafe_allow_html=True)
with col_m3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">건당 평균 금액</div>
        <div class="metric-value">{avg_amount:,}원</div>
    </div>
    """, unsafe_allow_html=True)
with col_m4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">수기 메모(목적) 인식</div>
        <div class="metric-value" style="color: #059669;">{handwritten_count} / {total_count}</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

# 2열 레이아웃: 영수증 실물 뷰어 vs 데이터 에디터
col_view, col_edit = st.columns([1, 1.3])

with col_view:
    st.subheader("👁️ 스캔 영수증 실물 뷰어")
    if st.session_state.pdf_images:
        num_pages = len(st.session_state.pdf_images)
        page_idx = st.slider("페이지 선택", 1, num_pages, 1) if num_pages > 1 else 1
        current_img = st.session_state.pdf_images[page_idx - 1]
        st.image(
            current_img,
            caption=f"스캔본 {page_idx}페이지 (영수증 실물 및 상단 증빙서 첨부지)",
            use_container_width=True
        )
        st.info("💡 영수증 여백에 적힌 수기 메모(예: '이정복 보건소', '영등포 보건소' 등)와 품목/수량을 직접 대조해보실 수 있습니다.")
    else:
        st.warning("표시할 영수증 이미지가 없습니다.")

with col_edit:
    st.subheader("📊 정리된 법인카드 내역서 (직접 수정 가능)")
    st.caption("※ 셀을 더블 클릭하여 수정할 수 있습니다. 사용자 등 확인 불가 항목은 공란으로 비워져 있습니다.")

    display_rows = []
    for r in st.session_state.receipt_data:
        display_rows.append({
            "연번": r.get("seq", 1),
            "사용일자": str(r.get("date", "")),
            "사용시각": str(r.get("time", "")),
            "사용장소": r.get("store_name", ""),
            "사용목적(수기메모)": r.get("purpose", ""),
            "대상인원": r.get("attendees", "3명"),
            "카드끝4자리": str(r.get("card_last4", "9716")),
            "금액(원)": int(str(r.get("amount", 0)).replace(",", "")),
            "사용자": r.get("user_name", "")
        })

    df = pd.DataFrame(display_rows)

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "연번": st.column_config.NumberColumn("연번", width=50),
            "사용일자": st.column_config.TextColumn("사용일자", width=100),
            "사용시각": st.column_config.TextColumn("사용시각", width=85),
            "사용장소": st.column_config.TextColumn("사용장소 (가맹점)", width=130),
            "사용목적(수기메모)": st.column_config.TextColumn("사용목적 (수기메모)", width=130),
            "대상인원": st.column_config.SelectboxColumn(
                "대상인원 (수량기반 자동산출)",
                options=["1명", "2명", "3명", "4명", "5명", "6명", "7명", "8명", "9명", "10명"],
                width=100
            ),
            "카드끝4자리": st.column_config.TextColumn("카드끝4자리", width=85),
            "금액(원)": st.column_config.NumberColumn("금액(원)", format="%d원", width=95),
            "사용자": st.column_config.TextColumn("사용자 (확인불가시 공란)", width=100),
        },
        height=380
    )

    # 수정된 데이터 세션에 반영
    updated_receipts = []
    for idx, row in edited_df.iterrows():
        u_name = str(row["사용자"]).strip() if pd.notna(row["사용자"]) and str(row["사용자"]).strip() != "None" else ""
        updated_receipts.append({
            "seq": int(row["연번"]),
            "date": str(row["사용일자"]),
            "time": str(row["사용시각"]),
            "store_name": str(row["사용장소"]),
            "purpose": str(row["사용목적(수기메모)"]),
            "attendees": str(row["대상인원"]),
            "card_last4": str(row["카드끝4자리"]),
            "amount": int(row["금액(원)"]),
            "user_name": u_name,
            "department": ""
        })
    st.session_state.receipt_data = updated_receipts

    st.write("")

    # ==============================================================================
    # [STEP 3] 정리된 엑셀 파일 저장 및 다운로드 영역
    # ==============================================================================
    st.markdown("""
    <div class="step-header">
        <span>💾 [3단계] 정리된 엑셀 파일 저장</span>
    </div>
    """, unsafe_allow_html=True)

    try:
        # 임시 템플릿 파일 생성
        temp_tpl_path = "temp_active_template.xlsx"
        if st.session_state.excel_template_bytes:
            with open(temp_tpl_path, "wb") as f:
                f.write(st.session_state.excel_template_bytes)
        else:
            temp_tpl_path = DEFAULT_EXCEL_TEMPLATE

        excel_bytes = excel_manager.fill_receipts_to_excel(
            template_path=temp_tpl_path,
            receipts=st.session_state.receipt_data
        )

        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        out_filename = f"법인카드_사용내역서_정리완료_{now_str}.xlsx"

        st.download_button(
            label="📥 정리된 엑셀 파일 저장 및 다운로드 (.xlsx)",
            data=excel_bytes,
            file_name=out_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        st.caption("✔ 원본 엑셀의 서식, 폰트, 테두리 및 합계 수식(=SUM(H6:H17))이 100% 보존된 채로 저장됩니다.")
    except Exception as e:
        st.error(f"엑셀 파일 생성 중 오류: {e}")

# 하단 세부 품목 아코디언
st.markdown("---")
with st.expander("🔍 영수증별 세부 품목 목록 및 개수 (인원 산출 근거)"):
    for r in st.session_state.receipt_data:
        st.markdown(f"**연번 {r.get('seq')}번 | {r.get('store_name')} ({r.get('purpose', '수기메모 없음')}) - 총 {r.get('amount', 0):,}원 | 대상인원: {r.get('attendees')}**")
        items = r.get("items", [])
        if items:
            item_df = pd.DataFrame(items)
            item_df.columns = ["품목명", "수량", "금액"]
            st.table(item_df)
        else:
            st.text("품목 상세 정보가 없습니다.")
