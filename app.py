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
    initial_sidebar_state="expanded"
)

# 2. 커스텀 CSS 스타일링
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
        margin-bottom: 1.2rem;
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
    .api-badge-ok {
        background-color: #ECFDF5;
        border: 1px solid #A7F3D0;
        color: #065F46;
        padding: 8px 12px;
        border-radius: 8px;
        font-size: 0.88rem;
        font-weight: 600;
        margin-bottom: 12px;
    }
    .api-badge-warn {
        background-color: #FFFBEB;
        border: 1px solid #FDE68A;
        color: #92400E;
        padding: 8px 12px;
        border-radius: 8px;
        font-size: 0.88rem;
        font-weight: 600;
        margin-bottom: 12px;
    }
</style>
""", unsafe_allow_html=True)

DEFAULT_EXCEL_TEMPLATE = "법인카드 사용내역서.xlsx"

# 3. API 키 감지 함수 (Streamlit Secrets / 사용자 입력 / 환경변수)
def detect_gemini_api_key(manual_key: str = "") -> tuple[str, str]:
    # 1) 사이드바 직접 입력값 우선
    if manual_key and manual_key.strip():
        return manual_key.strip(), "사이드바 직접 입력"
    
    # 2) Streamlit Cloud Secrets 확인
    try:
        if hasattr(st, "secrets"):
            if "GEMINI_API_KEY" in st.secrets and str(st.secrets["GEMINI_API_KEY"]).strip():
                return str(st.secrets["GEMINI_API_KEY"]).strip(), "Streamlit Secrets 연동"
            if "GOOGLE_API_KEY" in st.secrets and str(st.secrets["GOOGLE_API_KEY"]).strip():
                return str(st.secrets["GOOGLE_API_KEY"]).strip(), "Streamlit Secrets 연동"
    except Exception:
        pass

    # 3) OS 환경변수 확인
    env_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if env_key and env_key.strip():
        return env_key.strip(), "환경 변수(OS)"

    return "", "미설정"

# 4. 세션 상태 초기화 (클린 세션)
if "receipt_data" not in st.session_state:
    st.session_state.receipt_data = []

if "pdf_images" not in st.session_state:
    st.session_state.pdf_images = []

if "excel_template_bytes" not in st.session_state:
    if os.path.exists(DEFAULT_EXCEL_TEMPLATE):
        with open(DEFAULT_EXCEL_TEMPLATE, "rb") as f:
            st.session_state.excel_template_bytes = f.read()
    else:
        st.session_state.excel_template_bytes = None

def format_display_date(val):
    """표시용 날짜 포맷 (YYYY-MM-DD)"""
    if not val:
        return ""
    s = str(val).replace("-", "").replace(".", "").replace("/", "").strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    return str(val)

# 상단 헤더
st.markdown('<div class="main-title">🧾 법인카드 영수증 자동 정리 & 엑셀 생성</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">영수증 스캔 파일과 엑셀 양식을 업로드하고 <b>[실행]</b>을 누르면, '
    'AI가 가맹점/일시/금액과 <b>수기 메모(보건소명/병원명 등 사용목적)</b>를 인식하여 엑셀에 자동으로 정리합니다.</div>',
    unsafe_allow_html=True
)

# 사이드바
with st.sidebar:
    st.header("⚙️ 시스템 설정")
    
    sidebar_key_input = st.text_input(
        "Gemini API Key 직접 입력",
        type="password",
        value="",
        help="Google AI Studio(aistudio.google.com)에서 발급받은 API 키를 입력하세요. Streamlit Secrets에 등록되어 있다면 비워두셔도 자동 감지됩니다."
    )

    active_api_key, key_source = detect_gemini_api_key(sidebar_key_input)
    
    if active_api_key:
        st.markdown(f'<div class="api-badge-ok">🟢 Gemini AI 활성화됨<br><span style="font-size:0.78rem; font-weight:normal;">({key_source})</span></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="api-badge-warn">⚠️ API Key 필요<br><span style="font-size:0.78rem; font-weight:normal;">새 영수증 분석을 위해 위에 키를 입력하거나 Secrets에 등록하세요.</span></div>', unsafe_allow_html=True)
        st.markdown("[👉 무료 API Key 발급받기](https://aistudio.google.com/app/apikey)")

    st.markdown("---")
    st.subheader("🧪 샘플 데이터 테스트")
    if st.button("내장 샘플 영수증 불러오기", use_container_width=True):
        st.session_state.receipt_data = [dict(r) for r in SAMPLE_RECEIPTS]
        sample_pdf = "영수증.pdf"
        if os.path.exists(sample_pdf):
            with open(sample_pdf, "rb") as f:
                st.session_state.pdf_images = receipt_parser.extract_images_from_pdf(f.read())
        st.success("내장 샘플(파리바게뜨 5건)이 로드되었습니다.")
        st.rerun()

    if st.button("🗑️ 전체 데이터 초기화 (비우기)", type="secondary", use_container_width=True):
        st.session_state.receipt_data = []
        st.session_state.pdf_images = []
        st.info("데이터가 초기화되었습니다.")
        st.rerun()

# ==============================================================================
# [STEP 1] 파일 업로드 및 실행 영역
# ==============================================================================
st.markdown("""
<div class="step-header">
    <span>📂 [1단계] 파일 업로드</span>
</div>
""", unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2)

with col_up1:
    uploaded_excel = st.file_uploader(
        "1. 정리해서 넣을 엑셀 파일 (.xlsx) [선택]",
        type=["xlsx"],
        key="uploader_excel",
        help="법인카드 사용내역서 양식 엑셀 파일입니다. (미업로드 시 기본 표준 서식이 적용됩니다.)"
    )
    if uploaded_excel:
        st.session_state.excel_template_bytes = uploaded_excel.read()
        st.caption(f"✔ 업로드된 서식 사용: `{uploaded_excel.name}`")
    else:
        st.caption("ℹ 미업로드 시 기본 '법인카드 사용내역서.xlsx' 서식으로 자동 적용됩니다.")

with col_up2:
    uploaded_receipt = st.file_uploader(
        "2. 스캔한 영수증 파일 (PDF 또는 이미지) [필수]",
        type=["pdf", "png", "jpg", "jpeg"],
        key="uploader_receipt",
        help="분석할 스캔 영수증 파일(PDF, PNG, JPG)을 업로드하세요."
    )
    if uploaded_receipt:
        st.caption(f"✔ 분석 대기 중: `{uploaded_receipt.name}` ({len(uploaded_receipt.getvalue()) // 1024:,} KB)")
    else:
        st.caption("※ 분석할 영수증 파일을 여기에 업로드해주세요.")

# 🚀 실행 버튼
col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])

with col_btn2:
    run_btn = st.button("🚀 영수증 분석 & 엑셀 자동 정리 실행", type="primary", use_container_width=True)

with col_btn3:
    if st.session_state.receipt_data or st.session_state.pdf_images:
        if st.button("🗑️ 현재 내역 지우기", use_container_width=True):
            st.session_state.receipt_data = []
            st.session_state.pdf_images = []
            st.rerun()

if run_btn:
    if not uploaded_receipt:
        st.warning("⚠️ 스캔한 영수증 파일(PDF 또는 이미지)을 먼저 업로드해 주세요.")
    elif not active_api_key:
        st.error("❌ Gemini API 키가 감지되지 않았습니다! 좌측 사이드바에 API 키를 입력하시거나 Streamlit Secrets에 GEMINI_API_KEY를 설정해주세요.")
    else:
        receipt_bytes = uploaded_receipt.getvalue()
        receipt_filename = uploaded_receipt.name

        # 1. 이미지 렌더링 (미리보기용)
        try:
            if receipt_filename.lower().endswith(".pdf"):
                st.session_state.pdf_images = receipt_parser.extract_images_from_pdf(receipt_bytes)
            else:
                st.session_state.pdf_images = [Image.open(io.BytesIO(receipt_bytes))]
        except Exception as img_err:
            st.error(f"영수증 이미지를 불러오는 중 오류 발생: {img_err}")
            st.session_state.pdf_images = []

        # 2. 진행 상태 프로그레스 바 준비
        progress_bar = st.progress(0)
        status_text = st.empty()

        def on_progress(current_page, total_pages, msg):
            pct = int((current_page / total_pages) * 100)
            progress_bar.progress(pct)
            status_text.info(f"⏳ {msg}")

        # 3. AI 파싱 실행
        try:
            with st.spinner("AI가 영수증 이미지와 수기 메모를 정밀 분석 중입니다..."):
                parsed_results = receipt_parser.parse_receipt_document(
                    file_bytes=receipt_bytes,
                    file_name=receipt_filename,
                    api_key=active_api_key,
                    progress_callback=on_progress
                )
                st.session_state.receipt_data = parsed_results
                progress_bar.progress(100)
                status_text.empty()
                st.success(f"🎉 총 {len(parsed_results)}건의 영수증 내역을 성공적으로 인식하여 정리했습니다!")
                st.rerun()
        except Exception as parse_err:
            progress_bar.empty()
            status_text.empty()
            st.error(f"❌ 영수증 AI 분석 실패: {parse_err}")

st.markdown("---")

# ==============================================================================
# [STEP 2] 결과 검토 및 실시간 수정 영역
# ==============================================================================
if st.session_state.receipt_data:
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

    # KPI 지표 카드
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
                caption=f"스캔본 {page_idx}페이지 실물",
                use_container_width=True
            )
            st.info("💡 영수증 여백에 적힌 수기 메모와 품목/금액을 우측 표와 직접 대조해보실 수 있습니다.")
        else:
            st.info("표시할 영수증 이미지가 없습니다.")

    with col_edit:
        st.subheader("📊 정리된 법인카드 내역서 (직접 수정 가능)")
        st.caption("※ 셀을 더블 클릭하여 수정할 수 있습니다. 사용자 등 확인 불가 항목은 기본 공란으로 비워져 있습니다.")

        display_rows = []
        for r in st.session_state.receipt_data:
            display_rows.append({
                "연번": r.get("seq", 1),
                "사용일자": format_display_date(r.get("date", "")),
                "사용시각": str(r.get("time", "")),
                "사용장소": r.get("store_name", ""),
                "사용목적(수기메모)": r.get("purpose", ""),
                "대상인원": r.get("attendees", ""),
                "카드끝4자리": str(r.get("card_last4", "")),
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
                    "대상인원 (자동산출)",
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
            template_source = st.session_state.excel_template_bytes if st.session_state.excel_template_bytes else DEFAULT_EXCEL_TEMPLATE

            excel_bytes = excel_manager.fill_receipts_to_excel(
                template_source=template_source,
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
            st.caption("✔ 원본 엑셀 서식, 글꼴, 셀 테두리 및 합계 수식(=SUM(H6:H17))이 100% 보존되어 저장됩니다.")
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
else:
    st.info("👆 위 [1단계]에서 영수증 스캔 파일(PDF 또는 이미지)을 업로드하고 **[실행]** 버튼을 누르면 새 영수증 분석 결과가 나타납니다.")
