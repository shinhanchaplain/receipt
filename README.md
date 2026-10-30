# 🧾 스마트 법인카드 영수증 파싱 & 엑셀 자동화 대시보드

스캔한 영수증(PDF 및 이미지)에서 가맹점, 결제일시, 금액, 구매품목은 물론 **볼펜/매직으로 기재된 수기 메모(보건소명/병원명 등)**를 AI(Gemini Vision)로 정밀 인식하여 `법인카드 사용내역서.xlsx` 양식에 자동으로 매핑하고 다운로드할 수 있는 웹 대시보드 애플리케이션입니다.

---

## ✨ 주요 기능

1. **스캔 영수증 PDF/이미지 자동 분석**:
   - PDF 내 각 페이지 분할 및 고화질 렌더링
   - 한 페이지에 여러 장 부착된 영수증 개별 분리 인식
2. **수기(손글씨) 메모 완벽 인식**:
   - 영수증 여백에 적힌 '이정복 보건소', '영등포 보건소', '동작구 보건소', '양주시 보건소' 등 수기 목적을 정확히 추출하여 **사용목적(구체적 명기)** 컬럼에 자동 입력
3. **스마트 대상인원 산출**:
   - 총 구매 수량(개수)과 결제 금액을 감안하여 대상인원(예: `3명`, `8명`)을 자동 제안
4. **인터랙티브 웹 대시보드 (Streamlit)**:
   - 좌측: 스캔 영수증 실물 뷰어 (페이지별 고해상도 미리보기)
   - 우측: 실시간 데이터 편집 그리드 (`st.data_editor`)를 통한 간편 수정
5. **기존 엑셀 서식 100% 보존**:
   - 기존 `법인카드 사용내역서.xlsx`의 폰트, 테두리, 배경색 및 하단 소계 수식(`=SUM(H6:H17)`)을 그대로 유지하며 데이터 주입
6. **무료 웹 호스팅 지원**:
   - Streamlit Community Cloud를 통해 몇 분 안에 누구나 웹페이지 URL로 무료 배포 가능

---

## 🚀 로컬 실행 방법

### 1. 패키지 설치
```bash
pip install -r requirements.txt
```

### 2. Streamlit 대시보드 실행
```bash
streamlit run app.py
```
브라우저에서 `http://localhost:8501`로 대시보드가 자동으로 열립니다.

---

## 🌐 무료 웹페이지 앱으로 배포(게시)하는 방법

### 1단계: GitHub 저장소에 코드 업로드
1. GitHub(https://github.com)에 로그인 후 새 레포지토리(예: `receipt-dashboard`)를 생성합니다.
2. 작업 폴더의 파일들을 Git으로 커밋하고 푸시합니다:
```bash
git init
git add .
git commit -m "feat: 영수증 엑셀 자동화 대시보드 완성"
git branch -M main
git remote add origin https://github.com/사용자아이디/receipt-dashboard.git
git push -u origin main
```

### 2단계: Streamlit Community Cloud에 배포
1. [Streamlit Community Cloud](https://share.streamlit.io/)에 접속하여 GitHub 계정으로 로그인합니다.
2. **`New app`** 버튼 클릭
3. **Repository**: 방금 생성한 `receipt-dashboard` 선택
4. **Main file path**: `app.py` 입력
5. **Deploy!** 버튼을 클릭하면 1~2분 만에 전 세계 어디서나 접속 가능한 고유 URL이 생성됩니다.
6. (선택사항) Streamlit 앱 설정의 `Secrets` 탭에 `GEMINI_API_KEY="본인키"`를 입력해 두면 모든 사용자가 별도 키 입력 없이 바로 신규 영수증을 인식할 수 있습니다.

---

## 📂 프로젝트 구조

- `app.py`: Streamlit 웹 대시보드 메인 UI
- `receipt_parser.py`: PDF 페이지 렌더링 및 Gemini Vision 수기 메모/영수증 인식 엔진
- `excel_manager.py`: openpyxl 기반 원본 서식/수식 유지 엑셀 데이터 주입기
- `sample_data.py`: 기본 스캔본(영수증.pdf) 사전 분석 데이터셋
- `법인카드 사용내역서.xlsx`: 회사/기관 표준 법인카드 양식 템플릿
- `영수증.pdf`: 3페이지 분량의 스캔 영수증 실물 예시
- `requirements.txt`: 필수 파이썬 라이브러리 목록
