# Auto-Spec: AI-based E-Form Variable Specifier (전자 서식 변수 자동 매핑 도구)

## 📖 프로젝트 개요 (Project Overview)
**Auto-Spec**은 다양한 형식(이미지, PDF, DOCX)의 전자 서식을 읽어 들여 텍스트를 추출하고, 이를 사전에 정의된 메타데이터(한글/영문 변수명)와 인공지능(AI)을 통해 의미론적(Semantic)으로 자동 매핑해 주는 데스크탑 애플리케이션입니다. 

외부 API에 의존하지 않고 사용자의 로컬 환경에서 **오프라인**으로 작동하며, 정교한 텍스트 인식(OCR)과 자연어 임베딩 모델(SentenceTransformers)을 결합하여 사람의 개입을 최소화하면서도 정확한 변수명 명세서(Specification)를 작성할 수 있도록 돕습니다.

---

## ✨ 주요 기능 (Key Features)
1. **멀티 포맷 문서 지원**
   * `.png`, `.jpg` 등 일반적인 이미지 형식 지원
   * **다중 페이지 PDF 완벽 지원**: 문서를 한 번에 스크롤 형식으로 읽어 들이고 전 페이지 OCR 수행
   * **DOCX 문서 지원**: 이미지 변환 없이 문서 내부의 단락과 표 텍스트를 초고속으로 직접 추출
2. **오프라인 AI 텍스트 추출 및 매핑**
   * `EasyOCR`을 이용한 고성능 다국어 텍스트 추출
   * `SentenceTransformers(paraphrase-multilingual-MiniLM-L12-v2)` 기반의 의미 유사도(Semantic Similarity) 모델로 가장 적합한 메타데이터 변수 자동 추천
3. **대화형(Interactive) 미리보기 및 피드백 UI**
   * A4 비율에 최적화된 넓은 뷰어 제공
   * 이미지 상의 텍스트 바운딩 박스를 **직접 클릭하여 매핑 대상에서 제외/포함** 가능
   * 데이터 표(Grid) 항목 클릭 시 원본 이미지의 텍스트 영역이 노란색으로 **깜빡이는(Blink)** 직관적 UX
4. **결과 자동화 및 내보내기**
   * 사용자 수동 보정 지원 (드롭다운을 통한 2~5순위 후보 선택 및 전체 변수 선택)
   * 매핑 결과를 엑셀(`.xlsx`), `.csv`, `.json` 형식의 명세서로 저장 후 자동으로 실행(열기)

---

## 🛠 사용 기술 (Tech Stack)
* **GUI / 프레임워크**: Python, PySide6 (Qt)
* **OCR (광학 문자 인식)**: EasyOCR, OpenCV
* **AI 언어 모델**: `sentence-transformers` (PyTorch)
* **문서 처리**: `PyMuPDF (fitz)` (PDF 처리), `python-docx` (Word 처리), `pandas` (데이터 테이블 관리)

---

## 🚀 프롬프트 및 작동 로직 (Core Prompt & Logic)
이 프로그램은 사용자의 지시(Prompt)와 피드백을 바탕으로 AI 어시스턴트와의 페어 프로그래밍을 통해 완성되었습니다. 시스템 내부적으로 작동하는 핵심 매핑 프롬프트/로직은 다음과 같습니다:

1. **데이터 규격화 (Metadata Formatting)**:
   * 입력된 메타데이터는 `1열: 영문 변수명`, `2열: 한글 변수명` 구조를 따릅니다.
   * AI는 `한글 변수명`을 기준으로 사용자 서식에서 추출된 텍스트와 의미를 대조합니다.
2. **의미 유사도 추론 (Semantic Matching)**:
   * 서식에 표기된 문구(예: "주민번호", "성함")와 메타데이터의 표준 변수명(예: "주민등록번호", "성명")이 글자 단위로 일치하지 않아도, 다국어 임베딩 벡터 간의 코사인 유사도(Cosine Similarity)를 계산하여 의미가 가장 가까운 항목 Top 5를 추론합니다.
3. **사용자 제어 (Human-in-the-loop)**:
   * 텍스트 추출 직후, 단순 안내 문구나 불필요한 약관 등은 사용자가 시각적으로 직접 클릭하여 매핑을 회피(Exclude)할 수 있도록 설계하여 AI 추론 자원의 낭비를 막고 명세서의 품질을 높입니다.

---

## 💻 설치 및 실행 방법 (How to Run)

1. **가상환경 설정 및 접속**
   ```bash
   python -m venv venv
   source venv/bin/activate  # Mac/Linux
   # venv\Scripts\activate   # Windows
   ```

2. **의존성 패키지 설치**
   ```bash
   pip install -r requirements.txt
   ```

3. **프로그램 실행**
   ```bash
   python ui_main.py
   ```

4. **사용 순서**
   * **1단계**: `1. 메타데이터 로드` - 변수명 매핑의 기준이 될 엑셀/CSV 파일 업로드
   * **2단계**: `2. 서식 문서 로드` - PDF, DOCX, 이미지 등 서식 파일 업로드
   * **3단계**: `3. 텍스트 추출 (OCR)` - 텍스트를 추출하고 불필요한 박스를 클릭해 매핑에서 제외
   * **4단계**: `4. 매핑 실행` - AI 의미 유사도 분석 시작
   * **5단계**: `5. 명세서 저장` - 최종 결과를 파일로 내보내고 확인
