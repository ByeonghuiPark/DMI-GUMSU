# 🚀 HWPX 자동화 도구 EXE 빌드 가이드

## 📋 준비사항

### 1. Python 설치 (3.8 이상)
```bash
# Python 다운로드
https://www.python.org/downloads/

# 설치 시 주의사항:
☑️ "Add Python to PATH" 체크
☑️ "Install for all users" 체크 (권장)
```

### 2. Tesseract OCR 설치
```bash
# Windows
https://github.com/UB-Mannheim/tesseract/wiki
→ tesseract-ocr-w64-setup-v5.3.0.exe 다운로드
→ 설치 시 "Additional language data" → Korean 선택

# 설치 경로 (기본값)
C:\Program Files\Tesseract-OCR\

# 환경변수 설정 (자동으로 되지 않은 경우)
PATH에 C:\Program Files\Tesseract-OCR\ 추가
```

### 3. 필수 파일 준비
```
프로젝트 폴더/
├── hwpx_automation_gui_final.py
├── enhanced_document_extractor.py  
├── enhanced_hwpx_processor.py
├── hwpx_image_inserter.py
├── build_exe.py                    (빌드 스크립트)
├── build.bat                       (배치 파일)
└── requirements.txt                (의존성 목록)
```

## 🛠️ 빌드 방법

### 방법 1: 배치 파일 사용 (추천)
```bash
# Windows 탐색기에서 build.bat 더블클릭
# 또는 명령 프롬프트에서:
build.bat
```

### 방법 2: Python 스크립트 직접 실행
```bash
# 의존성 설치
pip install pyinstaller pdfplumber pillow pytesseract lxml

# 빌드 실행
python build_exe.py
```

### 방법 3: 수동 PyInstaller 실행
```bash
# 기본 빌드
pyinstaller --onefile --windowed hwpx_automation_gui_final.py

# 고급 옵션 포함
pyinstaller --onefile --windowed --clean --noconfirm ^
  --hidden-import=tkinter --hidden-import=pdfplumber ^
  --hidden-import=pytesseract --hidden-import=PIL ^
  --add-data="enhanced_document_extractor.py;." ^
  --add-data="enhanced_hwpx_processor.py;." ^
  --add-data="hwpx_image_inserter.py;." ^
  hwpx_automation_gui_final.py
```

## 📦 빌드 과정

### 1. 의존성 확인
```
✅ pyinstaller    (EXE 빌드)
✅ tkinter        (GUI)
✅ pdfplumber     (PDF 처리)
✅ pytesseract    (OCR)
✅ Pillow         (이미지 처리)
✅ lxml           (XML 처리)
```

### 2. Tesseract 확인
```
🔍 C:\Program Files\Tesseract-OCR\ 경로 확인
🔍 tesseract.exe 실행 파일 확인  
🔍 tessdata\ 폴더 (언어 데이터) 확인
```

### 3. 빌드 실행
```
🧹 이전 빌드 파일 정리 (build/, dist/)
📦 PyInstaller 실행
🎯 단일 EXE 파일 생성
📁 추가 파일 복사 (사용법, 샘플)
```

### 4. 결과 확인
```
dist/
├── HWPX_Automation.exe     (메인 실행 파일)
├── 사용법.txt              (사용자 가이드)
├── sample_terms.csv        (샘플 참조 데이터)
└── sample_terms.json       (샘플 참조 데이터)
```

## ⚙️ 빌드 옵션 설명

### PyInstaller 주요 옵션
```bash
--onefile          # 단일 EXE 파일 생성
--windowed          # 콘솔 창 숨기기
--clean             # 임시 파일 정리
--noconfirm         # 기존 파일 덮어쓰기
--hidden-import     # 숨겨진 모듈 포함
--add-data          # 데이터 파일 포함
--add-binary        # 바이너리 파일 포함
--exclude-module    # 불필요한 모듈 제외
--icon              # 아이콘 설정
```

### 크기 최적화 옵션
```bash
--exclude-module=matplotlib
--exclude-module=numpy.random._pickle
--exclude-module=tkinter.test
--exclude-module=unittest
--upx-dir=C:\upx    # UPX 압축 (별도 설치 필요)
```

## 🧪 테스트

### 1. 로컬 테스트
```bash
# 빌드 후 즉시 실행
dist\HWPX_Automation.exe

# 기능 테스트 체크리스트:
☑️ GUI 정상 실행
☑️ 파일 선택 대화상자
☑️ PDF 문서 분석
☑️ 이미지 OCR 처리
☑️ HWPX 파일 텍스트 치환
☑️ 이미지 삽입
```

### 2. 다른 PC에서 테스트
```bash
# 테스트 환경:
- Windows 10/11
- Tesseract 미설치 PC
- Python 미설치 PC
- 방화벽/백신 환경

# 확인사항:
☑️ 실행 파일 정상 구동
☑️ OCR 기능 작동 (Tesseract 포함된 경우)
☑️ 모든 기능 정상 작동
```

## 🔧 문제해결

### 빌드 실패 시
```bash
# 1. Python 버전 확인
python --version  # 3.8 이상 필요

# 2. 패키지 재설치
pip uninstall pyinstaller
pip install pyinstaller

# 3. 가상환경 사용
python -m venv build_env
build_env\Scripts\activate
pip install -r requirements.txt
python build_exe.py

# 4. 로그 확인
pyinstaller --log-level=DEBUG [options] script.py
```

### 실행 파일 오류 시
```bash
# 1. 의존성 확인
pyinstaller --collect-all tkinter script.py

# 2. 숨겨진 임포트 추가
--hidden-import=missing_module

# 3. 콘솔 모드로 디버깅
pyinstaller --onefile --console script.py
```

### OCR 오류 시
```bash
# 1. Tesseract 경로 확인
where tesseract

# 2. 언어 데이터 확인
tesseract --list-langs

# 3. 수동 경로 설정
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
```

## 📊 최적화 팁

### 1. 파일 크기 줄이기
```bash
# 불필요한 모듈 제외
--exclude-module=matplotlib,scipy,numpy

# UPX 압축 사용
# https://upx.github.io/ 에서 UPX 다운로드
--upx-dir=C:\upx

# 가상환경에서 빌드 (깨끗한 환경)
python -m venv clean_env
clean_env\Scripts\activate
pip install [필수 패키지만]
```

### 2. 실행 속도 향상
```bash
# 사전 컴파일
--optimize=2

# 단일 파일 대신 디렉토리 배포
--onedir  # 대신 --onefile 사용하지 않음
```

### 3. 호환성 향상
```bash
# Windows 호환성
--target-arch=x86_64  # 64비트
--target-arch=x86     # 32비트 (더 넓은 호환성)

# 디지털 서명 (코드 서명 인증서 있는 경우)
signtool sign /f cert.p12 /p password dist\app.exe
```

## 🎯 배포 준비

### 1. 파일 구성
```
HWPX_Automation_v1.0/
├── HWPX_Automation.exe     # 메인 실행 파일
├── 사용법.txt              # 한글 사용 가이드
├── README.md               # 영문 가이드
├── sample_terms.csv        # 샘플 CSV
├── sample_terms.json       # 샘플 JSON
└── 설치가이드.txt          # 설치 도움말
```

### 2. 압축 파일 생성
```bash
# ZIP 파일로 배포
"HWPX_Automation_v1.0.zip"

# 또는 설치 프로그램 생성 (NSIS, Inno Setup 등)
```

### 3. 배포 체크리스트
```
☑️ 다양한 Windows 버전에서 테스트 완료
☑️ 바이러스 검사 완료
☑️ 디지털 서명 적용 (선택사항)
☑️ 사용자 가이드 포함
☑️ 샘플 데이터 포함
☑️ 버전 정보 명시
```

이제 완전한 EXE 빌드 시스템이 준비되었습니다! 🎉