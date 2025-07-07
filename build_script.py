#!/usr/bin/env python3
"""
HWPX 자동화 도구 EXE 빌드 스크립트
PyInstaller를 사용하여 단일 EXE 파일 생성
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path
import platform

def check_dependencies():
    """필수 의존성 확인"""
    print("🔍 의존성 확인 중...")
    
    required_packages = [
        'pyinstaller', 'tkinter', 'pdfplumber', 'Pillow', 
        'pytesseract', 'lxml', 'pathlib'
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            if package == 'tkinter':
                import tkinter
            elif package == 'pyinstaller':
                import PyInstaller
            elif package == 'pdfplumber':
                import pdfplumber
            elif package == 'Pillow':
                import PIL
            elif package == 'pytesseract':
                import pytesseract
            elif package == 'lxml':
                import lxml
            elif package == 'pathlib':
                import pathlib
            print(f"✅ {package}")
        except ImportError:
            missing_packages.append(package)
            print(f"❌ {package}")
    
    if missing_packages:
        print(f"\n📦 다음 패키지를 설치해주세요:")
        for package in missing_packages:
            print(f"  pip install {package}")
        return False
    
    return True

def find_tesseract():
    """Tesseract 설치 경로 찾기"""
    possible_paths = [
        "C:\\Program Files\\Tesseract-OCR",
        "C:\\Program Files (x86)\\Tesseract-OCR",
        "/usr/bin",
        "/usr/local/bin",
        "/opt/homebrew/bin"
    ]
    
    for path in possible_paths:
        tesseract_exe = os.path.join(path, "tesseract.exe" if platform.system() == "Windows" else "tesseract")
        if os.path.exists(tesseract_exe):
            print(f"✅ Tesseract 발견: {path}")
            return path
    
    print("⚠️ Tesseract를 찾을 수 없습니다. 수동으로 설치해주세요.")
    return None

def create_requirements_txt():
    """requirements.txt 파일 생성"""
    requirements = """
# HWPX 자동화 도구 의존성
pyinstaller>=5.0.0
pdfplumber>=0.7.0
Pillow>=9.0.0
pytesseract>=0.3.10
lxml>=4.6.0
chardet>=4.0.0
tqdm>=4.60.0
pandas>=1.3.0
colorlog>=6.0.0
regex>=2021.0.0
opencv-python>=4.5.0
pathlib2>=2.3.0
""".strip()
    
    with open("requirements.txt", "w", encoding="utf-8") as f:
        f.write(requirements)
    print("📝 requirements.txt 생성됨")

def build_exe():
    """EXE 파일 빌드"""
    print("\n🚀 EXE 빌드 시작")
    
    # 프로젝트 설정
    main_script = "hwpx_automation_gui_final.py"
    app_name = "HWPX_Automation"
    
    # 메인 스크립트 존재 확인
    if not os.path.exists(main_script):
        print(f"❌ 메인 스크립트를 찾을 수 없습니다: {main_script}")
        return False
    
    # 기본 PyInstaller 명령어
    cmd = [
        "pyinstaller",
        "--name", app_name,
        "--onefile",                    # 단일 EXE 파일
        "--windowed",                   # 콘솔 창 숨기기
        "--clean",                      # 임시 파일 정리
        "--noconfirm",                 # 기존 파일 덮어쓰기
        
        # 숨겨진 임포트
        "--hidden-import", "PIL._tkinter_finder",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.scrolledtext",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.messagebox",
        "--hidden-import", "pdfplumber",
        "--hidden-import", "pytesseract",
        "--hidden-import", "lxml",
        "--hidden-import", "lxml.etree",
        "--hidden-import", "xml.etree.ElementTree",
        "--hidden-import", "zipfile",
        "--hidden-import", "tempfile",
        "--hidden-import", "pathlib",
        "--hidden-import", "threading",
        "--hidden-import", "json",
        "--hidden-import", "csv",
        "--hidden-import", "re",
        "--hidden-import", "datetime",
        
        # 제외할 모듈 (크기 줄이기)
        "--exclude-module", "matplotlib",
        "--exclude-module", "numpy.random._pickle",
        "--exclude-module", "tkinter.test",
        "--exclude-module", "unittest",
        "--exclude-module", "test",
        "--exclude-module", "pydoc",
        "--exclude-module", "doctest",
        
        main_script
    ]
    
    # 추가 Python 파일들 포함
    additional_files = [
        "enhanced_document_extractor.py",
        "enhanced_hwpx_processor.py", 
        "hwpx_image_inserter.py"
    ]
    
    for file in additional_files:
        if os.path.exists(file):
            cmd.extend(["--add-data", f"{file};."])
            print(f"📄 포함됨: {file}")
    
    # Tesseract 포함 (Windows)
    if platform.system() == "Windows":
        tesseract_path = find_tesseract()
        if tesseract_path:
            tesseract_exe = os.path.join(tesseract_path, "tesseract.exe")
            tessdata_path = os.path.join(tesseract_path, "tessdata")
            
            if os.path.exists(tesseract_exe):
                cmd.extend(["--add-binary", f"{tesseract_exe};tesseract"])
            
            if os.path.exists(tessdata_path):
                cmd.extend(["--add-data", f"{tessdata_path};tessdata"])
    
    # 아이콘 파일 (있는 경우)
    icon_files = ["icon.ico", "assets/icon.ico", "resources/icon.ico"]
    for icon_file in icon_files:
        if os.path.exists(icon_file):
            cmd.extend(["--icon", icon_file])
            print(f"🎨 아이콘 설정: {icon_file}")
            break
    
    try:
        print("📦 PyInstaller 실행 중...")
        print(f"💻 명령어: {' '.join(cmd)}")
        
        # PyInstaller 실행
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode == 0:
            # 성공
            exe_path = Path("dist") / f"{app_name}.exe"
            if exe_path.exists():
                file_size = exe_path.stat().st_size / (1024 * 1024)  # MB
                print(f"\n✅ 빌드 성공!")
                print(f"📁 위치: {exe_path.absolute()}")
                print(f"📏 크기: {file_size:.1f} MB")
                
                # 추가 파일 복사
                copy_additional_files()
                
                return True
            else:
                print("❌ EXE 파일이 생성되지 않았습니다.")
                return False
        else:
            # 실패
            print(f"❌ PyInstaller 실행 실패:")
            print(f"stdout: {result.stdout}")
            print(f"stderr: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ 빌드 오류: {e}")
        return False

def copy_additional_files():
    """추가 파일들을 dist 폴더에 복사"""
    dist_dir = Path("dist")
    
    # 복사할 파일들
    files_to_copy = [
        "README.md",
        "사용법.txt",
        "sample_terms.csv",
        "sample_terms.json"
    ]
    
    for file_name in files_to_copy:
        if os.path.exists(file_name):
            try:
                shutil.copy2(file_name, dist_dir / file_name)
                print(f"📄 복사됨: {file_name}")
            except Exception as e:
                print(f"⚠️ 복사 실패 {file_name}: {e}")
    
    # 사용자 가이드 생성
    create_user_guide(dist_dir)

def create_user_guide(dist_dir):
    """사용자 가이드 파일 생성"""
    guide_content = """
# HWPX 자동화 도구 사용법

## 📋 개요
이 도구는 세금계산서나 사업자등록증에서 정보를 추출하여 
HWPX 파일의 텍스트를 자동으로 치환하고 이미지를 삽입하는 프로그램입니다.

## 🚀 시작하기

### 1단계: 문서 분석
- "파일 선택" 버튼을 클릭하여 세금계산서(PDF) 또는 사업자등록증(이미지) 선택
- "문서 분석" 버튼을 클릭하여 정보 추출
- 추출된 정보가 오른쪽에 표시됩니다

### 2단계: 참조 데이터 생성
- "자동 참조 데이터 생성" 버튼으로 자동 생성하거나
- "기존 파일 불러오기"로 CSV/JSON 파일 로드
- 필요시 수동으로 항목 추가/삭제/편집

### 3단계: 텍스트 치환
- HWPX 템플릿 파일 선택
- 치환 옵션 설정 (대소문자 구분, 전체 단어 매칭 등)
- "텍스트 치환 실행" 버튼 클릭

### 4단계: 이미지 삽입 (선택사항)
- 삽입할 이미지 파일 선택
- 삽입 위치 설정 (표 번호, 행, 열)
- 이미지 크기 및 정렬 옵션 설정
- "이미지 삽입 실행" 버튼 클릭

## 📁 지원 파일 형식

### 입력 문서:
- PDF: 세금계산서 등
- 이미지: PNG, JPG, JPEG, BMP, TIFF (사업자등록증 등)
- 텍스트: TXT 파일

### 참조 데이터:
- JSON: {"원본": "치환"}
- CSV: 1열(원본), 2열(치환)

### 출력:
- HWPX: 한글 문서 파일
- TXT: 치환된 텍스트

## ⚠️ 주의사항
- 이미지 파일은 OCR 처리되므로 선명한 이미지를 사용하세요
- 큰 파일 처리 시 시간이 걸릴 수 있습니다
- 원본 파일은 자동으로 백업됩니다

## 🛠️ 문제해결
- OCR 오류: 이미지 품질을 확인하세요
- 한글 인식 문제: Tesseract 한국어 데이터 설치 확인
- 메모리 부족: 큰 파일은 나누어 처리하세요

## 📞 지원
문제 발생 시 로그 메시지를 참조하거나 개발자에게 문의하세요.
""".strip()
    
    guide_path = dist_dir / "사용법.txt"
    with open(guide_path, "w", encoding="utf-8") as f:
        f.write(guide_content)
    print(f"📖 사용자 가이드 생성: {guide_path}")

def clean_build():
    """빌드 임시 파일 정리"""
    dirs_to_clean = ["build", "__pycache__", "*.egg-info"]
    files_to_clean = ["*.spec"]
    
    print("🧹 임시 파일 정리 중...")
    
    # 디렉토리 삭제
    for dir_pattern in dirs_to_clean:
        if "*" in dir_pattern:
            # 와일드카드 패턴
            import glob
            for path in glob.glob(dir_pattern):
                if os.path.isdir(path):
                    shutil.rmtree(path)
                    print(f"🗑️ 삭제됨: {path}")
        else:
            # 일반 디렉토리
            if os.path.exists(dir_pattern):
                shutil.rmtree(dir_pattern)
                print(f"🗑️ 삭제됨: {dir_pattern}")
    
    # 파일 삭제
    import glob
    for file_pattern in files_to_clean:
        for file_path in glob.glob(file_pattern):
            os.remove(file_path)
            print(f"🗑️ 삭제됨: {file_path}")

def main():
    """메인 함수"""
    print("🎯 HWPX 자동화 도구 EXE 빌더")
    print("=" * 50)
    
    # 명령행 인수 처리
    if "--clean" in sys.argv:
        clean_build()
        return
    
    if "--requirements" in sys.argv:
        create_requirements_txt()
        return
    