#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build script for PDF Extractor
PyInstaller를 사용해 실행 파일을 생성하는 스크립트
"""

import os
import sys
import subprocess
import platform
from pathlib import Path

# Window CP1252 환경 stdout 인코딩 설정
try:
    sys.stdout.reconfigure(encoding="utf-8")
except AttributeError:
    pass

def get_build_command():
    """플랫폼별 빌드 명령어 생성"""
    system = platform.system()
    
    # 기본 PyInstaller 옵션
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name=pdf-extractor",
        "--add-data=pdf_extractor.py:.",
    ]
    
    # 숨겨진 import 추가
    hidden_imports = [
        "tkinter",
        "tkinter.ttk",
        "tkinter.filedialog", 
        "tkinter.messagebox",
        "tkinter.scrolledtext",
        "pandas",
        "openpyxl",
        "pdfplumber",
        "PyPDF2",
        "fitz",
        "camelot",
        "tabula",
        "cv2",
        "PIL",
        "pytesseract",
        "easyocr",
        "numpy"
    ]
    
    for imp in hidden_imports:
        cmd.extend(["--hidden-import", imp])
    
    # 제외할 모듈들 (크기 최적화)
    excludes = [
        "matplotlib",
        "scipy",
        "IPython",
        "jupyter",
        "notebook"
    ]
    
    for exc in excludes:
        cmd.extend(["--exclude-module", exc])
    
    # 플랫폼별 설정
    if system == "Windows":
        # 아이콘이 있다면 추가
        if os.path.exists("icon.ico"):
            cmd.extend(["--icon=icon.ico"])
        # 버전 정보가 있다면 추가
        if os.path.exists("version_info.txt"):
            cmd.extend(["--version-file=version_info.txt"])
    elif system == "Darwin":  # macOS
        # 아이콘이 있다면 추가
        if os.path.exists("icon.icns"):
            cmd.extend(["--icon=icon.icns"])
        cmd.extend(["--osx-bundle-identifier=com.pdfextractor.app"])
    
    # 메인 파일 추가
    cmd.append("main.py")
    
    return cmd


def create_spec_file():
    """spec 파일 생성 (고급 설정용)"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('pdf_extractor.py', '.')],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'pandas',
        'openpyxl',
        'pdfplumber',
        'PyPDF2',
        'fitz',
        'camelot',
        'tabula',
        'cv2',
        'PIL',
        'pytesseract',
        'easyocr',
        'numpy'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'IPython',
        'jupyter',
        'notebook'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='pdf-extractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# macOS용 앱 번들 생성
import platform
if platform.system() == 'Darwin':
    app = BUNDLE(
        exe,
        name='PDF Extractor.app',
        bundle_identifier='com.pdfextractor.app',
        version='0.0.1'
    )
'''
    
    with open('pdf-extractor.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)


def build_executable():
    """실행 파일 빌드"""
    print("PDF Extractor 빌드를 시작합니다...")
    
    # 빌드 디렉토리 정리
    if os.path.exists("build"):
        import shutil
        shutil.rmtree("build")
    if os.path.exists("dist"):
        import shutil
        shutil.rmtree("dist")
    
    # spec 파일 생성
    create_spec_file()
    
    try:
        # PyInstaller 실행
        system = platform.system()
        
        if os.path.exists("pdf-extractor.spec"):
            # spec 파일 사용
            cmd = ["pyinstaller", "--clean", "pdf-extractor.spec"]
        else:
            # 명령행 옵션 사용
            cmd = get_build_command()
        
        print(f"실행 명령어: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        print("빌드 성공!")
        print(f"출력 결과:\n{result.stdout}")
        
        # 생성된 파일 확인
        dist_path = Path("dist")
        if dist_path.exists():
            files = list(dist_path.iterdir())
            print(f"\n생성된 파일들:")
            for file in files:
                print(f"  - {file}")
                
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"빌드 실패: {e}")
        print(f"에러 출력:\n{e.stderr}")
        return False
    except Exception as e:
        print(f"예상치 못한 오류: {e}")
        return False


def check_dependencies():
    """필요한 의존성 확인"""
    print("의존성 확인 중...")
    
    # 패키지명과 실제 import 이름 매핑
    required_packages = {
        "tkinter": "tkinter",  # Python 기본 라이브러리
        "pandas": "pandas", 
        "openpyxl": "openpyxl",
        "pdfplumber": "pdfplumber",
        "PyPDF2": "PyPDF2",
        "PyMuPDF": "fitz",
        "camelot-py": "camelot",
        "tabula-py": "tabula",
        "opencv-python": "cv2",
        "Pillow": "PIL",
        "pyinstaller": "PyInstaller"
    }
    
    # 선택적 패키지 (없어도 빌드 진행 가능)
    optional_packages = {
        "pytesseract": "pytesseract",
        "easyocr": "easyocr"
    }
    
    missing_required = []
    missing_optional = []
    
    # 필수 패키지 확인
    for package_name, import_name in required_packages.items():
        try:
            __import__(import_name)
            print(f"✓ {package_name} 설치됨")
        except ImportError:
            print(f"✗ {package_name} 누락")
            missing_required.append(package_name)
    
    # 선택적 패키지 확인
    for package_name, import_name in optional_packages.items():
        try:
            __import__(import_name)
            print(f"✓ {package_name} 설치됨 (선택사항)")
        except ImportError:
            print(f"⚠ {package_name} 누락 (선택사항)")
            missing_optional.append(package_name)
    
    if missing_required:
        print(f"\n필수 패키지 누락: {missing_required}")
        print("다음 명령어로 설치하세요:")
        print(f"pip install {' '.join(missing_required)}")
        return False
    
    if missing_optional:
        print(f"\n선택적 패키지 누락: {missing_optional}")
        print("이 패키지들은 일부 기능에만 필요하므로 빌드를 계속 진행합니다.")
    
    print("모든 필수 의존성이 설치되어 있습니다.")
    return True


def main():
    """메인 함수"""
    print("=== PDF Extractor 빌드 스크립트 ===")
    print(f"플랫폼: {platform.system()} {platform.machine()}")
    print(f"Python 버전: {sys.version}")
    
    # 의존성 확인
    if not check_dependencies():
        print("의존성 문제로 빌드를 중단합니다.")
        sys.exit(1)
    
    # 빌드 실행
    success = build_executable()
    
    if success:
        print("\n🎉 Tkinter 기반 PDF Extractor 빌드가 성공적으로 완료되었습니다!")
        print("📁 dist/ 폴더에서 실행 파일을 확인하세요.")
        print("🚀 PyQt5 의존성 문제가 해결되어 안정적인 빌드가 가능합니다.")
    else:
        print("\n❌ 빌드에 실패했습니다.")
        sys.exit(1)


if __name__ == "__main__":
    main()
