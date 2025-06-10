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
        "PyQt5.QtCore",
        "PyQt5.QtWidgets", 
        "PyQt5.QtGui",
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
        cmd.extend([
            "--icon=icon.ico",  # 아이콘이 있다면
            "--version-file=version_info.txt"  # 버전 정보가 있다면
        ])
    elif system == "Darwin":  # macOS
        cmd.extend([
            "--icon=icon.icns",  # 아이콘이 있다면
            "--osx-bundle-identifier=com.pdfextractor.app"
        ])
    
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
        'PyQt5.QtCore',
        'PyQt5.QtWidgets', 
        'PyQt5.QtGui',
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
        icon='icon.icns',
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
    
    required_packages = [
        "PyQt5", "pandas", "openpyxl", "pdfplumber", 
        "PyPDF2", "PyMuPDF", "camelot-py", "tabula-py",
        "opencv-python", "Pillow", "pytesseract", "easyocr",
        "pyinstaller"
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package.replace("-", "_"))
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print(f"누락된 패키지: {missing_packages}")
        print("다음 명령어로 설치하세요:")
        print(f"pip install {' '.join(missing_packages)}")
        return False
    
    print("모든 의존성이 설치되어 있습니다.")
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
        print("\n🎉 빌드가 성공적으로 완료되었습니다!")
        print("dist/ 폴더에서 실행 파일을 확인하세요.")
    else:
        print("\n❌ 빌드에 실패했습니다.")
        sys.exit(1)


if __name__ == "__main__":
    main()
