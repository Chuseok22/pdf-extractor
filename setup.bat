@echo off
REM PDF Extractor 개발 환경 설정 스크립트 (Windows)

echo 🚀 PDF Extractor 개발 환경을 설정합니다...

REM Python 버전 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python이 설치되어 있지 않습니다.
    echo https://python.org에서 Python 3.8 이상을 설치하세요.
    pause
    exit /b 1
)

echo ✅ Python 설치 확인됨

REM 가상환경 생성
if not exist ".venv" (
    echo 📦 Python 가상환경을 생성합니다...
    python -m venv .venv
    echo ✅ 가상환경 생성 완료
) else (
    echo ✅ 가상환경이 이미 존재합니다
)

REM 가상환경 활성화
echo 🔄 가상환경을 활성화합니다...
call .venv\Scripts\activate.bat

REM pip 업그레이드
echo 📈 pip을 업그레이드합니다...
python -m pip install --upgrade pip

REM 의존성 설치
echo 📦 의존성을 설치합니다...
pip install -r requirements.txt

REM Chocolatey 확인
echo 🔧 시스템 의존성을 확인합니다...
choco --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Chocolatey가 설치되어 있지 않습니다.
    echo 다음 사이트에서 Chocolatey를 설치하세요: https://chocolatey.org/install
    echo 또는 Tesseract를 수동으로 설치하세요: https://github.com/UB-Mannheim/tesseract/wiki
) else (
    echo ✅ Chocolatey 설치 확인됨
    
    REM Tesseract 설치 확인
    tesseract --version >nul 2>&1
    if errorlevel 1 (
        echo 📦 Tesseract OCR을 설치합니다...
        choco install tesseract -y
        echo ✅ Tesseract OCR 설치 완료
    ) else (
        echo ✅ Tesseract OCR이 이미 설치되어 있습니다
    )
)

REM Java 확인 (tabula용)
java -version >nul 2>&1
if errorlevel 1 (
    echo 📦 Java를 설치합니다...
    choco install openjdk11 -y
    echo ✅ Java 설치 완료
) else (
    echo ✅ Java가 이미 설치되어 있습니다
)

echo.
echo 🎉 개발 환경 설정이 완료되었습니다!
echo.
echo 다음 명령어로 애플리케이션을 실행할 수 있습니다:
echo   .venv\Scripts\activate.bat
echo   python main.py
echo.
echo 실행 파일을 빌드하려면:
echo   python build.py
echo.
pause
