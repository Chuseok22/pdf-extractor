#!/bin/bash
# PDF Extractor 개발 환경 설정 스크립트 (macOS/Linux)

set -e

echo "🚀 PDF Extractor 개발 환경을 설정합니다..."

# Python 버전 확인
python_version=$(python3 --version 2>&1 | cut -d' ' -f2 | cut -d'.' -f1-2)
required_version="3.8"

if [ "$(printf '%s\n' "$required_version" "$python_version" | sort -V | head -n1)" != "$required_version" ]; then
    echo "❌ Python 3.8 이상이 필요합니다. 현재 버전: $python_version"
    exit 1
fi

echo "✅ Python 버전 확인: $python_version"

# 가상환경 생성
if [ ! -d ".venv" ]; then
    echo "📦 Python 가상환경을 생성합니다..."
    python3 -m venv .venv
    echo "✅ 가상환경 생성 완료"
else
    echo "✅ 가상환경이 이미 존재합니다"
fi

# 가상환경 활성화
echo "🔄 가상환경을 활성화합니다..."
source .venv/bin/activate

# pip 업그레이드
echo "📈 pip을 업그레이드합니다..."
pip install --upgrade pip

# 의존성 설치
echo "📦 의존성을 설치합니다..."
pip install -r requirements.txt

# 시스템 의존성 확인 및 설치
echo "🔧 시스템 의존성을 확인합니다..."

# macOS에서 Homebrew 확인
if [[ "$OSTYPE" == "darwin"* ]]; then
    if ! command -v brew &> /dev/null; then
        echo "❌ Homebrew가 설치되어 있지 않습니다."
        echo "다음 명령어로 Homebrew를 설치하세요:"
        echo '/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
        exit 1
    fi
    
    # Tesseract 설치 확인
    if ! command -v tesseract &> /dev/null; then
        echo "📦 Tesseract OCR을 설치합니다..."
        brew install tesseract
        echo "✅ Tesseract OCR 설치 완료"
    else
        echo "✅ Tesseract OCR이 이미 설치되어 있습니다"
    fi
    
    # Java 설치 확인 (tabula용)
    if ! command -v java &> /dev/null; then
        echo "📦 Java를 설치합니다..."
        brew install openjdk@11
        echo "✅ Java 설치 완료"
    else
        echo "✅ Java가 이미 설치되어 있습니다"
    fi
fi

# 프로젝트 구조 확인
echo "📁 프로젝트 구조를 확인합니다..."
required_files=("main.py" "pdf_extractor.py" "build.py" "requirements.txt")

for file in "${required_files[@]}"; do
    if [ -f "$file" ]; then
        echo "✅ $file"
    else
        echo "❌ $file이 누락되었습니다"
    fi
done

echo ""
echo "🎉 개발 환경 설정이 완료되었습니다!"
echo ""
echo "다음 명령어로 애플리케이션을 실행할 수 있습니다:"
echo "  source .venv/bin/activate"
echo "  python main.py"
echo ""
echo "실행 파일을 빌드하려면:"
echo "  python build.py"
