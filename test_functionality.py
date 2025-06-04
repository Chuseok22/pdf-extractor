#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
간소화된 PDF Extractor 기능 테스트
"""

import sys
import os
from pathlib import Path

# 현재 디렉토리를 시스템 경로에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from pdf_extractor import PDFExtractor

def test_extractor_initialization():
    """PDFExtractor 초기화 테스트"""
    print("=" * 50)
    print("PDFExtractor 초기화 테스트")
    print("=" * 50)
    
    try:
        extractor = PDFExtractor()
        print("✅ PDFExtractor 초기화 성공")
        return extractor
    except Exception as e:
        print(f"❌ PDFExtractor 초기화 실패: {e}")
        return None

def test_methods_availability(extractor):
    """메서드 사용 가능성 테스트"""
    print("\n" + "=" * 50)
    print("메서드 사용 가능성 테스트")
    print("=" * 50)
    
    methods_to_check = [
        'extract_tables',
        'is_password_protected',
        'extract_to_excel'
    ]
    
    for method_name in methods_to_check:
        if hasattr(extractor, method_name):
            print(f"✅ {method_name} 메서드 사용 가능")
        else:
            print(f"❌ {method_name} 메서드 없음")

def test_with_sample_pdf():
    """샘플 PDF로 실제 추출 테스트"""
    print("\n" + "=" * 50)
    print("샘플 PDF 추출 테스트")
    print("=" * 50)
    
    # test 폴더의 example.pdf 사용
    sample_pdf = Path(current_dir) / "test" / "example.pdf"
    
    if not sample_pdf.exists():
        print("❌ 샘플 PDF 파일이 없습니다 (test/example.pdf)")
        return
    
    print(f"📄 테스트 파일: {sample_pdf.name}")
    
    try:
        extractor = PDFExtractor()
        
        # 암호 보호 확인
        is_protected = extractor.is_password_protected(str(sample_pdf))
        print(f"🔒 암호 보호: {'예' if is_protected else '아니오'}")
        
        # 테이블 추출 시도 (output 디렉토리 포함)
        print("📊 테이블 추출 시도중...")
        output_dir = Path(current_dir) / "output"
        output_dir.mkdir(exist_ok=True)
        
        success, message = extractor.extract_tables(str(sample_pdf), str(output_dir))
        
        if success:
            print(f"✅ 추출 성공: {message}")
            
            # 생성된 Excel 파일 확인
            excel_files = list(output_dir.glob("*.xlsx"))
            if excel_files:
                print(f"📊 생성된 파일: {len(excel_files)}개")
                for excel_file in excel_files:
                    file_size = excel_file.stat().st_size
                    print(f"   - {excel_file.name} ({file_size} bytes)")
        else:
            print(f"❌ 추출 실패: {message}")
            
    except Exception as e:
        print(f"❌ 추출 테스트 실패: {e}")

def main():
    """메인 테스트 함수"""
    print("🔍 PDF Extractor 기능 테스트 시작")
    
    # 1. PDFExtractor 초기화 테스트
    extractor = test_extractor_initialization()
    if not extractor:
        return
    
    # 2. 메서드 가용성 테스트
    test_methods_availability(extractor)
    
    # 3. 샘플 PDF로 실제 테스트
    test_with_sample_pdf()
    
    print("\n" + "=" * 50)
    print("✅ 기능 테스트 완료")
    print("=" * 50)
    print("\n💡 GUI 테스트:")
    print("   python main.py 명령으로 GUI를 실행하여 실제 추출 기능을 테스트해보세요")

if __name__ == "__main__":
    main()