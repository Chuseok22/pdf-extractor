#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실제 PDF 파일로 추출 기능 테스트
"""

import os
import sys
import time
from pathlib import Path

# pdf_extractor 모듈 임포트
from pdf_extractor import PDFExtractor

def test_extraction():
    """실제 PDF 파일로 추출 테스트"""
    print("🧪 PDF 추출 기능 테스트 시작")
    
    # 파일 경로 설정
    pdf_path = "test/example.pdf"
    output_dir = "output"
    
    if not os.path.exists(pdf_path):
        print(f"❌ PDF 파일을 찾을 수 없습니다: {pdf_path}")
        return False
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"📁 출력 폴더 생성: {output_dir}")
    
    # 추출 옵션 설정
    options = {
        'min_rows': 2,
        'min_cols': 2,
        'separate_sheets': True,
        'include_text': True,
        'max_pages': None,  # 모든 페이지 (0 대신 None 사용)
        'batch_size': 5
    }
    
    print(f"📄 처리할 파일: {pdf_path}")
    print(f"📂 출력 폴더: {output_dir}")
    print(f"⚙️  추출 옵션: {options}")
    
    # PDFExtractor 인스턴스 생성
    extractor = PDFExtractor()
    
    # 추출 시작
    start_time = time.time()
    print(f"\n🚀 추출 시작...")
    
    try:
        success = extractor.extract_to_excel(pdf_path, output_dir, options)
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        if success:
            print(f"✅ 추출 완료! ({elapsed_time:.2f}초)")
            
            # 결과 파일 확인
            output_files = list(Path(output_dir).glob("*.xlsx"))
            print(f"📊 생성된 파일:")
            for file in output_files:
                file_size = file.stat().st_size
                print(f"  - {file.name} ({file_size} bytes)")
            
            return True
        else:
            print(f"❌ 추출 실패 ({elapsed_time:.2f}초)")
            return False
            
    except Exception as e:
        print(f"💥 오류 발생: {str(e)}")
        import traceback
        print(traceback.format_exc())
        return False

if __name__ == "__main__":
    success = test_extraction()
    if success:
        print("\n🎉 테스트 성공!")
    else:
        print("\n😞 테스트 실패!")
        sys.exit(1)
