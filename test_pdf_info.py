#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF 파일 정보 분석 스크립트
"""

import fitz  # PyMuPDF
import os

def analyze_pdf(pdf_path):
    """PDF 파일 분석"""
    try:
        print(f"📄 PDF 파일 분석: {os.path.basename(pdf_path)}")
        print(f"📁 파일 크기: {os.path.getsize(pdf_path) / 1024:.1f} KB")
        
        doc = fitz.open(pdf_path)
        
        print(f"📖 총 페이지 수: {len(doc)}")
        print(f"📝 메타데이터:")
        metadata = doc.metadata
        for key, value in metadata.items():
            if value:
                print(f"   {key}: {value}")
        
        print(f"\n📊 각 페이지 정보:")
        for page_num in range(min(3, len(doc))):  # 처음 3페이지만 분석
            page = doc.load_page(page_num)
            text = page.get_text()
            text_blocks = page.get_text("dict")["blocks"]
            
            print(f"\n페이지 {page_num + 1}:")
            print(f"  - 텍스트 길이: {len(text)} 문자")
            print(f"  - 텍스트 블록 수: {len([b for b in text_blocks if 'lines' in b])}")
            
            # 텍스트 미리보기 (처음 100자)
            preview = text.strip()[:100].replace('\n', ' ')
            if preview:
                print(f"  - 미리보기: {preview}...")
        
        doc.close()
        return True
        
    except Exception as e:
        print(f"❌ 오류: {str(e)}")
        return False

if __name__ == "__main__":
    pdf_path = "test/example.pdf"
    if os.path.exists(pdf_path):
        analyze_pdf(pdf_path)
    else:
        print(f"❌ 파일을 찾을 수 없습니다: {pdf_path}")
