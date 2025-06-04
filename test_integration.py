#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
최종 통합 테스트 스크립트
간소화된 PDF Extractor의 모든 기능을 포괄적으로 테스트
"""

import sys
import os
from pathlib import Path

# 현재 디렉토리를 시스템 경로에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from pdf_extractor import PDFExtractor

def test_complete_workflow():
    """완전한 작업 흐름 테스트"""
    print("🧪 PDF Extractor 통합 테스트")
    print("=" * 60)
    
    extractor = PDFExtractor()
    test_files = []
    
    # 테스트 파일 확인
    test_dirs = [
        Path(current_dir) / "test",
        Path(current_dir) / "test_scenarios" / "single_file",
        Path(current_dir) / "test_scenarios" / "small_batch"
    ]
    
    for test_dir in test_dirs:
        if test_dir.exists():
            pdf_files = list(test_dir.glob("*.pdf"))
            test_files.extend(pdf_files)
    
    if not test_files:
        print("❌ 테스트용 PDF 파일이 없습니다.")
        return False
    
    print(f"📁 테스트 파일 {len(test_files)}개 발견")
    
    # 출력 디렉토리 생성
    output_dir = Path(current_dir) / "output" / "integration_test"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    success_count = 0
    total_count = 0
    
    for pdf_file in test_files[:3]:  # 최대 3개 파일만 테스트
        total_count += 1
        print(f"\n📄 테스트 {total_count}: {pdf_file.name}")
        print("-" * 40)
        
        try:
            # 1. 암호 보호 확인
            is_protected = extractor.is_password_protected(str(pdf_file))
            print(f"🔒 암호 보호: {'예' if is_protected else '아니오'}")
            
            if is_protected:
                print("⏭️  암호 보호 파일은 건너뜁니다")
                continue
            
            # 2. 추출 실행
            success, message = extractor.extract_tables(
                str(pdf_file), 
                str(output_dir),
                method="pdfplumber"
            )
            
            if success:
                print(f"✅ {message}")
                success_count += 1
                
                # 생성된 파일 확인
                pdf_name = pdf_file.stem
                excel_file = output_dir / f"{pdf_name}_extracted.xlsx"
                if excel_file.exists():
                    file_size = excel_file.stat().st_size
                    print(f"📊 Excel 파일: {excel_file.name} ({file_size} bytes)")
                    
            else:
                print(f"❌ {message}")
                
        except Exception as e:
            print(f"❌ 오류 발생: {e}")
    
    # 결과 요약
    print("\n" + "=" * 60)
    print("📊 테스트 결과 요약")
    print("=" * 60)
    print(f"✅ 성공: {success_count}/{total_count} 파일")
    print(f"📁 출력 디렉토리: {output_dir}")
    
    if success_count > 0:
        # 생성된 파일 목록
        excel_files = list(output_dir.glob("*.xlsx"))
        print(f"📄 생성된 Excel 파일: {len(excel_files)}개")
        for excel_file in excel_files:
            file_size = excel_file.stat().st_size
            print(f"   - {excel_file.name} ({file_size} bytes)")
    
    success_rate = (success_count / total_count) * 100 if total_count > 0 else 0
    print(f"📈 성공률: {success_rate:.1f}%")
    
    return success_count > 0

def test_gui_compatibility():
    """GUI 호환성 테스트"""
    print("\n🖥️  GUI 호환성 테스트")
    print("=" * 60)
    
    try:
        # Qt 환경 테스트
        from PyQt5.QtWidgets import QApplication
        app = QApplication.instance()
        if app is None:
            app = QApplication([])
        
        print("✅ PyQt5 GUI 환경 정상")
        
        # GUI 클래스 import 테스트
        from main import PDFExtractorGUI, PDFPasswordDialog
        print("✅ GUI 클래스 import 정상")
        
        # GUI 인스턴스 생성 테스트 (실제 화면 표시 없이)
        window = PDFExtractorGUI()
        print("✅ GUI 인스턴스 생성 정상")
        
        return True
        
    except Exception as e:
        print(f"❌ GUI 테스트 실패: {e}")
        return False

def main():
    """메인 테스트 함수"""
    print("🔍 PDF Extractor - 최종 통합 테스트")
    print("=" * 60)
    
    # 1. 추출 기능 테스트
    extraction_success = test_complete_workflow()
    
    # 2. GUI 호환성 테스트
    gui_success = test_gui_compatibility()
    
    # 최종 결과
    print("\n" + "🏁 최종 결과" + " " * 50)
    print("=" * 60)
    print(f"📊 추출 기능: {'✅ 정상' if extraction_success else '❌ 실패'}")
    print(f"🖥️  GUI 호환성: {'✅ 정상' if gui_success else '❌ 실패'}")
    
    overall_success = extraction_success and gui_success
    print(f"\n🎯 전체 상태: {'✅ 모든 테스트 통과' if overall_success else '❌ 일부 테스트 실패'}")
    
    if overall_success:
        print("\n💡 사용 준비 완료!")
        print("   GUI 실행: python main.py")
        print("   직접 사용: from pdf_extractor import PDFExtractor")
    
    return overall_success

if __name__ == "__main__":
    main()
