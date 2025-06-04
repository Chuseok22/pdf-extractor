#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
대용량 PDF 시뮬레이션 테스트
실제 대용량 PDF가 없을 때 여러 파일로 테스트
"""

import os
import sys
import time
import shutil
from pathlib import Path

def create_test_scenarios():
    """다양한 테스트 시나리오 생성"""
    
    # 테스트 디렉토리 생성
    test_scenarios_dir = Path("test_scenarios")
    test_scenarios_dir.mkdir(exist_ok=True)
    
    # 기존 example.pdf를 여러 개 복사하여 대용량 시뮬레이션
    source_pdf = Path("test/example.pdf")
    if not source_pdf.exists():
        print("❌ 원본 PDF 파일이 없습니다: test/example.pdf")
        return False
    
    scenarios = [
        {"name": "single_file", "count": 1, "desc": "단일 파일 테스트"},
        {"name": "small_batch", "count": 3, "desc": "소규모 배치 (3개 파일)"},
        {"name": "medium_batch", "count": 5, "desc": "중간 배치 (5개 파일)"},
        {"name": "large_batch", "count": 10, "desc": "대규모 배치 (10개 파일)"}
    ]
    
    for scenario in scenarios:
        scenario_dir = test_scenarios_dir / scenario["name"]
        scenario_dir.mkdir(exist_ok=True)
        
        print(f"📁 {scenario['desc']} 시나리오 생성...")
        
        for i in range(scenario["count"]):
            target_file = scenario_dir / f"document_{i+1:02d}.pdf"
            shutil.copy2(source_pdf, target_file)
        
        print(f"   ✅ {scenario['count']}개 파일 생성 완료")
    
    return True

def test_scenario_performance():
    """각 시나리오별 성능 테스트"""
    from pdf_extractor import PDFExtractor
    
    scenarios_dir = Path("test_scenarios")
    if not scenarios_dir.exists():
        print("❌ 테스트 시나리오가 없습니다. create_test_scenarios()를 먼저 실행하세요.")
        return
    
    output_dir = Path("output_performance")
    output_dir.mkdir(exist_ok=True)
    
    extractor = PDFExtractor()
    
    options = {
        'min_rows': 2,
        'min_cols': 2,
        'separate_sheets': True,
        'include_text': True,
        'max_pages': 500,
        'batch_size': 5
    }
    
    results = []
    
    for scenario_dir in sorted(scenarios_dir.iterdir()):
        if not scenario_dir.is_dir():
            continue
            
        pdf_files = list(scenario_dir.glob("*.pdf"))
        if not pdf_files:
            continue
        
        print(f"\n🧪 {scenario_dir.name} 시나리오 테스트")
        print(f"📄 파일 수: {len(pdf_files)}")
        
        start_time = time.time()
        successful_files = 0
        
        for pdf_file in pdf_files:
            try:
                success = extractor.extract_to_excel(
                    str(pdf_file), 
                    str(output_dir), 
                    options
                )
                if success:
                    successful_files += 1
                    print(f"   ✅ {pdf_file.name}")
                else:
                    print(f"   ❌ {pdf_file.name}")
            except Exception as e:
                print(f"   💥 {pdf_file.name}: {str(e)}")
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        result = {
            'scenario': scenario_dir.name,
            'total_files': len(pdf_files),
            'successful_files': successful_files,
            'elapsed_time': elapsed_time,
            'files_per_second': successful_files / elapsed_time if elapsed_time > 0 else 0
        }
        results.append(result)
        
        print(f"   📊 결과: {successful_files}/{len(pdf_files)} 성공")
        print(f"   ⏱️  소요 시간: {elapsed_time:.2f}초")
        print(f"   🚀 처리 속도: {result['files_per_second']:.2f} 파일/초")
    
    # 성능 요약
    print(f"\n🎯 성능 테스트 요약:")
    print(f"{'시나리오':<15} {'파일수':<8} {'성공':<8} {'시간':<10} {'속도':<12}")
    print("-" * 60)
    
    for result in results:
        print(f"{result['scenario']:<15} "
              f"{result['total_files']:<8} "
              f"{result['successful_files']:<8} "
              f"{result['elapsed_time']:.2f}초{'':<6} "
              f"{result['files_per_second']:.2f} 파일/초")

if __name__ == "__main__":
    print("🔥 PDF Extractor 성능 테스트 도구")
    print("=" * 50)
    
    if len(sys.argv) > 1 and sys.argv[1] == "create":
        print("\n1️⃣ 테스트 시나리오 생성 중...")
        if create_test_scenarios():
            print("✅ 테스트 시나리오 생성 완료!")
        else:
            print("❌ 테스트 시나리오 생성 실패!")
            sys.exit(1)
    else:
        print("\n2️⃣ 성능 테스트 실행 중...")
        test_scenario_performance()
        print("\n🎉 성능 테스트 완료!")
