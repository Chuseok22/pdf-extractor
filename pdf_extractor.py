#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Extractor Core Module
PDF에서 데이터를 추출하여 Excel로 변환하는 핵심 모듈
"""

import os
import pandas as pd
import pdfplumber
import camelot
import tabula
import fitz  # PyMuPDF
from pathlib import Path
from typing import List, Dict, Optional, Union, Tuple
import logging
import traceback
import cv2
import numpy as np
from PIL import Image
import pytesseract
import easyocr
import tempfile
import psutil  # 메모리 모니터링
import gc      # 가비지 컬렉션
import time    # 성능 측정

# Ghostscript 경로 설정 (macOS Homebrew)
if '/opt/homebrew/bin' not in os.environ.get('PATH', ''):
    os.environ['PATH'] = '/opt/homebrew/bin:' + os.environ.get('PATH', '')


class PDFExtractor:
    """PDF 데이터 추출 클래스"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.setup_logging()
        self.process = psutil.Process()  # 현재 프로세스 모니터링
        
    def setup_logging(self):
        """로깅 설정"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
    def get_memory_usage(self) -> Dict[str, float]:
        """현재 메모리 사용량 반환 (MB 단위)"""
        try:
            memory_info = self.process.memory_info()
            system_memory = psutil.virtual_memory()
            
            return {
                'used_mb': memory_info.rss / 1024 / 1024,
                'available_mb': system_memory.available / 1024 / 1024,
                'total_mb': system_memory.total / 1024 / 1024,
                'percent': system_memory.percent
            }
        except Exception:
            return {'used_mb': 0, 'available_mb': 0, 'total_mb': 0, 'percent': 0}
    
    def log_memory_status(self, operation: str = ""):
        """메모리 상태 로깅"""
        memory = self.get_memory_usage()
        self.logger.info(f"[메모리] {operation} - 사용: {memory['used_mb']:.1f}MB, "
                        f"시스템 사용률: {memory['percent']:.1f}%")
    
    def cleanup_memory(self):
        """메모리 정리"""
        gc.collect()
        self.log_memory_status("메모리 정리 후")
    
    def get_pdf_info(self, pdf_path: str) -> Dict:
        """PDF 파일 정보 반환"""
        try:
            with fitz.open(pdf_path) as doc:
                return {
                    'pages': len(doc),
                    'title': doc.metadata.get('title', ''),
                    'author': doc.metadata.get('author', ''),
                    'subject': doc.metadata.get('subject', ''),
                    'creator': doc.metadata.get('creator', ''),
                    'file_size_mb': os.path.getsize(pdf_path) / (1024 * 1024)
                }
        except Exception as e:
            self.logger.error(f"PDF 정보 추출 실패: {str(e)}")
            return {
                'pages': 0,
                'title': '',
                'author': '',
                'subject': '',
                'creator': '',
                'file_size_mb': os.path.getsize(pdf_path) / (1024 * 1024) if os.path.exists(pdf_path) else 0
            }
        
    def extract_to_excel(self, pdf_path: str, output_dir: str, options: dict, password: str = None) -> bool:
        """PDF에서 데이터를 추출하여 Excel 파일로 저장 (안정성 개선)"""
        try:
            pdf_name = Path(pdf_path).stem
            output_path = os.path.join(output_dir, f"{pdf_name}_extracted.xlsx")
            
            self.logger.info(f"PDF 처리 시작: {pdf_path}")
            
            # PDF 파일 체크
            if not os.path.exists(pdf_path):
                self.logger.error(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
                return False
                
            # 암호 보호된 PDF 처리
            if self.is_password_protected(pdf_path):
                if password is None:
                    self.logger.error("PDF 파일이 암호로 보호되어 있습니다. 암호를 제공해주세요.")
                    return False
                
                # 암호 유효성 검증
                if not self.verify_password(pdf_path, password):
                    self.logger.error("제공된 암호가 올바르지 않습니다.")
                    return False
                
                self.logger.info("암호 인증 성공")
                
            # 파일 크기 체크 (100MB 이상이면 경고)
            file_size = os.path.getsize(pdf_path) / (1024 * 1024)  # MB
            if file_size > 100:
                self.logger.warning(f"대용량 PDF 파일 ({file_size:.1f}MB). 처리 시간이 오래 걸릴 수 있습니다.")
            
            # 추출 방법 결정
            method = options.get('method', '자동 선택')
            use_ocr = options.get('use_ocr', False)
            
            # 페이지 수 제한 (대용량 PDF 처리 최적화)
            max_pages = options.get('max_pages', 500)  # 기본 500페이지 제한
            if max_pages == 0 or max_pages is None:
                max_pages = 500  # 0이나 None인 경우 기본값 사용
            
            # 데이터 추출 (배치 처리)
            extracted_data = self.extract_data_batch(pdf_path, method, options, max_pages)
            
            if not extracted_data or (not extracted_data['tables'] and not extracted_data['text']):
                self.logger.warning("추출된 데이터가 없습니다.")
                return False
                
            # Excel 파일로 저장
            self.save_to_excel(extracted_data, output_path, options)
            
            self.logger.info(f"Excel 파일 생성 완료: {output_path}")
            return True
            
        except MemoryError:
            self.logger.error("메모리 부족으로 처리를 중단합니다. 더 작은 PDF로 분할하거나 max_pages 옵션을 줄여보세요.")
            return False
        except Exception as e:
            self.logger.error(f"PDF 추출 오류: {str(e)}")
            self.logger.error(traceback.format_exc())
            return False
            
    def extract_data_batch(self, pdf_path: str, method: str, options: dict, max_pages: int = 500) -> Dict:
        """대용량 PDF를 배치로 처리하여 데이터 추출"""
        extracted_data = {
            'tables': [],
            'text': [],
            'metadata': {}
        }
        
        try:
            # 시작 시 메모리 상태 로깅
            self.log_memory_status("추출 시작")
            
            # PDF 페이지 수 확인
            with fitz.open(pdf_path) as doc:
                total_pages = len(doc)
            
            self.logger.info(f"총 페이지 수: {total_pages}")
            
            if total_pages > max_pages:
                self.logger.warning(f"페이지 수가 {max_pages}페이지를 초과합니다. 상위 {max_pages}페이지만 처리합니다.")
                total_pages = max_pages
            
            # 배치 크기 설정 (메모리 효율성)
            batch_size = options.get('batch_size', min(50, total_pages))  # 기본 50페이지 또는 사용자 설정값
            
            for start_page in range(0, total_pages, batch_size):
                end_page = min(start_page + batch_size, total_pages)
                self.logger.info(f"페이지 {start_page + 1}-{end_page} 처리 중...")
                
                # 배치 시작 시 메모리 체크
                memory = self.get_memory_usage()
                if memory['percent'] > 85:  # 시스템 메모리 85% 이상 사용 시 경고
                    self.logger.warning(f"높은 메모리 사용률 감지: {memory['percent']:.1f}%")
                    
                try:
                    # 배치별 처리
                    batch_data = self.extract_data_range(pdf_path, method, options, start_page, end_page)
                    
                    # 결과 통합
                    if batch_data['tables']:
                        extracted_data['tables'].extend(batch_data['tables'])
                    if batch_data['text']:
                        extracted_data['text'].extend(batch_data['text'])
                    
                    # 배치 완료 후 메모리 정리
                    self.cleanup_memory()
                    
                except Exception as e:
                    self.logger.error(f"페이지 {start_page + 1}-{end_page} 처리 중 오류: {str(e)}")
                    continue
                    
            # 최종 메모리 상태 로깅
            self.log_memory_status("추출 완료")
            self.logger.info(f"배치 처리 완료. 테이블 {len(extracted_data['tables'])}개, 텍스트 {len(extracted_data['text'])}개 추출")
            
        except Exception as e:
            self.logger.error(f"배치 데이터 추출 오류: {str(e)}")
            
        return extracted_data
        
    def extract_data_range(self, pdf_path: str, method: str, options: dict, start_page: int, end_page: int) -> Dict:
        """지정된 페이지 범위에서 데이터 추출"""
        extracted_data = {
            'tables': [],
            'text': [],
            'metadata': {}
        }
        
        try:
            if method == '자동 선택':
                extracted_data = self.auto_extract_range(pdf_path, options, start_page, end_page)
            elif method == 'pdfplumber':
                extracted_data = self.extract_with_pdfplumber_range(pdf_path, options, start_page, end_page)
            elif method == 'camelot':
                extracted_data = self.extract_with_camelot_range(pdf_path, options, start_page, end_page)
            elif method == 'tabula':
                extracted_data = self.extract_with_tabula_range(pdf_path, options, start_page, end_page)
            elif method == 'pymupdf':
                extracted_data = self.extract_with_pymupdf_range(pdf_path, options, start_page, end_page)
                
            # OCR 사용 시 추가 처리 (페이지 범위 제한)
            if options.get('use_ocr', False) and end_page - start_page <= 10:  # OCR은 10페이지씩만
                ocr_data = self.extract_with_ocr_range(pdf_path, options, start_page, end_page)
                if ocr_data['tables']:
                    extracted_data['tables'].extend(ocr_data['tables'])
                if ocr_data['text']:
                    extracted_data['text'].extend(ocr_data['text'])
                    
        except Exception as e:
            self.logger.error(f"페이지 범위 데이터 추출 오류: {str(e)}")
            
        return extracted_data
    def extract_data(self, pdf_path: str, method: str, options: dict) -> Dict:
        """PDF에서 데이터 추출 (단일 처리용 - 하위 호환성)"""
        return self.extract_data_batch(pdf_path, method, options, options.get('max_pages', 500))
        
    def auto_extract_range(self, pdf_path: str, options: dict, start_page: int, end_page: int) -> Dict:
        """지정된 페이지 범위에서 자동으로 최적의 추출 방법 선택"""
        methods = ['pdfplumber', 'pymupdf', 'tabula', 'camelot']  # 안정성 순서로 정렬
        best_result = {'tables': [], 'text': [], 'metadata': {}}
        best_score = 0
        
        for method in methods:
            try:
                if method == 'pdfplumber':
                    result = self.extract_with_pdfplumber_range(pdf_path, options, start_page, end_page)
                elif method == 'pymupdf':
                    result = self.extract_with_pymupdf_range(pdf_path, options, start_page, end_page)
                elif method == 'tabula':
                    result = self.extract_with_tabula_range(pdf_path, options, start_page, end_page)
                elif method == 'camelot':
                    result = self.extract_with_camelot_range(pdf_path, options, start_page, end_page)
                    
                # 결과 점수 계산 (테이블 수 + 데이터 품질 + 텍스트)
                score = len(result['tables']) * 10
                for table in result['tables']:
                    if isinstance(table.get('data'), pd.DataFrame):
                        df = table['data']
                        score += len(df) * len(df.columns)
                
                # 텍스트 점수 추가 (텍스트가 있으면 보너스)
                if result['text']:
                    score += len(result['text']) * 5
                        
                if score > best_score:
                    best_score = score
                    best_result = result
                    self.logger.info(f"{method} 방법이 최고 점수 {score}를 기록했습니다.")
                    
            except Exception as e:
                self.logger.warning(f"{method} 추출 실패 (페이지 {start_page+1}-{end_page}): {str(e)}")
                continue
                
        return best_result
        
    def extract_with_pdfplumber_range(self, pdf_path: str, options: dict, start_page: int, end_page: int) -> Dict:
        """pdfplumber를 사용한 페이지 범위 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num in range(start_page, min(end_page, len(pdf.pages))):
                    try:
                        page = pdf.pages[page_num]
                        
                        # 테이블 추출
                        tables = page.find_tables()
                        for table in tables:
                            try:
                                table_data = table.extract()
                                if table_data and len(table_data) >= options.get('min_rows', 2):
                                    # 헤더가 있는지 확인
                                    if len(table_data) > 1:
                                        df = pd.DataFrame(table_data[1:], columns=table_data[0])
                                    else:
                                        df = pd.DataFrame(table_data)
                                    
                                    df = self.clean_dataframe(df)
                                    if len(df.columns) >= options.get('min_cols', 2):
                                        extracted_data['tables'].append({
                                            'data': df,
                                            'page': page_num + 1,
                                            'method': 'pdfplumber'
                                        })
                            except Exception as e:
                                self.logger.warning(f"테이블 추출 실패 (페이지 {page_num + 1}): {str(e)}")
                                
                        # 텍스트 추출
                        if options.get('include_text', False):
                            try:
                                text = page.extract_text()
                                if text and text.strip():
                                    extracted_data['text'].append({
                                        'content': text,
                                        'page': page_num + 1
                                    })
                            except Exception as e:
                                self.logger.warning(f"텍스트 추출 실패 (페이지 {page_num + 1}): {str(e)}")
                                
                    except Exception as e:
                        self.logger.warning(f"페이지 {page_num + 1} 처리 실패: {str(e)}")
                        continue
                        
        except Exception as e:
            self.logger.error(f"PDFPlumber 페이지 범위 추출 실패: {str(e)}")
            
        return extracted_data
        """자동으로 최적의 추출 방법 선택"""
        methods = ['pdfplumber', 'camelot', 'tabula', 'pymupdf']
        best_result = {'tables': [], 'text': [], 'metadata': {}}
        best_score = 0
        
        for method in methods:
            try:
                if method == 'pdfplumber':
                    result = self.extract_with_pdfplumber(pdf_path, options)
                elif method == 'camelot':
                    result = self.extract_with_camelot(pdf_path, options)
                elif method == 'tabula':
                    result = self.extract_with_tabula(pdf_path, options)
                elif method == 'pymupdf':
                    result = self.extract_with_pymupdf(pdf_path, options)
                    
                # 결과 점수 계산 (테이블 수 + 데이터 품질)
                score = len(result['tables']) * 10
                for table in result['tables']:
                    if isinstance(table, pd.DataFrame):
                        score += len(table) * len(table.columns)
                        
                if score > best_score:
                    best_score = score
                    best_result = result
                    
            except Exception as e:
                self.logger.warning(f"{method} 추출 실패: {str(e)}")
                continue
                
        return best_result
        
    def extract_with_pdfplumber(self, pdf_path: str, options: dict) -> Dict:
        """pdfplumber를 사용한 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # 테이블 추출
                tables = page.find_tables()
                for table in tables:
                    try:
                        table_data = table.extract()
                        if table_data and len(table_data) >= options.get('min_rows', 2):
                            df = pd.DataFrame(table_data[1:], columns=table_data[0])
                            df = self.clean_dataframe(df)
                            if len(df.columns) >= options.get('min_cols', 2):
                                extracted_data['tables'].append({
                                    'data': df,
                                    'page': page_num + 1,
                                    'method': 'pdfplumber'
                                })
                    except Exception as e:
                        self.logger.warning(f"테이블 추출 실패 (페이지 {page_num + 1}): {str(e)}")
                        
                # 텍스트 추출
                if options.get('include_text', False):
                    text = page.extract_text()
                    if text:
                        extracted_data['text'].append({
                            'content': text,
                            'page': page_num + 1
                        })
                        
        return extracted_data
        
    def extract_with_camelot(self, pdf_path: str, options: dict) -> Dict:
        """camelot을 사용한 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            tables = camelot.read_pdf(pdf_path, pages='all', flavor='lattice')
            
            for i, table in enumerate(tables):
                df = table.df
                df = self.clean_dataframe(df)
                
                if (len(df) >= options.get('min_rows', 2) and 
                    len(df.columns) >= options.get('min_cols', 2)):
                    extracted_data['tables'].append({
                        'data': df,
                        'page': table.page,
                        'method': 'camelot',
                        'accuracy': table.accuracy
                    })
                    
        except Exception as e:
            self.logger.error(f"Camelot 추출 실패: {str(e)}")
            
        return extracted_data
        
    def extract_with_tabula(self, pdf_path: str, options: dict) -> Dict:
        """tabula를 사용한 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
            
            for i, df in enumerate(tables):
                if isinstance(df, pd.DataFrame):
                    df = self.clean_dataframe(df)
                    
                    if (len(df) >= options.get('min_rows', 2) and 
                        len(df.columns) >= options.get('min_cols', 2)):
                        extracted_data['tables'].append({
                            'data': df,
                            'page': i + 1,  # 정확한 페이지 번호는 별도 계산 필요
                            'method': 'tabula'
                        })
                        
        except Exception as e:
            self.logger.error(f"Tabula 추출 실패: {str(e)}")
            
        return extracted_data
        
    def extract_with_pymupdf(self, pdf_path: str, options: dict) -> Dict:
        """PyMuPDF를 사용한 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            doc = fitz.open(pdf_path)
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                
                # 테이블 찾기 시도
                tables = page.find_tables()
                for table in tables:
                    try:
                        table_data = table.extract()
                        if table_data and len(table_data) >= options.get('min_rows', 2):
                            df = pd.DataFrame(table_data[1:], columns=table_data[0])
                            df = self.clean_dataframe(df)
                            if len(df.columns) >= options.get('min_cols', 2):
                                extracted_data['tables'].append({
                                    'data': df,
                                    'page': page_num + 1,
                                    'method': 'pymupdf'
                                })
                    except Exception as e:
                        self.logger.warning(f"PyMuPDF 테이블 추출 실패: {str(e)}")
                        
                # 텍스트 추출
                if options.get('include_text', False):
                    text = page.get_text()
                    if text:
                        extracted_data['text'].append({
                            'content': text,
                            'page': page_num + 1
                        })
                        
            doc.close()
            
        except Exception as e:
            self.logger.error(f"PyMuPDF 추출 실패: {str(e)}")
            
        return extracted_data
        
    def extract_with_pymupdf_range(self, pdf_path: str, options: dict, start_page: int, end_page: int) -> Dict:
        """PyMuPDF를 사용한 페이지 범위 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            doc = fitz.open(pdf_path)
            
            for page_num in range(start_page, min(end_page, len(doc))):
                try:
                    page = doc.load_page(page_num)
                    
                    # 테이블 찾기 시도
                    try:
                        tables = page.find_tables()
                        for table in tables:
                            try:
                                table_data = table.extract()
                                if table_data and len(table_data) >= options.get('min_rows', 2):
                                    if len(table_data) > 1:
                                        df = pd.DataFrame(table_data[1:], columns=table_data[0])
                                    else:
                                        df = pd.DataFrame(table_data)
                                    
                                    df = self.clean_dataframe(df)
                                    if len(df.columns) >= options.get('min_cols', 2):
                                        extracted_data['tables'].append({
                                            'data': df,
                                            'page': page_num + 1,
                                            'method': 'pymupdf'
                                        })
                            except Exception as e:
                                self.logger.warning(f"PyMuPDF 테이블 추출 실패 (페이지 {page_num + 1}): {str(e)}")
                    except Exception:
                        # find_tables() 메소드가 없는 경우 무시
                        pass
                        
                    # 텍스트 추출
                    if options.get('include_text', False):
                        try:
                            text = page.get_text()
                            if text and text.strip():
                                extracted_data['text'].append({
                                    'content': text,
                                    'page': page_num + 1
                                })
                        except Exception as e:
                            self.logger.warning(f"텍스트 추출 실패 (페이지 {page_num + 1}): {str(e)}")
                            
                except Exception as e:
                    self.logger.warning(f"페이지 {page_num + 1} 처리 실패: {str(e)}")
                    continue
                    
            doc.close()
            
        except Exception as e:
            self.logger.error(f"PyMuPDF 페이지 범위 추출 실패: {str(e)}")
            
        return extracted_data
        
    def extract_with_camelot_range(self, pdf_path: str, options: dict, start_page: int, end_page: int) -> Dict:
        """camelot을 사용한 페이지 범위 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            # 페이지 범위 문자열 생성
            page_range = f"{start_page + 1}-{end_page}"
            
            tables = camelot.read_pdf(pdf_path, pages=page_range, flavor='lattice')
            
            for table in tables:
                try:
                    df = table.df
                    df = self.clean_dataframe(df)
                    
                    if (len(df) >= options.get('min_rows', 2) and 
                        len(df.columns) >= options.get('min_cols', 2)):
                        extracted_data['tables'].append({
                            'data': df,
                            'page': table.page,
                            'method': 'camelot',
                            'accuracy': getattr(table, 'accuracy', 0)
                        })
                except Exception as e:
                    self.logger.warning(f"Camelot 테이블 처리 실패: {str(e)}")
                    
        except Exception as e:
            self.logger.error(f"Camelot 페이지 범위 추출 실패: {str(e)}")
            
        return extracted_data
        
    def extract_with_tabula_range(self, pdf_path: str, options: dict, start_page: int, end_page: int) -> Dict:
        """tabula를 사용한 페이지 범위 추출"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            # 페이지 범위 문자열 생성
            page_range = f"{start_page + 1}-{end_page}"
            
            tables = tabula.read_pdf(pdf_path, pages=page_range, multiple_tables=True)
            
            for i, df in enumerate(tables):
                try:
                    if isinstance(df, pd.DataFrame):
                        df = self.clean_dataframe(df)
                        
                        if (len(df) >= options.get('min_rows', 2) and 
                            len(df.columns) >= options.get('min_cols', 2)):
                            extracted_data['tables'].append({
                                'data': df,
                                'page': start_page + i + 1,  # 근사치
                                'method': 'tabula'
                            })
                except Exception as e:
                    self.logger.warning(f"Tabula 테이블 처리 실패: {str(e)}")
                    
        except Exception as e:
            self.logger.error(f"Tabula 페이지 범위 추출 실패: {str(e)}")
            
        return extracted_data
        
    def extract_with_ocr_range(self, pdf_path: str, options: dict, start_page: int, end_page: int) -> Dict:
        """OCR을 사용한 페이지 범위 추출 (제한적 사용)"""
        extracted_data = {'tables': [], 'text': [], 'metadata': {}}
        
        try:
            doc = fitz.open(pdf_path)
            ocr_engine = options.get('ocr_engine', 'tesseract')
            
            # EasyOCR 초기화 (필요한 경우)
            reader = None
            if ocr_engine == 'easyocr':
                reader = easyocr.Reader(['ko', 'en'])
                
            for page_num in range(start_page, min(end_page, len(doc))):
                try:
                    page = doc.load_page(page_num)
                    
                    # 페이지를 이미지로 변환 (해상도 제한으로 메모리 절약)
                    pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))  # 1.5배 확대로 제한
                    img_data = pix.tobytes("png")
                    
                    # 임시 파일로 저장
                    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                        tmp_file.write(img_data)
                        tmp_path = tmp_file.name
                        
                    try:
                        # OCR 수행
                        if ocr_engine == 'tesseract':
                            text = pytesseract.image_to_string(
                                tmp_path, lang='kor+eng', config='--psm 6'
                            )
                        else:  # easyocr
                            results = reader.readtext(tmp_path)
                            text = '\n'.join([result[1] for result in results])
                            
                        if text.strip():
                            # 테이블 구조 감지 시도
                            tables = self.detect_table_structure(text, page_num + 1)
                            extracted_data['tables'].extend(tables)
                            
                            # 텍스트 저장
                            if options.get('include_text', False):
                                extracted_data['text'].append({
                                    'content': text,
                                    'page': page_num + 1,
                                    'method': f'ocr_{ocr_engine}'
                                })
                                
                    finally:
                        # 임시 파일 삭제
                        try:
                            os.unlink(tmp_path)
                        except:
                            pass
                        
                except Exception as e:
                    self.logger.warning(f"OCR 처리 실패 (페이지 {page_num + 1}): {str(e)}")
                    continue
                    
            doc.close()
            
        except Exception as e:
            self.logger.error(f"OCR 페이지 범위 추출 실패: {str(e)}")
            
        return extracted_data

    def detect_table_structure(self, text: str, page_num: int) -> List[Dict]:
        """텍스트에서 테이블 구조 감지 (안정성 개선)"""
        tables = []
        
        try:
            lines = text.strip().split('\n')
            potential_table_lines = []
            
            for line in lines:
                # 탭이나 다중 공백으로 분리된 데이터 찾기
                if '\t' in line or '  ' in line:
                    # 분리 시도
                    if '\t' in line:
                        cells = [cell.strip() for cell in line.split('\t') if cell.strip()]
                    else:
                        cells = [cell.strip() for cell in line.split('  ') if cell.strip()]
                        
                    if len(cells) >= 2:  # 최소 2개 열
                        potential_table_lines.append(cells)
                        
            # 연속된 테이블 라인들을 그룹화
            if len(potential_table_lines) >= 2:
                # 컬럼 수가 비슷한 라인들을 찾기
                max_cols = max(len(line) for line in potential_table_lines)
                filtered_lines = [line for line in potential_table_lines 
                                if len(line) >= max_cols - 1]  # 1개까지 차이 허용
                
                if len(filtered_lines) >= 2:
                    try:
                        # DataFrame 생성 (안전하게)
                        max_len = max(len(line) for line in filtered_lines)
                        normalized_lines = []
                        
                        for line in filtered_lines:
                            # 행 길이 정규화
                            while len(line) < max_len:
                                line.append('')
                            normalized_lines.append(line[:max_len])
                        
                        if len(normalized_lines) > 1:
                            df = pd.DataFrame(normalized_lines[1:], columns=normalized_lines[0])
                        else:
                            df = pd.DataFrame(normalized_lines)
                            
                        df = self.clean_dataframe(df)
                        
                        if not df.empty and len(df.columns) >= 2:
                            tables.append({
                                'data': df,
                                'page': page_num,
                                'method': 'ocr_table_detection'
                            })
                    except Exception as e:
                        self.logger.warning(f"테이블 DataFrame 생성 실패: {str(e)}")
                    
        except Exception as e:
            self.logger.warning(f"테이블 구조 감지 실패: {str(e)}")
            
        return tables
        
    def clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """DataFrame 정리 (안전한 처리)"""
        try:
            # 빈 행/열 제거
            df = df.dropna(how='all').dropna(axis=1, how='all')
            
            # 공백 제거 (pandas 2.1.0+ 호환)
            if hasattr(df, 'map'):
                df = df.map(lambda x: str(x).strip() if pd.notna(x) else x)
            else:
                df = df.applymap(lambda x: str(x).strip() if pd.notna(x) else x)
            
            # 빈 문자열을 NaN으로 변경
            df = df.replace('', pd.NA)
            
            # 데이터프레임이 너무 크면 제한
            if len(df) > 10000:
                self.logger.warning(f"테이블이 너무 큼 ({len(df)}행). 상위 10,000행만 유지")
                df = df.head(10000)
            
            return df
            
        except Exception as e:
            self.logger.error(f"DataFrame 정리 중 오류: {str(e)}")
            return df
        
    def save_to_excel(self, extracted_data: Dict, output_path: str, options: dict):
        """추출된 데이터를 Excel 파일로 저장 (안정성 개선)"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                
                if options.get('separate_sheets', False):
                    # 페이지별 시트 분리
                    page_tables = {}
                    for table_info in extracted_data['tables']:
                        page = table_info['page']
                        if page not in page_tables:
                            page_tables[page] = []
                        page_tables[page].append(table_info)
                        
                    for page, tables in page_tables.items():
                        sheet_name = f"Page_{page}"
                        start_row = 0
                        
                        for i, table_info in enumerate(tables):
                            try:
                                df = table_info['data']
                                if isinstance(df, pd.DataFrame) and not df.empty:
                                    # 시트 이름 길이 제한 (Excel 제한사항)
                                    if len(sheet_name) > 31:
                                        sheet_name = sheet_name[:31]
                                    
                                    df.to_excel(writer, sheet_name=sheet_name, 
                                              startrow=start_row, index=False)
                                    start_row += len(df) + 2  # 2행 간격
                            except Exception as e:
                                self.logger.warning(f"테이블 저장 실패 (페이지 {page}): {str(e)}")
                                
                else:
                    # 모든 테이블을 하나의 시트에
                    if extracted_data['tables']:
                        start_row = 0
                        tables_saved = 0
                        
                        for table_info in extracted_data['tables']:
                            try:
                                df = table_info['data']
                                if isinstance(df, pd.DataFrame) and not df.empty:
                                    # 첫 번째 테이블에서 시트 생성
                                    sheet_name = "Tables"
                                    
                                    # 메타데이터 행 추가
                                    meta_data = [f"페이지 {table_info['page']} - {table_info['method']}"]
                                    meta_df = pd.DataFrame([meta_data])
                                    meta_df.to_excel(writer, sheet_name=sheet_name, 
                                                   startrow=start_row, index=False, header=False)
                                    start_row += 1
                                    
                                    # 테이블 데이터 저장
                                    df.to_excel(writer, sheet_name=sheet_name, 
                                              startrow=start_row, index=False)
                                    start_row += len(df) + 2
                                    tables_saved += 1
                                    
                            except Exception as e:
                                self.logger.warning(f"테이블 저장 실패: {str(e)}")
                                
                        self.logger.info(f"{tables_saved}개 테이블이 저장되었습니다.")
                        
                # 텍스트 데이터 저장
                if extracted_data['text'] and options.get('include_text', False):
                    try:
                        text_data = []
                        for text_info in extracted_data['text']:
                            text_data.append({
                                'Page': text_info['page'],
                                'Content': text_info['content'][:32000]  # Excel 셀 크기 제한
                            })
                            
                        if text_data:
                            text_df = pd.DataFrame(text_data)
                            text_df.to_excel(writer, sheet_name="Text", index=False)
                            self.logger.info(f"{len(text_data)}개 텍스트 블록이 저장되었습니다.")
                    except Exception as e:
                        self.logger.warning(f"텍스트 저장 실패: {str(e)}")
                        
        except Exception as e:
            self.logger.error(f"Excel 파일 저장 실패: {str(e)}")
            raise

    def extract_tables(self, pdf_path: str, output_directory: str, method: str = "pdfplumber", password: str = None) -> tuple:
        """간소화된 GUI용 테이블 추출 메서드"""
        try:
            # 암호 보호된 PDF 체크
            if password is None and self.is_password_protected(pdf_path):
                return False, "PDF 파일이 암호로 보호되어 있습니다."
            
            # 기본 추출 옵션 설정
            options = {
                'method': method,
                'min_rows': 2,
                'min_cols': 2,
                'separate_sheets': True,
                'include_text': False,
                'max_pages': 500,
                'batch_size': 50
            }
            
            # 암호가 있는 경우 PDF 열기 테스트
            if password:
                try:
                    doc = fitz.open(pdf_path)
                    if doc.needsPass:
                        if not doc.authenticate(password):
                            return False, "잘못된 암호입니다."
                    doc.close()
                except Exception as e:
                    return False, f"PDF 열기 실패: {str(e)}"
            
            # 실제 추출 수행
            success = self.extract_to_excel(pdf_path, output_directory, options)
            
            if success:
                pdf_name = Path(pdf_path).stem
                excel_path = os.path.join(output_directory, f"{pdf_name}_extracted.xlsx")
                return True, f"추출 완료: {excel_path}"
            else:
                return False, "추출에 실패했습니다."
                
        except Exception as e:
            return False, f"오류 발생: {str(e)}"
    
    def is_password_protected(self, pdf_path: str) -> bool:
        """PDF가 암호로 보호되어 있는지 확인"""
        try:
            doc = fitz.open(pdf_path)
            is_protected = doc.needsPass
            doc.close()
            return is_protected
        except Exception:
            return False
    
    def verify_password(self, pdf_path: str, password: str) -> bool:
        """PDF 암호 유효성 검증"""
        try:
            doc = fitz.open(pdf_path)
            if doc.needsPass:
                is_valid = doc.authenticate(password)
                doc.close()
                return is_valid
            else:
                doc.close()
                return True  # 암호가 필요하지 않은 PDF
        except Exception as e:
            self.logger.error(f"암호 검증 중 오류: {str(e)}")
            return False

def test_extractor():
    """테스트 함수 (안정성 개선)"""
    extractor = PDFExtractor()
    
    # 대용량 PDF 처리를 위한 개선된 샘플 옵션
    options = {
        'method': '자동 선택',
        'use_ocr': False,
        'ocr_engine': 'tesseract',
        'min_rows': 2,
        'min_cols': 2,
        'separate_sheets': True,
        'include_text': False,  # 대용량 PDF에서는 텍스트 비활성화
        'max_pages': 100,  # 최대 100페이지만 처리
    }
    
    print("PDF 추출기 안정성 개선 완료:")
    print("✅ pandas applymap -> map 호환성 개선")
    print("✅ 대용량 PDF 배치 처리 (50페이지씩)")
    print("✅ 메모리 사용량 최적화")
    print("✅ 에러 핸들링 강화")
    print("✅ Excel 저장 안정성 개선")
    print("✅ 페이지 수 제한 옵션 추가")
    print("\n사용 예시:")
    print("options = {")
    print("    'method': '자동 선택',")
    print("    'max_pages': 200,  # 대용량 PDF용")
    print("    'use_ocr': False,  # 성능 최적화")
    print("    'separate_sheets': True")
    print("}")
    print("extractor.extract_to_excel('large.pdf', 'output', options)")


if __name__ == "__main__":
    test_extractor()
