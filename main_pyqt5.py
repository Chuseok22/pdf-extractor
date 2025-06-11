#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Excel Extractor - Simplified Version
PDF 파일에서 데이터를 추출하여 Excel 파일로 저장하는 간단한 GUI 애플리케이션
"""

import sys
import os
import traceback
import json
import time
import psutil
from pathlib import Path
from typing import List

# Qt 환경 설정 (PyQt5 GUI 문제 해결)
def setup_qt_environment():
    """Qt 환경 변수 설정으로 GUI 플러그인 문제 해결"""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    venv_path = os.path.join(current_dir, '.venv', 'lib', 'python3.13', 'site-packages', 'PyQt5')
    
    qt_lib_path = os.path.join(venv_path, 'Qt5', 'lib')
    qt_plugins_path = os.path.join(venv_path, 'Qt5', 'plugins')
    
    # Qt 플러그인 경로 설정
    os.environ['QT_PLUGIN_PATH'] = qt_plugins_path
    os.environ['DYLD_LIBRARY_PATH'] = qt_lib_path
    os.environ['DYLD_FRAMEWORK_PATH'] = qt_lib_path

# Qt 환경 초기화
setup_qt_environment()

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
    QWidget, QPushButton, QLabel, QTextEdit, QFileDialog,
    QProgressBar, QMessageBox, QInputDialog, QTabWidget,
    QGroupBox, QGridLayout, QComboBox, QSpinBox, QCheckBox,
    QDialog, QDialogButtonBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QTimer, QTime
from PyQt5.QtGui import QFont, QPalette, QColor

from pdf_extractor import PDFExtractor

# 테마 색상 정의 (라이트 모드만 사용)
THEME = {
    'background': '#FFFFFF',
    'surface': '#FAFAFA',
    'text': '#212121',
    'primary': '#4CAF50',
    'secondary': '#2196F3',
    'accent': '#FF9800',
    'border': '#E0E0E0',
    'hover': '#F5F5F5'
}


class ExtractorThread(QThread):
    """PDF 추출 작업을 백그라운드에서 수행하는 스레드"""
    progress = pyqtSignal(int)  # 전체 진행률 (0-100)
    file_progress = pyqtSignal(int)  # 현재 파일 진행률 (0-100)
    status = pyqtSignal(str)  # 상태 메시지
    current_file = pyqtSignal(str)  # 현재 처리 중인 파일명
    progress_details = pyqtSignal(str)  # 상세 진행 정보 (페이지 등)
    log_message = pyqtSignal(str)  # 로그 메시지
    processing_stats = pyqtSignal(int, int, float)  # 처리된 파일 수, 전체 파일 수, 경과 시간
    finished = pyqtSignal(bool, str)
    
    def __init__(self, pdf_files: List[str], output_dir: str, options: dict):
        super().__init__()
        self.pdf_files = pdf_files
        self.output_dir = output_dir
        self.options = options
        self.extractor = PDFExtractor()
        
    def run(self):
        try:
            import time
            start_time = time.time()
            
            total_files = len(self.pdf_files)
            self.status.emit("추출 작업 시작...")
            self.log_message.emit(f"📁 총 {total_files}개 파일 처리 시작")
            
            for i, pdf_file in enumerate(self.pdf_files):
                filename = os.path.basename(pdf_file)
                self.current_file.emit(f"현재 파일: {filename}")
                
                # 파일 정보 표시
                try:
                    file_size = os.path.getsize(pdf_file) / (1024 * 1024)  # MB
                    size_info = f"({file_size:.1f}MB)" if file_size > 1 else f"({file_size*1024:.0f}KB)"
                    self.log_message.emit(f"📄 처리 시작: {filename} {size_info}")
                    
                    if file_size > 50:
                        self.status.emit(f"대용량 파일 처리 중...")
                        self.log_message.emit(f"⚠️  대용량 파일 감지: {file_size:.1f}MB")
                    else:
                        self.status.emit(f"파일 처리 중...")
                except Exception as e:
                    self.log_message.emit(f"❌ 파일 정보 읽기 실패: {str(e)}")
                
                # 파일 진행률 초기화
                self.file_progress.emit(0)
                
                # PDF 페이지 수 확인 (간단한 체크)
                try:
                    import fitz  # PyMuPDF
                    doc = fitz.open(pdf_file)
                    total_pages = len(doc)
                    doc.close()
                    self.progress_details.emit(f"0/{total_pages} 페이지")
                    self.log_message.emit(f"📖 총 {total_pages}페이지 감지")
                except:
                    total_pages = 0
                    self.progress_details.emit("페이지 수 확인 불가")
                
                # PDF 추출 수행
                self.log_message.emit(f"🔄 데이터 추출 시작...")
                self.file_progress.emit(50)  # 중간 진행률
                
                success = self.extractor.extract_to_excel(
                    pdf_file, self.output_dir, self.options
                )
                
                if not success:
                    self.log_message.emit(f"❌ 실패: {filename}")
                    self.finished.emit(False, f"파일 처리 실패: {filename}")
                    return
                
                # 파일 완료
                self.file_progress.emit(100)
                if total_pages > 0:
                    self.progress_details.emit(f"{total_pages}/{total_pages} 페이지")
                
                # 전체 진행률 업데이트
                overall_progress = int((i + 1) / total_files * 100)
                self.progress.emit(overall_progress)
                
                elapsed_time = time.time() - start_time
                self.log_message.emit(f"✅ 완료: {filename} ({elapsed_time:.1f}초)")
                self.status.emit(f"완료: {i+1}/{total_files} 파일")
                
                # 처리 통계 업데이트
                self.processing_stats.emit(i + 1, total_files, elapsed_time)
            
            total_time = time.time() - start_time
            self.status.emit("모든 파일 처리 완료!")
            self.log_message.emit(f"🎉 전체 작업 완료! (총 {total_time:.1f}초)")
            self.finished.emit(True, f"총 {total_files}개 파일이 성공적으로 처리되었습니다.")
            
        except Exception as e:
            self.log_message.emit(f"💥 오류 발생: {str(e)}")
            self.finished.emit(False, f"오류 발생: {str(e)}")


class PDFExtractorGUI(QMainWindow):
    """PDF to Excel Extractor GUI 메인 클래스"""
    
    def __init__(self):
        super().__init__()
        self.pdf_files = []
        self.output_directory = ""
        self.init_ui()
        self.load_settings()
        
    def init_ui(self):
        """UI 초기화"""
        self.setWindowTitle("PDF to Excel Extractor v0.0.1")
        self.setGeometry(100, 100, 800, 600)
        
        # 메인 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃
        main_layout = QVBoxLayout(central_widget)
        
        # 제목
        title_label = QLabel("PDF to Excel Extractor")
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 상단 버튼들
        top_button_layout = QHBoxLayout()
        
        # 도움말 버튼
        help_btn = QPushButton("❓ 도움말")
        help_btn.clicked.connect(self.show_help)
        top_button_layout.addWidget(help_btn)
        
        top_button_layout.addStretch()
        
        main_layout.addLayout(top_button_layout)
        
        # 탭 위젯
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        
        # 메인 탭
        main_tab = self.create_main_tab()
        tab_widget.addTab(main_tab, "파일 추출")
        
        # 설정 탭
        settings_tab = self.create_settings_tab()
        tab_widget.addTab(settings_tab, "추출 설정")
        
        # 진행률 표시 그룹
        progress_group = QGroupBox("진행 상황")
        progress_layout = QVBoxLayout(progress_group)
        
        # 전체 진행률
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid grey;
                border-radius: 5px;
                text-align: center;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
        """)
        progress_layout.addWidget(self.progress_bar)
        
        # 현재 파일 진행률 (파일 내 페이지 진행률)
        self.file_progress_bar = QProgressBar()
        self.file_progress_bar.setVisible(False)
        self.file_progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid grey;
                border-radius: 3px;
                text-align: center;
                font-size: 10px;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
                border-radius: 2px;
            }
        """)
        progress_layout.addWidget(self.file_progress_bar)
        
        # 상세 진행 정보
        progress_info_layout = QHBoxLayout()
        self.current_file_label = QLabel("현재 파일: 없음")
        self.current_file_label.setStyleSheet("font-size: 11px; color: #333;")
        progress_info_layout.addWidget(self.current_file_label)
        
        self.progress_details_label = QLabel("0/0 페이지")
        self.progress_details_label.setStyleSheet("font-size: 11px; color: #666;")
        progress_info_layout.addWidget(self.progress_details_label)
        progress_info_layout.addStretch();
        
        progress_layout.addLayout(progress_info_layout)
        main_layout.addWidget(progress_group)
        
        # 상태 및 메모리 사용량 표시
        status_layout = QHBoxLayout()
        self.status_label = QLabel("대기 중...")
        self.status_label.setStyleSheet("font-weight: bold; color: #2E7D32;")
        status_layout.addWidget(self.status_label)
        
        # 메모리 사용량 표시
        self.memory_label = QLabel("메모리: 0MB")
        self.memory_label.setStyleSheet("color: gray; font-size: 10px;")
        status_layout.addWidget(self.memory_label)
        
        # 처리 시간 및 속도 표시
        self.time_label = QLabel("처리 시간: 0초")
        self.time_label.setStyleSheet("color: gray; font-size: 10px;")
        status_layout.addWidget(self.time_label)
        
        # 처리 속도 표시
        self.speed_label = QLabel("처리 속도: -")
        self.speed_label.setStyleSheet("color: gray; font-size: 10px;")
        status_layout.addWidget(self.speed_label)
        status_layout.addStretch()
        
        main_layout.addLayout(status_layout)
        
        # 시스템 성능 모니터링 그룹
        perf_group = QGroupBox("시스템 성능")
        perf_layout = QGridLayout(perf_group)
        
        # CPU 사용률
        perf_layout.addWidget(QLabel("CPU 사용률:"), 0, 0)
        self.cpu_label = QLabel("0%")
        perf_layout.addWidget(self.cpu_label, 0, 1)
        
        # 메모리 사용률
        perf_layout.addWidget(QLabel("메모리 사용률:"), 1, 0)
        self.memory_label = QLabel("0 MB")
        perf_layout.addWidget(self.memory_label, 1, 1)
        
        # 처리 속도
        perf_layout.addWidget(QLabel("처리 속도:"), 2, 0)
        self.speed_label = QLabel("0 파일/초")
        perf_layout.addWidget(self.speed_label, 2, 1)
        
        # 예상 완료 시간
        perf_layout.addWidget(QLabel("예상 완료:"), 3, 0)
        self.eta_label = QLabel("--:--")
        perf_layout.addWidget(self.eta_label, 3, 1)
        
        progress_layout.addWidget(perf_group)
        
        # 성능 모니터링 타이머
        self.perf_timer = QTimer()
        self.perf_timer.timeout.connect(self.update_performance_info)
        self.perf_timer.start(1000)  # 1초마다 업데이트
        
        # 로그 영역
        log_group = QGroupBox("실시간 로그")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ddd;
                border-radius: 3px;
                font-family: 'Courier New', monospace;
                font-size: 11px;
            }
        """)
        log_layout.addWidget(self.log_text)
        
        # 로그 제어 버튼
        log_controls = QHBoxLayout()
        self.clear_log_btn = QPushButton("로그 지우기")
        self.clear_log_btn.clicked.connect(self.clear_log)
        self.clear_log_btn.setMaximumWidth(80)
        log_controls.addWidget(self.clear_log_btn)
        log_controls.addStretch()
        
        log_layout.addLayout(log_controls)
        main_layout.addWidget(log_group)
        
    def create_main_tab(self):
        """메인 탭 생성"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 파일 선택 그룹
        file_group = QGroupBox("PDF 파일 선택")
        file_layout = QVBoxLayout(file_group)
        
        file_btn_layout = QHBoxLayout()
        self.select_files_btn = QPushButton("PDF 파일 선택")
        self.select_files_btn.clicked.connect(self.select_pdf_files)
        file_btn_layout.addWidget(self.select_files_btn)
        
        self.select_folder_btn = QPushButton("폴더 선택 (일괄 처리)")
        self.select_folder_btn.clicked.connect(self.select_pdf_folder)
        file_btn_layout.addWidget(self.select_folder_btn)
        
        file_layout.addLayout(file_btn_layout)
        
        self.selected_files_label = QLabel("선택된 파일: 없음")
        file_layout.addWidget(self.selected_files_label)
        
        layout.addWidget(file_group)
        
        # 출력 설정 그룹
        output_group = QGroupBox("출력 설정")
        output_layout = QVBoxLayout(output_group)
        
        output_btn_layout = QHBoxLayout()
        self.output_dir_btn = QPushButton("출력 폴더 선택")
        self.output_dir_btn.clicked.connect(self.select_output_directory)
        output_btn_layout.addWidget(self.output_dir_btn)
        
        output_layout.addLayout(output_btn_layout)
        
        self.output_dir_label = QLabel("출력 폴더: 미선택")
        output_layout.addWidget(self.output_dir_label)
        
        layout.addWidget(output_group)
        
        # 실행 버튼
        self.extract_btn = QPushButton("추출 시작")
        self.extract_btn.clicked.connect(self.start_extraction)
        self.extract_btn.setEnabled(False)
        layout.addWidget(self.extract_btn)
        
        return tab
        
    def create_settings_tab(self):
        """설정 탭 생성"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # 추출 방법 설정
        method_group = QGroupBox("추출 방법 설정")
        method_layout = QGridLayout(method_group)
        
        method_layout.addWidget(QLabel("기본 추출 방법:"), 0, 0)
        self.method_combo = QComboBox()
        self.method_combo.addItems([
            "자동 선택", "pdfplumber", "camelot", "tabula", "pymupdf"
        ])
        method_layout.addWidget(self.method_combo, 0, 1)
        
        # OCR 설정
        self.ocr_checkbox = QCheckBox("OCR 사용 (이미지 PDF)")
        method_layout.addWidget(self.ocr_checkbox, 1, 0)
        
        self.ocr_combo = QComboBox()
        self.ocr_combo.addItems(["tesseract", "easyocr"])
        self.ocr_combo.setEnabled(False)
        self.ocr_checkbox.toggled.connect(self.ocr_combo.setEnabled)
        method_layout.addWidget(self.ocr_combo, 1, 1)
        
        layout.addWidget(method_group)
        
        # 테이블 감지 설정
        table_group = QGroupBox("테이블 감지 설정")
        table_layout = QGridLayout(table_group)
        
        table_layout.addWidget(QLabel("최소 행 수:"), 0, 0)
        self.min_rows_spin = QSpinBox()
        self.min_rows_spin.setMinimum(1)
        self.min_rows_spin.setMaximum(100)
        self.min_rows_spin.setValue(2)
        table_layout.addWidget(self.min_rows_spin, 0, 1)
        
        table_layout.addWidget(QLabel("최소 열 수:"), 1, 0)
        self.min_cols_spin = QSpinBox()
        self.min_cols_spin.setMinimum(1)
        self.min_cols_spin.setMaximum(50)
        self.min_cols_spin.setValue(2)
        table_layout.addWidget(self.min_cols_spin, 1, 1)
        
        layout.addWidget(table_group)
        
        # 대용량 PDF 처리 설정 (새로 추가)
        performance_group = QGroupBox("성능 최적화 설정")
        performance_layout = QGridLayout(performance_group)
        
        performance_layout.addWidget(QLabel("최대 처리 페이지:"), 0, 0)
        self.max_pages_spin = QSpinBox()
        self.max_pages_spin.setMinimum(10)
        self.max_pages_spin.setMaximum(2000)
        self.max_pages_spin.setValue(500)
        self.max_pages_spin.setToolTip("대용량 PDF 처리 시 메모리 절약을 위한 페이지 제한")
        performance_layout.addWidget(self.max_pages_spin, 0, 1)
        
        performance_layout.addWidget(QLabel("배치 크기:"), 1, 0)
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setMinimum(10)
        self.batch_size_spin.setMaximum(100)
        self.batch_size_spin.setValue(50)
        self.batch_size_spin.setToolTip("한 번에 처리할 페이지 수 (메모리 사용량 조절)")
        performance_layout.addWidget(self.batch_size_spin, 1, 1)
        
        layout.addWidget(performance_group)
        
        # 출력 설정
        output_group = QGroupBox("출력 형식 설정")
        output_layout = QVBoxLayout(output_group)
        
        self.separate_sheets_checkbox = QCheckBox("페이지별 시트 분리")
        output_layout.addWidget(self.separate_sheets_checkbox)
        
        self.include_text_checkbox = QCheckBox("텍스트 내용도 포함")
        output_layout.addWidget(self.include_text_checkbox)
        
        layout.addWidget(output_group)
        
        layout.addStretch()
        
        return tab
    
    def select_pdf_files(self):
        """PDF 파일 선택"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "PDF 파일 선택", "", "PDF files (*.pdf)"
        )
        if files:
            self.pdf_files = files
            self.update_file_label()
            self.check_ready_to_extract()
            
    def select_pdf_folder(self):
        """PDF 폴더 선택 (일괄 처리)"""
        folder = QFileDialog.getExistingDirectory(self, "PDF 폴더 선택")
        if folder:
            pdf_files = list(Path(folder).glob("*.pdf"))
            self.pdf_files = [str(f) for f in pdf_files]
            self.update_file_label()
            self.check_ready_to_extract()
            
    def select_output_directory(self):
        """출력 디렉토리 선택"""
        directory = QFileDialog.getExistingDirectory(self, "출력 폴더 선택")
        if directory:
            self.output_directory = directory
            self.output_dir_label.setText(f"출력 폴더: {directory}")
            self.check_ready_to_extract()
            
    def update_file_label(self):
        """선택된 파일 레이블 업데이트"""
        if self.pdf_files:
            count = len(self.pdf_files)
            if count == 1:
                filename = os.path.basename(self.pdf_files[0])
                self.selected_files_label.setText(f"선택된 파일: {filename}")
            else:
                self.selected_files_label.setText(f"선택된 파일: {count}개")
        else:
            self.selected_files_label.setText("선택된 파일: 없음")
            
    def update_memory_display(self):
        """메모리 사용량 표시 업데이트"""
        try:
            # 현재 프로세스의 메모리 사용량 (MB)
            process = psutil.Process()
            memory_info = process.memory_info()
            memory_mb = memory_info.rss / 1024 / 1024
            
            # 시스템 전체 메모리 사용률
            system_memory = psutil.virtual_memory()
            system_usage = system_memory.percent
            
            status_text = f"메모리: {memory_mb:.1f}MB | 시스템: {system_usage:.1f}%"
            if hasattr(self, 'memory_label'):
                self.memory_label.setText(status_text)
        except Exception as e:
            if hasattr(self, 'memory_label'):
                self.memory_label.setText("메모리 정보 확인 불가")
    
    def check_ready_to_extract(self):
        """추출 준비 상태 확인"""
        ready = bool(self.pdf_files and self.output_directory)
        if hasattr(self, 'extract_btn'):
            self.extract_btn.setEnabled(ready)
        
        # 상태 표시 업데이트
        if ready:
            if hasattr(self, 'status_label'):
                self.status_label.setText("추출 준비 완료")
        else:
            if hasattr(self, 'status_label'):
                self.status_label.setText("파일과 출력 폴더를 선택하세요")
    
    def get_extraction_options(self):
        """추출 옵션 수집"""
        max_pages_value = getattr(self, 'max_pages_spin', type('obj', (object,), {'value': lambda: 500})).value()
        # 0인 경우 500으로 변경 (모든 페이지 처리를 위함)
        if max_pages_value == 0:
            max_pages_value = 500
            
        return {
            'min_rows': getattr(self, 'min_rows_spin', type('obj', (object,), {'value': lambda: 3})).value(),
            'min_cols': getattr(self, 'min_cols_spin', type('obj', (object,), {'value': lambda: 2})).value(),
            'separate_sheets': getattr(self, 'separate_sheets_checkbox', type('obj', (object,), {'isChecked': lambda: True})).isChecked(),
            'include_text': getattr(self, 'include_text_checkbox', type('obj', (object,), {'isChecked': lambda: True})).isChecked(),
            'max_pages': max_pages_value,
            'batch_size': getattr(self, 'batch_size_spin', type('obj', (object,), {'value': lambda: 5})).value()
        }
        
    def start_extraction(self):
        """추출 시작"""
        if not self.pdf_files or not self.output_directory:
            QMessageBox.warning(self, "경고", "PDF 파일과 출력 폴더를 모두 선택해주세요.")
            return
            
        # UI 상태 초기화
        if hasattr(self, 'extract_btn'):
            self.extract_btn.setEnabled(False)
        if hasattr(self, 'progress_bar'):
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
        if hasattr(self, 'file_progress_bar'):
            self.file_progress_bar.setVisible(True)
            self.file_progress_bar.setValue(0)
        
        # 시작 시간 기록
        self.start_time = QTime.currentTime()
        
        # 추출 옵션 수집
        options = self.get_extraction_options()
        
        # 백그라운드 스레드에서 추출 수행
        self.extractor_thread = ExtractorThread(
            self.pdf_files, self.output_directory, options
        )
        
        # 시그널 연결
        if hasattr(self, 'progress_bar'):
            self.extractor_thread.progress.connect(self.progress_bar.setValue)
        if hasattr(self, 'file_progress_bar'):
            self.extractor_thread.file_progress.connect(self.file_progress_bar.setValue)
        if hasattr(self, 'status_label'):
            self.extractor_thread.status.connect(self.status_label.setText)
        if hasattr(self, 'current_file_label'):
            self.extractor_thread.current_file.connect(self.current_file_label.setText)
        if hasattr(self, 'progress_details_label'):
            self.extractor_thread.progress_details.connect(self.progress_details_label.setText)
        if hasattr(self, 'log_text'):
            self.extractor_thread.log_message.connect(self.add_log_message)
        if hasattr(self, 'speed_label'):
            self.extractor_thread.processing_stats.connect(self.update_processing_stats)
        
        self.extractor_thread.finished.connect(self.extraction_finished)
        self.extractor_thread.start()
        
        if hasattr(self, 'log_text'):
            self.add_log_message("🚀 추출 작업을 시작합니다...")
        
        # 시간 업데이트 타이머 시작
        if not hasattr(self, 'time_timer'):
            self.time_timer = QTimer()
            self.time_timer.timeout.connect(self.update_elapsed_time)
        self.time_timer.start(1000)  # 1초마다 업데이트
        
    def show_help(self):
        """도움말 다이얼로그 표시"""
        help_dialog = HelpDialog(self)
        help_dialog.exec_()
    
    def extraction_finished(self, success: bool, message: str):
        """추출 작업 완료 처리"""
        self.progress_bar.setVisible(False)
        self.file_progress_bar.setVisible(False)
        if hasattr(self, 'extract_btn'):
            self.extract_btn.setEnabled(True)
        
        if success:
            self.status_label.setText("✅ 완료!")
            self.status_label.setStyleSheet("font-weight: bold; color: #4CAF50;")
            
            # 처리 완료 다이얼로그 표시
            if self.output_directory:
                results_dialog = ResultsDialog(
                    self.output_directory, 
                    len(self.pdf_files), 
                    self
                )
                results_dialog.exec_()
        else:
            self.status_label.setText("❌ 실패")
            self.status_label.setStyleSheet("font-weight: bold; color: #F44336;")
            QMessageBox.critical(self, "오류", message)
    
    def add_log_message(self, message: str):
        """로그 메시지 추가"""
        if hasattr(self, 'log_text'):
            current_time = QTime.currentTime().toString('hh:mm:ss')
            formatted_message = f"[{current_time}] {message}"
            self.log_text.append(formatted_message)
            # 자동 스크롤
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.End)
            self.log_text.setTextCursor(cursor)
    
    def clear_log(self):
        """로그 지우기"""
        if hasattr(self, 'log_text'):
            self.log_text.clear()
            self.add_log_message("📝 로그가 지워졌습니다.")
    
    def update_elapsed_time(self):
        """경과 시간 및 처리 속도 업데이트"""
        if hasattr(self, 'start_time') and hasattr(self, 'time_label'):
            elapsed = self.start_time.secsTo(QTime.currentTime())
            if elapsed >= 60:
                minutes = elapsed // 60
                seconds = elapsed % 60
                time_text = f"처리 시간: {minutes}분 {seconds}초"
            else:
                time_text = f"처리 시간: {elapsed}초"
            self.time_label.setText(time_text)
            
            # 처리 속도 계산 (진행률 기준)
            if hasattr(self, 'speed_label') and hasattr(self, 'progress_bar') and elapsed > 0:
                progress = self.progress_bar.value()
                if progress > 0:
                    speed = progress / elapsed
                    if speed >= 1:
                        speed_text = f"처리 속도: {speed:.1f}%/초"
                    else:
                        eta = (100 - progress) / speed if speed > 0 else 0
                        speed_text = f"예상 완료: {eta:.0f}초 후"
                    self.speed_label.setText(speed_text)
                else:
                    self.speed_label.setText("처리 속도: 계산 중...")

    def save_settings(self):
        """설정을 파일에 저장"""
        settings = {
            'last_output_directory': self.output_directory,
            'extraction_method': getattr(self, 'method_combo', None) and self.method_combo.currentText(),
            'max_pages': getattr(self, 'max_pages_spin', None) and self.max_pages_spin.value(),
            'use_ocr': getattr(self, 'use_ocr_check', None) and self.use_ocr_check.isChecked(),
            'include_text': getattr(self, 'include_text_check', None) and self.include_text_check.isChecked(),
            'separate_sheets': getattr(self, 'separate_sheets_check', None) and self.separate_sheets_check.isChecked()
        }
        
        try:
            with open('settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"설정 저장 실패: {e}")
    
    def load_settings(self):
        """설정 파일 로드"""
        try:
            with open('settings.json', 'r', encoding='utf-8') as f:
                settings = json.load(f)
                
            self.output_directory = settings.get('last_output_directory', '')
                
        except Exception as e:
            print(f"설정 로드 실패: {e}")

    def update_performance_info(self):
        """성능 정보 업데이트"""
        try:
            # CPU 사용률
            cpu_percent = psutil.cpu_percent(interval=None)
            if hasattr(self, 'cpu_label'):
                self.cpu_label.setText(f"{cpu_percent:.1f}%")
                
                # CPU 사용률에 따른 색상 변경
                if cpu_percent > 80:
                    self.cpu_label.setStyleSheet("color: #F44336; font-weight: bold;")  # 빨간색
                elif cpu_percent > 60:
                    self.cpu_label.setStyleSheet("color: #FF9800; font-weight: bold;")  # 주황색
                else:
                    self.cpu_label.setStyleSheet("color: #4CAF50; font-weight: bold;")  # 초록색
            
            # 메모리 사용률
            memory = psutil.virtual_memory()
            memory_mb = (memory.total - memory.available) / (1024 * 1024)
            memory_percent = memory.percent
            if hasattr(self, 'memory_label'):
                self.memory_label.setText(f"{memory_mb:.0f} MB ({memory_percent:.1f}%)")
                
                # 메모리 사용률에 따른 색상 변경
                if memory_percent > 85:
                    self.memory_label.setStyleSheet("color: #F44336; font-weight: bold;")  # 빨간색
                elif memory_percent > 70:
                    self.memory_label.setStyleSheet("color: #FF9800; font-weight: bold;")  # 주황색
                else:
                    self.memory_label.setStyleSheet("color: #4CAF50; font-weight: bold;")  # 초록색
                    
        except Exception as e:
            print(f"성능 정보 업데이트 실패: {e}")

    def update_processing_stats(self, files_processed: int, total_files: int, elapsed_time: float):
        """처리 통계 업데이트"""
        if elapsed_time > 0:
            # 처리 속도 계산
            speed = files_processed / elapsed_time
            if hasattr(self, 'speed_label'):
                self.speed_label.setText(f"{speed:.2f} 파일/초")
            
            # 예상 완료 시간 계산
            if speed > 0 and hasattr(self, 'eta_label'):
                remaining_files = total_files - files_processed
                eta_seconds = remaining_files / speed
                
                if eta_seconds < 60:
                    eta_text = f"{eta_seconds:.0f}초"
                elif eta_seconds < 3600:
                    eta_text = f"{eta_seconds/60:.0f}분 {eta_seconds%60:.0f}초"
                else:
                    hours = eta_seconds // 3600
                    minutes = (eta_seconds % 3600) // 60
                    eta_text = f"{hours:.0f}시간 {minutes:.0f}분"
                
                self.eta_label.setText(eta_text)
            else:
                if hasattr(self, 'eta_label'):
                    self.eta_label.setText("계산 중...")
        else:
            if hasattr(self, 'speed_label'):
                self.speed_label.setText("0 파일/초")
            if hasattr(self, 'eta_label'):
                self.eta_label.setText("--:--")

    def closeEvent(self, event):
        """애플리케이션 종료 시 설정 저장"""
        self.save_settings()
        super().closeEvent(event)

class HelpDialog(QDialog):
    """도움말 다이얼로그"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDF Extractor 도움말")
        self.setFixedSize(600, 500)
        self.setup_ui()
        
    def setup_ui(self):
        """UI 설정"""
        layout = QVBoxLayout()
        
        # 도움말 내용
        help_text = QTextEdit()
        help_text.setReadOnly(True)
        help_text.setHtml("""
        <h2>PDF to Excel Extractor 사용법</h2>
        
        <h3>🔧 기본 사용법</h3>
        <p><b>1단계:</b> "PDF 파일 선택" 버튼을 클릭하여 처리할 PDF 파일들을 선택합니다.</p>
        <p><b>2단계:</b> "출력 폴더 선택" 버튼을 클릭하여 결과 파일이 저장될 폴더를 선택합니다.</p>
        <p><b>3단계:</b> 필요에 따라 추출 설정을 조정합니다.</p>
        <p><b>4단계:</b> "추출 시작" 버튼을 클릭하여 처리를 시작합니다.</p>
        
        <h3>⚙️ 추출 설정</h3>
        <p><b>최대 페이지 수:</b> 0으로 설정하면 모든 페이지를 처리합니다.</p>
        <p><b>추출 방법:</b></p>
        <ul>
            <li><b>camelot:</b> 표 구조가 명확한 PDF에 적합</li>
            <li><b>tabula:</b> 다양한 형태의 표에 대응</li>
            <li><b>pdfplumber:</b> 텍스트 기반 표 추출</li>
        </ul>
        <p><b>Ghostscript 사용:</b> 이미지 기반 PDF 처리 시 품질 향상</p>
        
        <h3>🎨 고급 기능</h3>
        <p><b>다크 모드:</b> 우상단 "다크 모드" 버튼으로 테마 전환</p>
        <p><b>성능 모니터링:</b> 실시간 CPU/메모리 사용량 확인</p>
        <p><b>진행률 추적:</b> 전체 진행률과 개별 파일 진행률 모니터링</p>
        <p><b>로그 기록:</b> 처리 과정의 상세 로그 확인 및 저장</p>
        
        <h3>🔍 문제 해결</h3>
        <p><b>Ghostscript 오류:</b> 다른 추출 방법으로 대체되므로 무시해도 됩니다.</p>
        <p><b>메모리 부족:</b> 대용량 파일 처리 시 최대 페이지 수를 제한하세요.</p>
        <p><b>빈 결과:</b> 다른 추출 방법을 시도해보세요.</p>
        
        <h3>💡 팁</h3>
        <p>• 처리 중에도 로그를 통해 진행 상황을 확인할 수 있습니다.</p>
        <p>• 설정은 자동으로 저장되어 다음 실행 시 복원됩니다.</p>
        <p>• 대용량 파일 처리 시 성능 모니터를 참고하세요.</p>
        """)
        
        layout.addWidget(help_text)
        
        # 버튼
        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(self.accept)
        layout.addWidget(button_box)
        
        self.setLayout(layout)

class ResultsDialog(QDialog):
    """처리 결과 확인 다이얼로그"""
    
    def __init__(self, output_folder: str, processed_files: int, parent=None):
        super().__init__(parent)
        self.output_folder = output_folder
        self.processed_files = processed_files
        self.setWindowTitle("처리 완료")
        self.setFixedSize(400, 200)
        self.setup_ui()
        
    def setup_ui(self):
        """UI 설정"""
        layout = QVBoxLayout()
        
        # 결과 정보
        info_label = QLabel(f"""
        <h3>✅ 처리가 완료되었습니다!</h3>
        <p><b>처리된 파일:</b> {self.processed_files}개</p>
        <p><b>출력 폴더:</b> {self.output_folder}</p>
        """)
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # 버튼들
        button_layout = QHBoxLayout()
        
        open_folder_btn = QPushButton("📁 폴더 열기")
        open_folder_btn.clicked.connect(self.open_folder)
        button_layout.addWidget(open_folder_btn)
        
        close_btn = QPushButton("닫기")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
    def open_folder(self):
        """결과 폴더 열기"""
        try:
            if sys.platform == "win32":
                os.startfile(self.output_folder)
            elif sys.platform == "darwin":  # macOS
                os.system(f"open '{self.output_folder}'")
            else:  # Linux
                os.system(f"xdg-open '{self.output_folder}'")
        except Exception as e:
            QMessageBox.warning(self, "오류", f"폴더를 열 수 없습니다: {e}")


def main():
    """메인 함수"""
    app = QApplication(sys.argv)
    
    # 애플리케이션 정보 설정
    app.setApplicationName("PDF to Excel Extractor")
    app.setApplicationVersion("0.0.1")
    app.setOrganizationName("PDF Extractor")
    
    # 메인 윈도우 생성 및 표시
    window = PDFExtractorGUI()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
