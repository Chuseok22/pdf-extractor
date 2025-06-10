#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF to Excel Extractor - Tkinter Version
PDF 파일에서 데이터를 추출하여 Excel 파일로 저장하는 GUI 애플리케이션
"""

import sys
import os
import traceback
import json
import time
import threading
from pathlib import Path
from typing import List
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.ttk import Progressbar, Notebook

from pdf_extractor import PDFExtractor


class ExtractorThread(threading.Thread):
    """PDF 추출 작업을 백그라운드에서 수행하는 스레드"""
    
    def __init__(self, pdf_files: List[str], output_dir: str, options: dict, callback_manager):
        super().__init__()
        self.pdf_files = pdf_files
        self.output_dir = output_dir
        self.options = options
        self.callback_manager = callback_manager
        self.extractor = PDFExtractor()
        self.daemon = True  # 메인 프로그램 종료시 함께 종료
        
    def run(self):
        try:
            start_time = time.time()
            
            total_files = len(self.pdf_files)
            self.callback_manager.status_update("추출 작업 시작...")
            self.callback_manager.log_message(f"📁 총 {total_files}개 파일 처리 시작")
            
            for i, pdf_file in enumerate(self.pdf_files):
                filename = os.path.basename(pdf_file)
                self.callback_manager.current_file_update(f"현재 파일: {filename}")
                
                try:
                    file_size = os.path.getsize(pdf_file) / (1024 * 1024)  # MB
                    size_info = f"({file_size:.1f}MB)" if file_size > 1 else f"({file_size*1024:.0f}KB)"
                    self.callback_manager.log_message(f"📄 처리 시작: {filename} {size_info}")
                    
                    if file_size > 50:
                        self.callback_manager.status_update(f"대용량 파일 처리 중...")
                        self.callback_manager.log_message(f"⚠️  대용량 파일 감지: {file_size:.1f}MB")
                    else:
                        self.callback_manager.status_update(f"파일 처리 중...")
                        
                except Exception as e:
                    self.callback_manager.log_message(f"파일 크기 확인 실패: {str(e)}")
                
                # PDF 처리
                try:
                    output_excel = os.path.join(
                        self.output_dir, 
                        f"{os.path.splitext(filename)[0]}_extracted.xlsx"
                    )
                    
                    self.callback_manager.file_progress_update(0)
                    
                    # PDF에서 텍스트와 테이블 추출
                    def progress_callback(current_page, total_pages):
                        progress = int((current_page / total_pages) * 100)
                        self.callback_manager.file_progress_update(progress)
                        self.callback_manager.progress_details_update(f"페이지 {current_page}/{total_pages}")
                    
                    self.extractor.extract_pdf(
                        pdf_file, 
                        output_excel, 
                        progress_callback=progress_callback,
                        **self.options
                    )
                    
                    self.callback_manager.file_progress_update(100)
                    self.callback_manager.log_message(f"✅ 완료: {filename}")
                    
                except Exception as e:
                    error_msg = f"❌ 실패: {filename} - {str(e)}"
                    self.callback_manager.log_message(error_msg)
                    print(f"처리 오류: {str(e)}")
                    traceback.print_exc()
                
                # 전체 진행률 업데이트
                overall_progress = int(((i + 1) / total_files) * 100)
                self.callback_manager.progress_update(overall_progress)
                
                elapsed_time = time.time() - start_time
                self.callback_manager.processing_stats_update(i + 1, total_files, elapsed_time)
            
            # 완료
            elapsed_time = time.time() - start_time
            self.callback_manager.status_update("모든 작업 완료!")
            self.callback_manager.log_message(f"🎉 추출 완료! 총 소요시간: {elapsed_time:.1f}초")
            self.callback_manager.finished(True, "추출이 완료되었습니다.")
            
        except Exception as e:
            error_msg = f"추출 중 오류 발생: {str(e)}"
            self.callback_manager.log_message(f"❌ {error_msg}")
            traceback.print_exc()
            self.callback_manager.finished(False, error_msg)


class CallbackManager:
    """스레드와 GUI 간 콜백 관리"""
    
    def __init__(self, main_window):
        self.main_window = main_window
    
    def status_update(self, message):
        self.main_window.after(0, self.main_window.update_status, message)
    
    def log_message(self, message):
        self.main_window.after(0, self.main_window.add_log_message, message)
    
    def current_file_update(self, message):
        self.main_window.after(0, self.main_window.update_current_file, message)
    
    def progress_update(self, value):
        self.main_window.after(0, self.main_window.update_progress, value)
    
    def file_progress_update(self, value):
        self.main_window.after(0, self.main_window.update_file_progress, value)
    
    def progress_details_update(self, message):
        self.main_window.after(0, self.main_window.update_progress_details, message)
    
    def processing_stats_update(self, processed, total, elapsed):
        self.main_window.after(0, self.main_window.update_processing_stats, processed, total, elapsed)
    
    def finished(self, success, message):
        self.main_window.after(0, self.main_window.extraction_finished, success, message)


class PDFExtractorGUI:
    """PDF Extractor GUI 메인 클래스"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF to Excel Extractor")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # 변수들 초기화
        self.pdf_files = []
        self.output_dir = ""
        self.extractor_thread = None
        self.start_time = 0
        
        # 설정 파일 경로
        self.settings_file = Path("settings.json")
        self.settings = self.load_settings()
        
        self.setup_gui()
        self.load_window_settings()
        
        # 종료 시 설정 저장
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_gui(self):
        """GUI 구성 요소 초기화"""
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 제목
        title_label = ttk.Label(main_frame, text="PDF to Excel Extractor", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # 파일 선택 섹션
        files_frame = ttk.LabelFrame(main_frame, text="1. PDF 파일 선택", padding="10")
        files_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)
        
        ttk.Button(files_frame, text="파일 선택", 
                  command=self.select_files).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(files_frame, text="폴더 선택", 
                  command=self.select_folder).grid(row=0, column=1, padx=(0, 10))
        
        self.files_label = ttk.Label(files_frame, text="선택된 파일이 없습니다.")
        self.files_label.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # 출력 폴더 선택 섹션
        output_frame = ttk.LabelFrame(main_frame, text="2. 출력 폴더 선택", padding="10")
        output_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Button(output_frame, text="폴더 선택", 
                  command=self.select_output_folder).grid(row=0, column=0, padx=(0, 10))
        
        self.output_label = ttk.Label(output_frame, text="출력 폴더를 선택하세요.")
        self.output_label.grid(row=0, column=1, sticky=(tk.W, tk.E))
        
        # 추출 옵션 섹션
        options_frame = ttk.LabelFrame(main_frame, text="3. 추출 옵션", padding="10")
        options_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 추출 방법 선택
        method_frame = ttk.Frame(options_frame)
        method_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        ttk.Label(method_frame, text="추출 방법:").grid(row=0, column=0, padx=(0, 10))
        self.method_var = tk.StringVar(value="auto")
        methods = [("자동 감지", "auto"), ("텍스트만", "text"), ("테이블만", "table")]
        for i, (text, value) in enumerate(methods):
            ttk.Radiobutton(method_frame, text=text, variable=self.method_var, 
                           value=value).grid(row=0, column=i+1, padx=5)
        
        # 고급 옵션
        advanced_frame = ttk.Frame(options_frame)
        advanced_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        self.ocr_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(advanced_frame, text="OCR 사용 (이미지 텍스트 추출)", 
                       variable=self.ocr_var).grid(row=0, column=0, sticky=tk.W)
        
        # 제어 버튼
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=4, column=0, columnspan=2, pady=(0, 10))
        
        self.start_button = ttk.Button(control_frame, text="추출 시작", 
                                      command=self.start_extraction)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.stop_button = ttk.Button(control_frame, text="중지", 
                                     command=self.stop_extraction, state="disabled")
        self.stop_button.pack(side=tk.LEFT)
        
        # 진행 상황 섹션
        progress_frame = ttk.LabelFrame(main_frame, text="진행 상황", padding="10")
        progress_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        # 상태 표시
        self.status_label = ttk.Label(progress_frame, text="대기 중...")
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 현재 파일
        self.current_file_label = ttk.Label(progress_frame, text="")
        self.current_file_label.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # 전체 진행률
        ttk.Label(progress_frame, text="전체 진행률:").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        self.progress_bar = Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 파일 진행률
        ttk.Label(progress_frame, text="현재 파일 진행률:").grid(row=4, column=0, sticky=tk.W, pady=(10, 0))
        self.file_progress_bar = Progressbar(progress_frame, mode='determinate')
        self.file_progress_bar.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 상세 정보
        self.details_label = ttk.Label(progress_frame, text="")
        self.details_label.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 처리 통계
        self.stats_label = ttk.Label(progress_frame, text="")
        self.stats_label.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # 로그 출력 섹션
        log_frame = ttk.LabelFrame(main_frame, text="작업 로그", padding="10")
        log_frame.grid(row=6, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 로그 제어 버튼
        log_control_frame = ttk.Frame(log_frame)
        log_control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        ttk.Button(log_control_frame, text="로그 지우기", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(log_control_frame, text="로그 저장", 
                  command=self.save_log).pack(side=tk.LEFT)
    
    def select_files(self):
        """PDF 파일들 선택"""
        files = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if files:
            self.pdf_files = list(files)
            self.update_files_display()
    
    def select_folder(self):
        """PDF 파일들이 있는 폴더 선택"""
        folder = filedialog.askdirectory(title="PDF 폴더 선택")
        if folder:
            pdf_files = list(Path(folder).glob("*.pdf"))
            if pdf_files:
                self.pdf_files = [str(f) for f in pdf_files]
                self.update_files_display()
            else:
                messagebox.showwarning("경고", "선택한 폴더에 PDF 파일이 없습니다.")
    
    def select_output_folder(self):
        """출력 폴더 선택"""
        folder = filedialog.askdirectory(title="출력 폴더 선택")
        if folder:
            self.output_dir = folder
            self.output_label.config(text=f"출력 폴더: {folder}")
    
    def update_files_display(self):
        """선택된 파일 목록 표시 업데이트"""
        count = len(self.pdf_files)
        if count == 0:
            self.files_label.config(text="선택된 파일이 없습니다.")
        elif count == 1:
            filename = os.path.basename(self.pdf_files[0])
            self.files_label.config(text=f"선택된 파일: {filename}")
        else:
            self.files_label.config(text=f"선택된 파일: {count}개")
    
    def start_extraction(self):
        """추출 작업 시작"""
        # 입력 검증
        if not self.pdf_files:
            messagebox.showerror("오류", "PDF 파일을 선택하세요.")
            return
        
        if not self.output_dir:
            messagebox.showerror("오류", "출력 폴더를 선택하세요.")
            return
        
        # 출력 폴더 존재 확인
        try:
            os.makedirs(self.output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("오류", f"출력 폴더를 생성할 수 없습니다: {str(e)}")
            return
        
        # 추출 옵션 수집
        options = {
            'method': self.method_var.get(),
            'use_ocr': self.ocr_var.get(),
        }
        
        # UI 상태 변경
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        
        # 진행 상황 초기화
        self.progress_bar['value'] = 0
        self.file_progress_bar['value'] = 0
        self.clear_log()
        
        self.start_time = time.time()
        
        # 추출 스레드 시작
        callback_manager = CallbackManager(self)
        self.extractor_thread = ExtractorThread(
            self.pdf_files, self.output_dir, options, callback_manager
        )
        self.extractor_thread.start()
    
    def stop_extraction(self):
        """추출 작업 중지"""
        if self.extractor_thread and self.extractor_thread.is_alive():
            # Python threading은 강제 종료가 어려우므로 사용자에게 알림
            result = messagebox.askyesno(
                "작업 중지", 
                "현재 진행 중인 작업을 중지하시겠습니까?\n"
                "진행 중인 파일 처리가 완료된 후 중지됩니다."
            )
            if result:
                self.add_log_message("❌ 사용자에 의해 작업이 중지됩니다...")
                self.update_status("작업 중지 중...")
                # 스레드는 자연스럽게 종료되도록 함
                self.extraction_finished(False, "작업이 중지되었습니다.")
    
    def extraction_finished(self, success, message):
        """추출 작업 완료 처리"""
        # UI 상태 복원
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        if success:
            self.update_status("추출 완료!")
            messagebox.showinfo("완료", message)
            
            # 출력 폴더 열기 옵션
            result = messagebox.askyesno("완료", f"{message}\n\n출력 폴더를 여시겠습니까?")
            if result:
                self.open_output_folder()
        else:
            self.update_status("추출 실패")
            messagebox.showerror("오류", message)
    
    def open_output_folder(self):
        """출력 폴더 열기"""
        try:
            import subprocess
            import platform
            
            if platform.system() == "Darwin":  # macOS
                subprocess.run(["open", self.output_dir])
            elif platform.system() == "Windows":
                subprocess.run(["explorer", self.output_dir])
            else:  # Linux
                subprocess.run(["xdg-open", self.output_dir])
        except Exception as e:
            print(f"폴더 열기 실패: {str(e)}")
    
    # 콜백 메서드들
    def update_status(self, message):
        self.status_label.config(text=message)
    
    def update_current_file(self, message):
        self.current_file_label.config(text=message)
    
    def update_progress(self, value):
        self.progress_bar['value'] = value
    
    def update_file_progress(self, value):
        self.file_progress_bar['value'] = value
    
    def update_progress_details(self, message):
        self.details_label.config(text=message)
    
    def update_processing_stats(self, processed, total, elapsed):
        if total > 0:
            remaining = total - processed
            if processed > 0:
                avg_time = elapsed / processed
                estimated_remaining = avg_time * remaining
                stats_text = (f"진행: {processed}/{total}개 파일 "
                            f"| 경과: {elapsed:.1f}초 "
                            f"| 예상 남은 시간: {estimated_remaining:.1f}초")
            else:
                stats_text = f"진행: {processed}/{total}개 파일 | 경과: {elapsed:.1f}초"
            
            self.stats_label.config(text=stats_text)
    
    def add_log_message(self, message):
        timestamp = time.strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def save_log(self):
        """로그를 파일로 저장"""
        try:
            log_content = self.log_text.get(1.0, tk.END)
            if not log_content.strip():
                messagebox.showinfo("정보", "저장할 로그가 없습니다.")
                return
            
            filename = filedialog.asksaveasfilename(
                title="로그 저장",
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(log_content)
                messagebox.showinfo("완료", f"로그가 저장되었습니다:\n{filename}")
        except Exception as e:
            messagebox.showerror("오류", f"로그 저장 중 오류가 발생했습니다:\n{str(e)}")
    
    def load_settings(self):
        """설정 파일 로드"""
        default_settings = {
            'window_geometry': '800x700+100+100',
            'last_output_dir': str(Path.home() / 'Desktop'),
            'extraction_method': 'auto',
            'use_ocr': False
        }
        
        try:
            if self.settings_file.exists():
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    saved_settings = json.load(f)
                    default_settings.update(saved_settings)
        except Exception as e:
            print(f"설정 로드 실패: {e}")
        
        return default_settings
    
    def save_settings(self):
        """설정 파일 저장"""
        try:
            settings = {
                'window_geometry': self.root.geometry(),
                'last_output_dir': self.output_dir,
                'extraction_method': self.method_var.get(),
                'use_ocr': self.ocr_var.get()
            }
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"설정 저장 실패: {e}")
    
    def load_window_settings(self):
        """윈도우 설정 로드"""
        try:
            geometry = self.settings.get('window_geometry', '800x700+100+100')
            self.root.geometry(geometry)
            
            # 기본값 설정
            self.output_dir = self.settings.get('last_output_dir', str(Path.home() / 'Desktop'))
            self.output_label.config(text=f"출력 폴더: {self.output_dir}")
            
            self.method_var.set(self.settings.get('extraction_method', 'auto'))
            self.ocr_var.set(self.settings.get('use_ocr', False))
            
        except Exception as e:
            print(f"윈도우 설정 로드 실패: {e}")
    
    def on_closing(self):
        """애플리케이션 종료 시 처리"""
        # 진행 중인 작업이 있다면 확인
        if self.extractor_thread and self.extractor_thread.is_alive():
            result = messagebox.askyesno(
                "프로그램 종료", 
                "작업이 진행 중입니다. 정말 종료하시겠습니까?"
            )
            if not result:
                return
        
        # 설정 저장
        self.save_settings()
        
        # 프로그램 종료
        self.root.quit()
        self.root.destroy()
    
    def run(self):
        """GUI 애플리케이션 실행"""
        self.root.mainloop()


def main():
    """메인 함수"""
    try:
        app = PDFExtractorGUI()
        app.run()
    except Exception as e:
        print(f"애플리케이션 시작 실패: {str(e)}")
        traceback.print_exc()
        
        # GUI 실행이 실패한 경우 콘솔 모드로 fallback
        print("\nGUI 모드 실행 실패. 콘솔 모드로 전환합니다.")
        print("사용법: python main_tkinter.py [--file PDF파일경로] [--directory PDF폴더경로]")


if __name__ == "__main__":
    main()
