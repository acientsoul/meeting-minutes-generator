# -*- coding: utf-8 -*-
"""
회의록 자동 생성 GUI (tkinter 버전 - Python 기본 포함)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import sys
from datetime import datetime
import threading

class MeetingRecorderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎤 회의록 자동 생성")
        self.root.geometry("700x800")
        self.generated_file = None
        
        # 회의 정보
        tk.Label(root, text="🎤 회의록 자동 생성", font=("Arial", 16, "bold")).pack(pady=10)
        tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, padx=5)
        
        # 입력 필드
        frame = tk.Frame(root)
        frame.pack(pady=10, padx=10, fill=tk.BOTH)
        
        tk.Label(frame, text="회의명:", width=10).grid(row=0, column=0, sticky=tk.W)
        self.entry_제목 = tk.Entry(frame, width=40)
        self.entry_제목.insert(0, "정기 회의")
        self.entry_제목.grid(row=0, column=1, sticky=tk.W)
        
        tk.Label(frame, text="장소:", width=10).grid(row=1, column=0, sticky=tk.W)
        self.entry_장소 = tk.Entry(frame, width=40)
        self.entry_장소.insert(0, "미정")
        self.entry_장소.grid(row=1, column=1, sticky=tk.W)
        
        tk.Label(frame, text="일시:", width=10).grid(row=2, column=0, sticky=tk.W)
        self.entry_일시 = tk.Entry(frame, width=40)
        self.entry_일시.insert(0, datetime.now().strftime("%Y.%m.%d"))
        self.entry_일시.grid(row=2, column=1, sticky=tk.W)
        
        tk.Label(frame, text="작성자:", width=10).grid(row=3, column=0, sticky=tk.W)
        self.entry_작성자 = tk.Entry(frame, width=40)
        self.entry_작성자.insert(0, "미정")
        self.entry_작성자.grid(row=3, column=1, sticky=tk.W)
        
        tk.Label(frame, text="참석자:", width=10).grid(row=4, column=0, sticky=tk.W)
        self.entry_참석자 = tk.Entry(frame, width=40)
        self.entry_참석자.insert(0, "미정")
        self.entry_참석자.grid(row=4, column=1, sticky=tk.W)
        
        # 파일 선택
        tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, padx=5)
        file_frame = tk.Frame(root)
        file_frame.pack(pady=10, padx=10, fill=tk.X)
        
        tk.Label(file_frame, text="음성파일:", width=10).pack(side=tk.LEFT)
        self.label_파일 = tk.Label(file_frame, text="선택 안 함", fg="gray")
        self.label_파일.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(file_frame, text="선택", command=self.select_file, width=8).pack(side=tk.RIGHT)
        
        # 출력 파일
        tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, padx=5)
        out_frame = tk.Frame(root)
        out_frame.pack(pady=10, padx=10, fill=tk.X)
        
        tk.Label(out_frame, text="출력파일:", width=10).pack(side=tk.LEFT)
        self.entry_출력 = tk.Entry(out_frame)
        self.entry_출력.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
        self.entry_출력.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 진행률
        tk.Frame(root, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, padx=5)
        
        self.progress = tk.Canvas(root, height=20, bg="lightgray")
        self.progress.pack(pady=10, padx=10, fill=tk.X)
        self.progress_fill = None
        
        self.label_상태 = tk.Label(root, text="상태: 준비 완료", fg="blue")
        self.label_상태.pack()
        
        # 로그
        self.log_text = scrolledtext.ScrolledText(root, height=8, width=80)
        self.log_text.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 버튼
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="🚀 변환 시작", command=self.convert, width=12).pack(side=tk.LEFT, padx=5)
        self.btn_폴더 = tk.Button(btn_frame, text="📂 폴더 열기", command=self.open_folder, width=12, state=tk.DISABLED)
        self.btn_폴더.pack(side=tk.LEFT, padx=5)
        self.btn_파일 = tk.Button(btn_frame, text="📄 파일 열기", command=self.open_file, width=12, state=tk.DISABLED)
        self.btn_파일.pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="초기화", command=self.reset, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="종료", command=root.quit, width=12).pack(side=tk.LEFT, padx=5)
        
        self.selected_file = None
        self.log("프로그램 시작됨\n1. 회의 정보 입력\n2. 음성 파일 선택\n3. '변환 시작' 클릭")
    
    def select_file(self):
        file = filedialog.askopenfilename(filetypes=[("M4A Files", "*.m4a")])
        if file:
            self.selected_file = file
            size = os.path.getsize(file) / (1024 * 1024)
            self.label_파일.config(text=f"✓ {os.path.basename(file)} ({size:.1f}MB)", fg="green")
            self.log(f"파일 선택: {os.path.basename(file)}")
    
    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update()
    
    def update_progress(self, value):
        self.progress.delete("all")
        width = self.progress.winfo_width()
        if width < 2:
            width = 680
        fill_width = int((value / 100) * width)
        self.progress.create_rectangle(0, 0, fill_width, 20, fill="green", outline="")
        self.progress.create_text(width/2, 10, text=f"{value}%", fill="black")
        self.root.update()
    
    def convert(self):
        if not self.selected_file:
            messagebox.showerror("오류", " 음성 파일을 선택해주세요")
            return
        
        # 스레드에서 실행
        thread = threading.Thread(target=self._convert_thread)
        thread.start()
    
    def _convert_thread(self):
        try:
            meeting_info = {
                '회의명': self.entry_제목.get() or '정기 회의',
                '장소': self.entry_장소.get() or '미정',
                '일시': self.entry_일시.get(),
                '작성자': self.entry_작성자.get() or '미정',
                '참석자': self.entry_참석자.get() or '미정',
            }
            
            output = self.entry_출력.get()
            if not output.endswith('.docx'):
                output += '.docx'
            
            self.label_상태.config(text="상태: 음성 변환 중...", fg="orange")
            self.update_progress(20)
            self.log("1단계: 음성 변환 중...")
            
            from speech_to_text import SpeechToText
            self.log("Whisper 모델 로딩...")
            self.update_progress(30)
            converter = SpeechToText(model="base")
            
            self.log("음성 파일 변환 중... (시간이 소요될 수 있습니다)")
            transcript = converter.convert_m4a_to_text(self.selected_file)
            
            self.update_progress(70)
            self.log("2단계: Word 문서 생성 중...")
            
            from document_generator import generate_meeting_minutes
            generate_meeting_minutes(meeting_info, transcript, output)
            
            self.update_progress(100)
            self.label_상태.config(text="상태: ✅ 완료!", fg="green")
            
            self.generated_file = os.path.abspath(output)
            self.log(f"✅ 성공! 파일: {self.generated_file}")
            
            self.btn_폴더.config(state=tk.NORMAL)
            self.btn_파일.config(state=tk.NORMAL)
            
        except Exception as e:
            self.label_상태.config(text=f"상태: ❌ 오류", fg="red")
            self.log(f"❌ 오류: {str(e)}")
            messagebox.showerror("오류", f"오류 발생:\n{str(e)}")
    
    def open_folder(self):
        if self.generated_file:
            folder = os.path.dirname(self.generated_file)
            if sys.platform == 'win32':
                os.startfile(folder)
            self.log("폴더 열기")
    
    def open_file(self):
        if self.generated_file and os.path.exists(self.generated_file):
            if sys.platform == 'win32':
                os.startfile(self.generated_file)
            self.log("파일 열기")
    
    def reset(self):
        self.entry_제목.delete(0, tk.END)
        self.entry_제목.insert(0, "정기 회의")
        self.entry_장소.delete(0, tk.END)
        self.entry_장소.insert(0, "미정")
        self.entry_일시.delete(0, tk.END)
        self.entry_일시.insert(0, datetime.now().strftime("%Y.%m.%d"))
        self.entry_작성자.delete(0, tk.END)
        self.entry_작성자.insert(0, "미정")
        self.entry_참석자.delete(0, tk.END)
        self.entry_참석자.insert(0, "미정")
        self.label_파일.config(text="선택 안 함", fg="gray")
        self.entry_출력.delete(0, tk.END)
        self.entry_출력.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
        self.update_progress(0)
        self.label_상태.config(text="상태: 준비 완료", fg="blue")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.btn_폴더.config(state=tk.DISABLED)
        self.btn_파일.config(state=tk.DISABLED)
        self.selected_file = None
        self.generated_file = None
        self.log("초기화됨")

if __name__ == '__main__':
    root = tk.Tk()
    app = MeetingRecorderApp(root)
    root.mainloop()
