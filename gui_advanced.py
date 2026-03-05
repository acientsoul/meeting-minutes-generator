# -*- coding: utf-8 -*-
"""
회의록 자동 생성 GUI (개선 버전)
- 드래그앤드롭 영역
- 미리보기 팝업
- 두 배치 인터페이스
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import sys
from datetime import datetime
import threading
from PIL import Image, ImageDraw, ImageTk
import io

class PreviewWindow:
    """미리보기 팝업 윈도우"""
    def __init__(self, parent, title, content):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("600x700")
        
        # 제목
        tk.Label(self.window, text=f"📄 {title}", 
                font=("Arial", 12, "bold")).pack(pady=10)
        
        # 분류 표시
        info_frame = tk.Frame(self.window)
        info_frame.pack(fill=tk.X, padx=10)
        
        # 미리보기 내용
        text_widget = scrolledtext.ScrolledText(self.window, 
                                               height=30, width=70,
                                               wrap=tk.WORD)
        text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)
        
        # 버튼
        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="📋 복사", 
                 command=lambda: self.copy_text(text_widget), 
                 width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="💾 저장", 
                 command=lambda: self.save_text(content, title), 
                 width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="닫기", 
                 command=self.window.destroy, 
                 width=10).pack(side=tk.LEFT, padx=5)
    
    def copy_text(self, text_widget):
        text_widget.config(state=tk.NORMAL)
        content = text_widget.get(1.0, tk.END)
        self.window.clipboard_clear()
        self.window.clipboard_append(content)
        self.window.update()
        messagebox.showinfo("완료", "텍스트가 클립보드에 복사되었습니다!")
        text_widget.config(state=tk.DISABLED)
    
    def save_text(self, content, title):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
            initialfile=f"{title}.txt"
        )
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("완료", f"파일이 저장되었습니다!\n{file_path}")

class DropZone(tk.Frame):
    """드래그앤드롭 영역"""
    def __init__(self, parent, on_drop, **kwargs):
        super().__init__(parent, **kwargs)
        self.on_drop = on_drop
        self.file_path = None
        
        # 드래그 오버 효과
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Button-1>", self.on_click)
        
        # 배경 및 텍스트
        self.config(bg="#f0f0f0", relief=tk.RIDGE, bd=2)
        
        inner_frame = tk.Frame(self, bg="#f0f0f0")
        inner_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        
        tk.Label(inner_frame, text="📁", font=("Arial", 40), 
                bg="#f0f0f0").pack()
        tk.Label(inner_frame, text="여기에 m4a 파일을 드래그해주세요",
                font=("Arial", 12, "bold"), bg="#f0f0f0").pack()
        tk.Label(inner_frame, text="또는 클릭하여 파일 선택",
                font=("Arial", 10), fg="gray", bg="#f0f0f0").pack()
        
        self.label_info = tk.Label(inner_frame, text="", 
                                   font=("Arial", 9), fg="green", bg="#f0f0f0")
        self.label_info.pack(pady=10)
    
    def on_enter(self, event):
        self.config(bg="#e0e0ff")
    
    def on_leave(self, event):
        self.config(bg="#f0f0f0")
    
    def on_click(self, event):
        file = filedialog.askopenfilename(filetypes=[("M4A Files", "*.m4a")])
        if file:
            self.set_file(file)
    
    def set_file(self, file_path):
        self.file_path = file_path
        size = os.path.getsize(file_path) / (1024 * 1024)
        info = f"✓ {os.path.basename(file_path)} ({size:.1f}MB)"
        self.label_info.config(text=info, fg="green")
        self.on_drop(file_path)

class MeetingRecorderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎤 회의록 자동 생성")
        self.root.geometry("900x950")
        self.generated_file = None
        
        # 메인 프레임 (왼쪽: 입력, 오른쪽: 미리보기)
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ==================== 왼쪽 패널 ====================
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Label(left_frame, text="🎤 회의록 자동 생성", 
                font=("Arial", 14, "bold")).pack(pady=10)
        tk.Frame(left_frame, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X)
        
        # 회의 정보
        info_frame = tk.Frame(left_frame)
        info_frame.pack(pady=10, fill=tk.X)
        
        tk.Label(info_frame, text="회의명:", width=10).grid(row=0, column=0, sticky=tk.W)
        self.entry_제목 = tk.Entry(info_frame, width=30)
        self.entry_제목.insert(0, "정기 회의")
        self.entry_제목.grid(row=0, column=1, sticky=tk.W)
        
        tk.Label(info_frame, text="장소:", width=10).grid(row=1, column=0, sticky=tk.W)
        self.entry_장소 = tk.Entry(info_frame, width=30)
        self.entry_장소.insert(0, "미정")
        self.entry_장소.grid(row=1, column=1, sticky=tk.W)
        
        tk.Label(info_frame, text="일시:", width=10).grid(row=2, column=0, sticky=tk.W)
        self.entry_일시 = tk.Entry(info_frame, width=30)
        self.entry_일시.insert(0, datetime.now().strftime("%Y.%m.%d"))
        self.entry_일시.grid(row=2, column=1, sticky=tk.W)
        
        tk.Label(info_frame, text="작성자:", width=10).grid(row=3, column=0, sticky=tk.W)
        self.entry_작성자 = tk.Entry(info_frame, width=30)
        self.entry_작성자.insert(0, "미정")
        self.entry_작성자.grid(row=3, column=1, sticky=tk.W)
        
        tk.Label(info_frame, text="참석자:", width=10).grid(row=4, column=0, sticky=tk.W)
        self.entry_참석자 = tk.Entry(info_frame, width=30)
        self.entry_참석자.insert(0, "미정")
        self.entry_참석자.grid(row=4, column=1, sticky=tk.W)
        
        tk.Frame(left_frame, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, pady=5)
        
        # 드래그앤드롭 영역
        tk.Label(left_frame, text="🎵 음성 파일 선택", font=("Arial", 10, "bold")).pack()
        self.drop_zone = DropZone(left_frame, self.on_file_selected, height=100)
        self.drop_zone.pack(fill=tk.BOTH, expand=True, pady=5)
        
        tk.Frame(left_frame, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, pady=5)
        
        # 출력 파일
        out_frame = tk.Frame(left_frame)
        out_frame.pack(pady=5, fill=tk.X)
        tk.Label(out_frame, text="출력파일:", width=10).pack(side=tk.LEFT)
        self.entry_출력 = tk.Entry(out_frame)
        self.entry_출력.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
        self.entry_출력.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        tk.Frame(left_frame, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, pady=5)
        
        # 진행률
        self.progress = tk.Canvas(left_frame, height=20, bg="lightgray")
        self.progress.pack(pady=5, fill=tk.X)
        
        self.label_상태 = tk.Label(left_frame, text="상태: 준비 완료", fg="blue")
        self.label_상태.pack()
        
        # 로그
        self.log_text = scrolledtext.ScrolledText(left_frame, height=6, width=50)
        self.log_text.pack(pady=5, fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)
        
        # 버튼
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="🚀 변환 시작", command=self.convert, 
                 width=12, bg="green", fg="white").pack(side=tk.LEFT, padx=3)
        self.btn_미리보기 = tk.Button(btn_frame, text="👁 미리보기", 
                                     command=self.show_preview, width=12, 
                                     state=tk.DISABLED)
        self.btn_미리보기.pack(side=tk.LEFT, padx=3)
        self.btn_폴더 = tk.Button(btn_frame, text="📂 폴더", command=self.open_folder, 
                                 width=12, state=tk.DISABLED)
        self.btn_폴더.pack(side=tk.LEFT, padx=3)
        self.btn_파일 = tk.Button(btn_frame, text="📄 파일", command=self.open_file, 
                                 width=12, state=tk.DISABLED)
        self.btn_파일.pack(side=tk.LEFT, padx=3)
        
        btn_frame2 = tk.Frame(left_frame)
        btn_frame2.pack(pady=5)
        tk.Button(btn_frame2, text="🔄 초기화", command=self.reset, width=12).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame2, text="❌ 종료", command=root.quit, width=12).pack(side=tk.LEFT, padx=3)
        
        # ==================== 오른쪽 패널 (미리보기) ====================
        right_frame = tk.Frame(main_frame, bg="white", relief=tk.SUNKEN, bd=1)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Label(right_frame, text="📋 미리보기", font=("Arial", 12, "bold"), 
                bg="white").pack(pady=10)
        
        self.preview_text = scrolledtext.ScrolledText(right_frame, height=40, width=40,
                                                      bg="#f9f9f9", wrap=tk.WORD)
        self.preview_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.preview_text.config(state=tk.DISABLED)
        
        self.selected_file = None
        self.transcript = None
        self.log("프로그램 시작됨\n1️⃣ 회의 정보 입력\n2️⃣ 음성 파일 선택\n3️⃣ '변환 시작' 클릭")
    
    def on_file_selected(self, file_path):
        self.selected_file = file_path
        self.log(f"파일 선택됨: {os.path.basename(file_path)}")
    
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
        self.progress.create_text(width/2, 10, text=f"{value}%", fill="white", font=("Arial", 10, "bold"))
        self.root.update()
    
    def convert(self):
        if not self.selected_file:
            messagebox.showerror("오류", "음성 파일을 선택해주세요")
            return
        
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
            self.transcript = converter.convert_m4a_to_text(self.selected_file)
            
            # 미리보기 업데이트
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)
            preview_content = f"【회의 정보】\n"
            preview_content += f"회의명: {meeting_info['회의명']}\n"
            preview_content += f"장소: {meeting_info['장소']}\n"
            preview_content += f"일시: {meeting_info['일시']}\n"
            preview_content += f"작성자: {meeting_info['작성자']}\n"
            preview_content += f"참석자: {meeting_info['참석자']}\n"
            preview_content += f"\n【회의 내용】\n"
            preview_content += self.transcript
            self.preview_text.insert(tk.END, preview_content)
            self.preview_text.config(state=tk.DISABLED)
            
            self.update_progress(70)
            self.log("2단계: Word 문서 생성 중...")
            
            from document_generator import generate_meeting_minutes
            generate_meeting_minutes(meeting_info, self.transcript, output)
            
            self.update_progress(100)
            self.label_상태.config(text="상태: ✅ 완료!", fg="green")
            
            self.generated_file = os.path.abspath(output)
            self.log(f"✅ 성공! 파일: {self.generated_file}")
            
            self.btn_폴더.config(state=tk.NORMAL)
            self.btn_파일.config(state=tk.NORMAL)
            self.btn_미리보기.config(state=tk.NORMAL)
            
        except Exception as e:
            self.label_상태.config(text=f"상태: ❌ 오류", fg="red")
            self.log(f"❌ 오류: {str(e)}")
            messagebox.showerror("오류", f"오류 발생:\n{str(e)}")
    
    def show_preview(self):
        if self.transcript:
            meeting_info = {
                '회의명': self.entry_제목.get(),
                '장소': self.entry_장소.get(),
                '일시': self.entry_일시.get(),
                '작성자': self.entry_작성자.get(),
                '참석자': self.entry_참석자.get(),
            }
            
            content = f"【회의 정보】\n"
            content += f"회의명: {meeting_info['회의명']}\n"
            content += f"장소: {meeting_info['장소']}\n"
            content += f"일시: {meeting_info['일시']}\n"
            content += f"작성자: {meeting_info['작성자']}\n"
            content += f"참석자: {meeting_info['참석자']}\n\n"
            content += f"【회의 내용】\n"
            content += self.transcript
            
            PreviewWindow(self.root, "회의록 미리보기", content)
    
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
        self.drop_zone.label_info.config(text="")
        self.entry_출력.delete(0, tk.END)
        self.entry_출력.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
        self.update_progress(0)
        self.label_상태.config(text="상태: 준비 완료", fg="blue")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.config(state=tk.DISABLED)
        self.btn_폴더.config(state=tk.DISABLED)
        self.btn_파일.config(state=tk.DISABLED)
        self.btn_미리보기.config(state=tk.DISABLED)
        self.selected_file = None
        self.transcript = None
        self.generated_file = None
        self.log("초기화됨")

if __name__ == '__main__':
    root = tk.Tk()
    app = MeetingRecorderApp(root)
    root.mainloop()
