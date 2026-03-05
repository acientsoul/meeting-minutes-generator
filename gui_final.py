# -*- coding: utf-8 -*-
"""
회의록 자동 생성 GUI (개선 버전)
- 드래그앤드롭 영역
- 미리보기 팝업
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import sys
from datetime import datetime
import threading

class PreviewWindow:
    """미리보기 팝업 윈도우"""
    def __init__(self, parent, title, content):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("600x700")
        
        tk.Label(self.window, text=f"📄 {title}", 
                font=("Arial", 12, "bold")).pack(pady=10)
        
        text_widget = scrolledtext.ScrolledText(self.window, height=30, width=70, wrap=tk.WORD)
        text_widget.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)
        
        btn_frame = tk.Frame(self.window)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="📋 복사", command=lambda: self.copy(text_widget), width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="💾 저장", command=lambda: self.save(content, title), width=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="X", command=self.window.destroy, width=10).pack(side=tk.LEFT, padx=5)
    
    def copy(self, tw):
        tw.config(state=tk.NORMAL)
        content = tw.get(1.0, tk.END)
        self.window.clipboard_clear()
        self.window.clipboard_append(content)
        self.window.update()
        messagebox.showinfo("완료", "클립보드 복사됨")
        tw.config(state=tk.DISABLED)
    
    def save(self, content, title):
        file = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"{title}.txt")
        if file:
            with open(file, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("완료", f"저장됨: {file}")

class DropFrame(tk.Frame):
    """드래그앤드롭 영역"""
    def __init__(self, parent, callback, **kwargs):
        super().__init__(parent, **kwargs)
        self.callback = callback
        self.file_path = None
        self.in_drag = False
        
        self.config(bg="lightgray", relief=tk.RIDGE, bd=3)
        self.bind("<Button-1>", self.on_click)
        self.bind("<Motion>", self.on_motion)
        self.bind("<Leave>", self.on_leave)
        
        # 파일 선택 전 상태
        self.empty_inner = tk.Frame(self, bg="lightgray")
        tk.Label(self.empty_inner, text="📁", font=("Arial", 50), bg="lightgray").pack()
        tk.Label(self.empty_inner, text="여기에 m4a 파일을 드래그", font=("Arial", 12, "bold"), bg="lightgray").pack()
        tk.Label(self.empty_inner, text="또는 클릭하여 선택", font=("Arial", 10), fg="gray", bg="lightgray").pack()
        self.info_label = tk.Label(self.empty_inner, text="", font=("Arial", 9, "bold"), fg="green", bg="lightgray")
        self.info_label.pack(pady=10)
        
        # 파일 선택 후 상태
        self.file_inner = tk.Frame(self, bg="#e8f5e9")
        tk.Label(self.file_inner, text="✓", font=("Arial", 60), fg="green", bg="#e8f5e9").pack()
        tk.Label(self.file_inner, text="파일이 선택되었습니다", font=("Arial", 12, "bold"), fg="green", bg="#e8f5e9").pack()
        self.file_info_label = tk.Label(self.file_inner, text="", font=("Arial", 10, "bold"), fg="darkgreen", bg="#e8f5e9")
        self.file_info_label.pack(pady=5)
        self.file_size_label = tk.Label(self.file_inner, text="", font=("Arial", 9), fg="gray", bg="#e8f5e9")
        self.file_size_label.pack()
        tk.Label(self.file_inner, text="다시 클릭하여 다른 파일 선택", font=("Arial", 9), fg="gray", bg="#e8f5e9").pack(pady=5)
        
        # 초기 상태 (파일 선택 전)
        self.empty_inner.pack(expand=True, fill=tk.BOTH)
        self.current_state = "empty"
    
    def on_motion(self, e):
        if self.current_state == "empty":
            self.config(bg="#e0e0ff")
    
    def on_leave(self, e):
        if self.current_state == "empty":
            self.config(bg="lightgray")
    
    def on_click(self, e):
        file = filedialog.askopenfilename(filetypes=[("M4A Files", "*.m4a")])
        if file:
            self.set_file(file)
    
    def set_file(self, path):
        self.file_path = path
        size = os.path.getsize(path) / (1024 * 1024)
        
        # UI 전환 (empty → file 상태)
        if self.current_state == "empty":
            self.empty_inner.pack_forget()
            self.current_state = "file"
            self.file_inner.pack(expand=True, fill=tk.BOTH)
            self.config(bg="#e8f5e9")
        
        # 파일 정보 업데이트
        self.file_info_label.config(text=f"📄 {os.path.basename(path)}")
        self.file_size_label.config(text=f"크기: {size:.2f} MB")
        
        self.callback(path)
    
    def reset(self):
        """초기 상태로 리셋"""
        if self.current_state == "file":
            self.file_inner.pack_forget()
            self.current_state = "empty"
            self.empty_inner.pack(expand=True, fill=tk.BOTH)
            self.config(bg="lightgray")
        self.file_path = None

class MeetingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎤 회의록 자동 생성")
        self.root.geometry("1000x900")
        self.generated_file = None
        self.transcript = None
        
        # 메인 프레임
        main = tk.Frame(root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 왼쪽 (입력)
        left = tk.Frame(main)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(left, text="🎤 회의록 자동 생성", font=("Arial", 14, "bold")).pack(pady=10)
        tk.Frame(left, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X)
        
        # 정보 입력
        info = tk.Frame(left)
        info.pack(pady=10, fill=tk.X)
        
        fields = [("회의명:", "제목"), ("장소:", "장소"), ("일시:", "일시"), ("작성자:", "작성자"), ("참석자:", "참석자")]
        defaults = ["정기 회의", "미정", datetime.now().strftime("%Y.%m.%d"), "미정", "미정"]
        
        self.entries = {}
        for i, (label, key) in enumerate(fields):
            tk.Label(info, text=label, width=8).grid(row=i, column=0, sticky=tk.W)
            entry = tk.Entry(info, width=35)
            entry.insert(0, defaults[i])
            entry.grid(row=i, column=1, sticky=tk.W, padx=5)
            self.entries[key] = entry
        
        tk.Frame(left, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, pady=5)
        
        # 파일 드롭 영역
        tk.Label(left, text="🎵 음성파일 선택", font=("Arial", 10, "bold")).pack()
        self.drop_area = DropFrame(left, self.on_file_selected, height=80)
        self.drop_area.pack(fill=tk.BOTH, expand=True, pady=5)
        
        tk.Frame(left, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X, pady=5)
        
        # 출력파일
        out_frm = tk.Frame(left)
        out_frm.pack(pady=5, fill=tk.X)
        tk.Label(out_frm, text="출력파일:", width=8).pack(side=tk.LEFT)
        self.entry_out = tk.Entry(out_frm)
        self.entry_out.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
        self.entry_out.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 진행률
        self.progress = tk.Canvas(left, height=20, bg="lightgray")
        self.progress.pack(pady=5, fill=tk.X)
        
        self.status_label = tk.Label(left, text="상태: 준비 완료", fg="blue")
        self.status_label.pack()
        
        # 로그
        self.log_area = scrolledtext.ScrolledText(left, height=6, width=50)
        self.log_area.pack(pady=5, fill=tk.BOTH, expand=True)
        self.log_area.config(state=tk.DISABLED)
        
        # 버튼
        btn1 = tk.Frame(left)
        btn1.pack(pady=10)
        tk.Button(btn1, text="🚀 변환 시작", command=self.convert, width=12, bg="green", fg="white").pack(side=tk.LEFT, padx=3)
        self.btn_preview = tk.Button(btn1, text="👁 미리보기", command=self.preview, width=12, state=tk.DISABLED)
        self.btn_preview.pack(side=tk.LEFT, padx=3)
        self.btn_folder = tk.Button(btn1, text="📂 폴더", command=self.open_folder, width=12, state=tk.DISABLED)
        self.btn_folder.pack(side=tk.LEFT, padx=3)
        self.btn_file = tk.Button(btn1, text="📄 파일", command=self.open_file, width=12, state=tk.DISABLED)
        self.btn_file.pack(side=tk.LEFT, padx=3)
        
        btn2 = tk.Frame(left)
        btn2.pack(pady=5)
        tk.Button(btn2, text="🔄 초기화", command=self.reset, width=12).pack(side=tk.LEFT, padx=3)
        tk.Button(btn2, text="❌ 종료", command=root.quit, width=12).pack(side=tk.LEFT, padx=3)
        
        # 오른쪽 (미리보기)
        right = tk.Frame(main, bg="white", relief=tk.SUNKEN, bd=1)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(right, text="📋 미리보기", font=("Arial", 12, "bold"), bg="white").pack(pady=10)
        self.preview_area = scrolledtext.ScrolledText(right, height=40, width=40, bg="#f9f9f9", wrap=tk.WORD)
        self.preview_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.preview_area.config(state=tk.DISABLED)
        
        self.selected_file = None
        self.log("프로그램 시작\n1. 회의정보 입력\n2. 음성파일 선택\n3. 변환 시작")
    
    def on_file_selected(self, path):
        self.selected_file = path
        self.log(f"파일 선택: {os.path.basename(path)}")
    
    def log(self, msg):
        self.log_area.config(state=tk.NORMAL)
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.root.update()
    
    def progress_update(self, val):
        self.progress.delete("all")
        w = self.progress.winfo_width()
        if w < 2: w = 680
        fw = int((val / 100) * w)
        self.progress.create_rectangle(0, 0, fw, 20, fill="green", outline="")
        self.progress.create_text(w/2, 10, text=f"{val}%", fill="white", font=("Arial", 10, "bold"))
        self.root.update()
    
    def convert(self):
        if not self.selected_file:
            messagebox.showerror("오류", "음성 파일을 선택해주세요")
            return
        threading.Thread(target=self.convert_thread).start()
    
    def convert_thread(self):
        try:
            info = {
                '회의명': self.entries['제목'].get() or '정기 회의',
                '장소': self.entries['장소'].get() or '미정',
                '일시': self.entries['일시'].get(),
                '작성자': self.entries['작성자'].get() or '미정',
                '참석자': self.entries['참석자'].get() or '미정',
            }
            
            out = self.entry_out.get()
            if not out.endswith('.docx'): out += '.docx'
            
            self.status_label.config(text="상태: 음성 변환 중...", fg="orange")
            self.progress_update(20)
            self.log("1단계: 음성 변환 중...")
            
            from speech_to_text import SpeechToText
            self.log("Whisper 로딩...")
            self.progress_update(30)
            converter = SpeechToText(model="base")
            
            self.log("변환 중... (시간이 소요될 수 있습니다)")
            self.transcript = converter.convert_m4a_to_text(self.selected_file)
            
            # 미리보기 업데이트
            preview = "【회의정보】\n"
            for k, v in info.items():
                preview += f"{k}: {v}\n"
            preview += f"\n【회의내용】\n{self.transcript}"
            
            self.preview_area.config(state=tk.NORMAL)
            self.preview_area.delete(1.0, tk.END)
            self.preview_area.insert(tk.END, preview)
            self.preview_area.config(state=tk.DISABLED)
            
            self.progress_update(70)
            self.log("2단계: Word 생성 중...")
            
            from document_generator import generate_meeting_minutes
            generate_meeting_minutes(info, self.transcript, out)
            
            self.progress_update(100)
            self.status_label.config(text="상태: ✅ 완료!", fg="green")
            
            self.generated_file = os.path.abspath(out)
            self.log(f"성공! {self.generated_file}")
            
            self.btn_folder.config(state=tk.NORMAL)
            self.btn_file.config(state=tk.NORMAL)
            self.btn_preview.config(state=tk.NORMAL)
            
        except Exception as e:
            self.status_label.config(text="상태: ❌ 오류", fg="red")
            self.log(f"❌ {str(e)}")
            messagebox.showerror("오류", f"{str(e)}")
    
    def preview(self):
        if self.transcript:
            info = "\n".join([f"{k}: {self.entries[k].get()}" if k in ['제목','장소','일시','작성자','참석자'] else '' 
                            for k in ['제목','장소','일시','작성자','참석자']])
            content = f"【회의정보】\n{info}\n\n【회의내용】\n{self.transcript}"
            PreviewWindow(self.root, "회의록 미리보기", content)
    
    def open_folder(self):
        if self.generated_file:
            os.startfile(os.path.dirname(self.generated_file))
            self.log("폴더 열기")
    
    def open_file(self):
        if self.generated_file and os.path.exists(self.generated_file):
            os.startfile(self.generated_file)
            self.log("파일 열기")
    
    def reset(self):
        for k, e in [("제목","정기 회의"),("장소","미정"),("일시",datetime.now().strftime("%Y.%m.%d")),("작성자","미정"),("참석자","미정")]:
            self.entries[k].delete(0, tk.END)
            self.entries[k].insert(0, e)
        self.drop_area.reset()
        self.entry_out.delete(0, tk.END)
        self.entry_out.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
        self.progress_update(0)
        self.status_label.config(text="상태: 준비 완료", fg="blue")
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.preview_area.config(state=tk.NORMAL)
        self.preview_area.delete(1.0, tk.END)
        self.preview_area.config(state=tk.DISABLED)
        self.btn_folder.config(state=tk.DISABLED)
        self.btn_file.config(state=tk.DISABLED)
        self.btn_preview.config(state=tk.DISABLED)
        self.selected_file = None
        self.transcript = None
        self.generated_file = None
        self.log("초기화됨")

if __name__ == '__main__':
    root = tk.Tk()
    app = MeetingApp(root)
    root.mainloop()
