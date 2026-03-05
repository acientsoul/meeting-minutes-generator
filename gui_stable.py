# -*- coding: utf-8 -*-
"""
회의록 자동 생성 GUI (안정 버전)
- 파일 선택 버튼
- 드래그앤드롭 기능
- 오른쪽 미리보기
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import os
import sys
import webbrowser
from datetime import datetime
import threading

# AI 회의록 생성 모듈
try:
    from ai_meeting_generator import (
        load_config, save_config, generate_ai_meeting_minutes,
        get_available_providers, GEMINI_AVAILABLE, OPENAI_AVAILABLE
    )
    AI_AVAILABLE = True
except ImportError:
    AI_AVAILABLE = False

# 드래그앤드롭 지원
try:
    import tkinterdnd2 as tkdnd
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

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
        messagebox.showinfo("완료", "클립보드에 복사됨")
        tw.config(state=tk.DISABLED)
    
    def save(self, content, title):
        file = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"{title}.txt")
        if file:
            with open(file, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("완료", f"저장됨: {file}")

class AISettingsWindow:
    """AI 설정 팝업 윈도우"""
    def __init__(self, parent, callback=None):
        self.window = tk.Toplevel(parent)
        self.window.title("🤖 AI 설정")
        self.window.geometry("550x680")
        self.window.resizable(False, False)
        self.callback = callback
        
        config = load_config() if AI_AVAILABLE else {}
        
        # 스크롤 가능한 영역
        canvas = tk.Canvas(self.window)
        scrollbar = tk.Scrollbar(self.window, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        tk.Label(scroll_frame, text="🤖 AI 회의록 생성 설정",
                font=("Arial", 14, "bold")).pack(pady=10)
        
        # AI 제공자 선택
        provider_frame = tk.LabelFrame(scroll_frame, text="AI 제공자 선택", font=("Arial", 10, "bold"), padx=15, pady=10)
        provider_frame.pack(fill=tk.X, padx=20, pady=5)
        
        self.provider_var = tk.StringVar(value=config.get("ai_provider", "gemini"))
        tk.Radiobutton(provider_frame, text="Google Gemini (무료 tier 가능)",
                       variable=self.provider_var, value="gemini",
                       font=("Arial", 10)).pack(anchor=tk.W)
        tk.Radiobutton(provider_frame, text="OpenAI GPT (유료)",
                       variable=self.provider_var, value="openai",
                       font=("Arial", 10)).pack(anchor=tk.W)
        
        # Gemini 설정
        gemini_frame = tk.LabelFrame(scroll_frame, text="Google Gemini 설정", font=("Arial", 10, "bold"), padx=15, pady=10)
        gemini_frame.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(gemini_frame, text="API Key (여러 개 입력 시 한 줄에 하나씩):", font=("Arial", 10)).pack(anchor=tk.W, pady=(3,0))
        self.gemini_keys_text = tk.Text(gemini_frame, width=50, height=4, font=("Arial", 10))
        self.gemini_keys_text.pack(fill=tk.X, pady=3)
        # 기존 키 로드 (다중 키 + 단일 키 합침)
        gemini_keys = config.get("gemini_api_keys", [])
        single_key = config.get("gemini_api_key", "")
        if single_key and single_key not in gemini_keys:
            gemini_keys.insert(0, single_key)
        if gemini_keys:
            self.gemini_keys_text.insert("1.0", "\n".join(gemini_keys))
        
        tk.Label(gemini_frame, text="💡 할당량 초과 시 자동으로 다음 키로 전환됩니다", 
                font=("Arial", 9), fg="gray").pack(anchor=tk.W)
        
        tk.Label(gemini_frame, text="모델:", font=("Arial", 10)).pack(anchor=tk.W, pady=(5,0))
        self.gemini_model = ttk.Combobox(gemini_frame, values=["gemini-2.0-flash", "gemini-2.5-flash", "gemini-2.5-pro"], width=42)
        self.gemini_model.set(config.get("gemini_model", "gemini-2.0-flash"))
        self.gemini_model.pack(anchor=tk.W, pady=3)
        
        gemini_link = tk.Label(gemini_frame, text="🔗 API 키 발급받기 (aistudio.google.com)",
                font=("Arial", 9), fg="blue", cursor="hand2")
        gemini_link.pack(anchor=tk.W, pady=3)
        gemini_link.bind("<Button-1>", lambda e: webbrowser.open("https://aistudio.google.com/apikey"))
        gemini_link.bind("<Enter>", lambda e: gemini_link.config(font=("Arial", 9, "underline")))
        gemini_link.bind("<Leave>", lambda e: gemini_link.config(font=("Arial", 9)))
        
        # OpenAI 설정
        openai_frame = tk.LabelFrame(scroll_frame, text="OpenAI 설정", font=("Arial", 10, "bold"), padx=15, pady=10)
        openai_frame.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(openai_frame, text="API Key (여러 개 입력 시 한 줄에 하나씩):", font=("Arial", 10)).pack(anchor=tk.W, pady=(3,0))
        self.openai_keys_text = tk.Text(openai_frame, width=50, height=4, font=("Arial", 10))
        self.openai_keys_text.pack(fill=tk.X, pady=3)
        openai_keys = config.get("openai_api_keys", [])
        single_okey = config.get("openai_api_key", "")
        if single_okey and single_okey not in openai_keys:
            openai_keys.insert(0, single_okey)
        if openai_keys:
            self.openai_keys_text.insert("1.0", "\n".join(openai_keys))
        
        tk.Label(openai_frame, text="💡 할당량 초과 시 자동으로 다음 키로 전환됩니다", 
                font=("Arial", 9), fg="gray").pack(anchor=tk.W)
        
        tk.Label(openai_frame, text="모델:", font=("Arial", 10)).pack(anchor=tk.W, pady=(5,0))
        self.openai_model = ttk.Combobox(openai_frame, values=["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini", "gpt-4.1"], width=42)
        self.openai_model.set(config.get("openai_model", "gpt-4o-mini"))
        self.openai_model.pack(anchor=tk.W, pady=3)
        
        openai_link = tk.Label(openai_frame, text="🔗 API 키 발급받기 (platform.openai.com)",
                font=("Arial", 9), fg="blue", cursor="hand2")
        openai_link.pack(anchor=tk.W, pady=3)
        openai_link.bind("<Button-1>", lambda e: webbrowser.open("https://platform.openai.com/api-keys"))
        openai_link.bind("<Enter>", lambda e: openai_link.config(font=("Arial", 9, "underline")))
        openai_link.bind("<Leave>", lambda e: openai_link.config(font=("Arial", 9)))
        
        # 버튼
        btn_frame = tk.Frame(scroll_frame)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="💾 저장", command=self.save_settings,
                 font=("Arial", 11, "bold"), bg="#4CAF50", fg="white",
                 padx=20, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="🔑 키 테스트", command=self.test_keys,
                 font=("Arial", 11), bg="#2196F3", fg="white",
                 padx=20, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="취소", command=self.window.destroy,
                 font=("Arial", 11), padx=20, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=10)
    
    def _parse_keys(self, text_widget):
        """텍스트 위젯에서 키 목록 파싱 (빈 줄 제거)"""
        raw = text_widget.get("1.0", tk.END).strip()
        return [k.strip() for k in raw.split("\n") if k.strip()]
    
    def test_keys(self):
        """등록된 키 유효성 테스트"""
        provider = self.provider_var.get()
        if provider == "gemini":
            keys = self._parse_keys(self.gemini_keys_text)
            model = self.gemini_model.get()
        else:
            keys = self._parse_keys(self.openai_keys_text)
            model = self.openai_model.get()
        
        if not keys:
            messagebox.showwarning("알림", "테스트할 API 키가 없습니다.", parent=self.window)
            return
        
        results = []
        for i, key in enumerate(keys):
            masked = key[:8] + "..." + key[-4:] if len(key) > 12 else key
            try:
                if provider == "gemini" and GEMINI_AVAILABLE:
                    from google import genai as _genai
                    client = _genai.Client(api_key=key)
                    client.models.generate_content(model=model, contents="Hi")
                    results.append(f"키 #{i+1} ({masked}): ✅ 정상")
                elif provider == "openai":
                    import openai as _openai
                    client = _openai.OpenAI(api_key=key)
                    client.chat.completions.create(
                        model=model,
                        messages=[{"role": "user", "content": "Hi"}],
                        max_tokens=5
                    )
                    results.append(f"키 #{i+1} ({masked}): ✅ 정상")
                else:
                    results.append(f"키 #{i+1} ({masked}): ⚠ 라이브러리 미설치")
            except Exception as e:
                err = str(e)[:120]
                if "429" in err or "RESOURCE_EXHAUSTED" in err:
                    results.append(f"키 #{i+1} ({masked}): ⚠ 할당량 초과")
                elif "401" in err or "403" in err or "400" in err or "INVALID" in err.upper() or "API_KEY" in err.upper():
                    results.append(f"키 #{i+1} ({masked}): ❌ 키 무효/만료")
                else:
                    results.append(f"키 #{i+1} ({masked}): ❌ {err[:50]}")
        
        messagebox.showinfo("키 테스트 결과", "\n".join(results), parent=self.window)
    
    def save_settings(self):
        gemini_keys = self._parse_keys(self.gemini_keys_text)
        openai_keys = self._parse_keys(self.openai_keys_text)
        
        config = {
            "ai_provider": self.provider_var.get(),
            "gemini_api_key": gemini_keys[0] if gemini_keys else "",
            "gemini_api_keys": gemini_keys,
            "openai_api_key": openai_keys[0] if openai_keys else "",
            "openai_api_keys": openai_keys,
            "gemini_model": self.gemini_model.get(),
            "openai_model": self.openai_model.get(),
        }
        save_config(config)
        
        key_count = len(gemini_keys) if self.provider_var.get() == "gemini" else len(openai_keys)
        messagebox.showinfo("완료", f"AI 설정이 저장되었습니다.\n등록된 키: {key_count}개", parent=self.window)
        if self.callback:
            self.callback(config)
        self.window.destroy()


class MeetingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎤 회의록 자동 생성 (AI 지원)")
        self.root.geometry("1050x900")
        self.generated_file = None
        self.transcript = None
        self.ai_minutes = None
        self.doc_content = None
        self.meeting_info = None
        self.selected_file = None
        
        # 메인 프레임
        main = tk.Frame(root)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ===== 왼쪽 패널 (입력) - 스크롤 가능 =====
        left_container = tk.Frame(main, width=480)
        left_container.pack(side=tk.LEFT, fill=tk.BOTH, padx=5)
        left_container.pack_propagate(False)
        
        left_canvas = tk.Canvas(left_container, highlightthickness=0)
        left_scrollbar = tk.Scrollbar(left_container, orient=tk.VERTICAL, command=left_canvas.yview)
        left = tk.Frame(left_canvas)
        
        left.bind("<Configure>", lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.create_window((0, 0), window=left, anchor="nw", width=460)
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        left_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        left_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 마우스 휠 스크롤 지원
        def _on_mousewheel(event):
            left_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        left_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 제목
        title_frame = tk.Frame(left, bg="#2196F3")
        title_frame.pack(fill=tk.X, pady=10)
        tk.Label(title_frame, text="🎤 회의록 자동 생성", font=("Arial", 14, "bold"), 
                bg="#2196F3", fg="white").pack(pady=10)
        
        # 회의 정보
        info_frame = tk.LabelFrame(left, text="📋 회의 정보", font=("Arial", 10, "bold"), padx=10, pady=10)
        info_frame.pack(fill=tk.X, pady=10)
        
        self.entries = {}
        fields = [("회의명:", "회의명"), ("장소:", "장소"), ("일시:", "일시"), ("작성자:", "작성자"), ("참석자:", "참석자"), ("업체이름:", "업체이름")]
        defaults = ["정기 회의", "미정", datetime.now().strftime("%Y.%m.%d"), "미정", "미정", "미정"]
        
        for i, (label, key) in enumerate(fields):
            tk.Label(info_frame, text=label, width=8, font=("Arial", 10)).grid(row=i, column=0, sticky=tk.W, pady=5)
            entry = tk.Entry(info_frame, width=35, font=("Arial", 10))
            entry.insert(0, defaults[i])
            entry.grid(row=i, column=1, sticky=tk.W, padx=5, pady=5)
            self.entries[key] = entry
        
        # 파일 선택 영역
        file_frame = tk.LabelFrame(left, text="🎵 음성 파일 선택", font=("Arial", 10, "bold"), padx=10, pady=5)
        file_frame.pack(fill=tk.X, pady=5)
        
        # 파일 선택 버튼이 있는 영역 (고정 높이)
        self.file_display = tk.Frame(file_frame, bg="lightgray", relief=tk.RIDGE, bd=2, height=120)
        self.file_display.pack(fill=tk.X, pady=5)
        self.file_display.pack_propagate(False)
        
        # 초기 상태 (파일 선택 전)
        self.empty_state = tk.Frame(self.file_display, bg="lightgray")
        tk.Label(self.empty_state, text="📁 여기에 m4a 파일을 선택해주세요", font=("Arial", 11, "bold"), bg="lightgray").pack(pady=15)
        tk.Label(self.empty_state, text="아래 버튼을 클릭하여 파일 선택", font=("Arial", 9), bg="lightgray", fg="gray").pack()
        self.empty_state.pack(fill=tk.BOTH, expand=True)
        
        # 파일 선택 후 상태
        self.file_state = tk.Frame(self.file_display, bg="#c8e6c9")
        tk.Label(self.file_state, text="✓ 파일이 선택되었습니다", font=("Arial", 11, "bold"), fg="green", bg="#c8e6c9").pack(pady=8)
        self.file_name_label = tk.Label(self.file_state, text="", font=("Arial", 10, "bold"), fg="darkgreen", bg="#c8e6c9")
        self.file_name_label.pack()
        self.file_size_label = tk.Label(self.file_state, text="", font=("Arial", 9), fg="gray", bg="#c8e6c9")
        self.file_size_label.pack()
        
        # 파일 선택 버튼
        btn_select = tk.Button(file_frame, text="📂 파일 선택", command=self.select_file, 
                              font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", pady=5, cursor="hand2")
        btn_select.pack(fill=tk.X, pady=3)
        
        # AI 설정 영역
        ai_frame = tk.LabelFrame(left, text="🤖 AI 회의록 생성", font=("Arial", 10, "bold"), padx=10, pady=5)
        ai_frame.pack(fill=tk.X, pady=5)
        
        ai_top = tk.Frame(ai_frame)
        ai_top.pack(fill=tk.X)
        
        self.ai_enabled = tk.BooleanVar(value=True)
        tk.Checkbutton(ai_top, text="AI로 회의록 자동 구성", variable=self.ai_enabled,
                       font=("Arial", 10)).pack(side=tk.LEFT)
        
        self.ai_status_label = tk.Label(ai_top, text="", font=("Arial", 9), fg="gray")
        self.ai_status_label.pack(side=tk.LEFT, padx=10)
        
        tk.Button(ai_top, text="⚙ AI 설정", command=self.open_ai_settings,
                 font=("Arial", 9), cursor="hand2").pack(side=tk.RIGHT)
        
        self.update_ai_status()
        
        # 출력 파일
        out_frame = tk.Frame(left)
        out_frame.pack(fill=tk.X, pady=5)
        tk.Label(out_frame, text="출력파일:", font=("Arial", 10)).pack(side=tk.LEFT)
        self.entry_out = tk.Entry(out_frame, font=("Arial", 10))
        self.entry_out.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d")}_미정.docx')
        self.entry_out.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 진행률
        progress_frame = tk.Frame(left)
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress = tk.Canvas(progress_frame, height=25, bg="#E0E0E0", highlightthickness=0)
        self.progress.pack(fill=tk.X)
        
        self.status_label = tk.Label(left, text="상태: 준비 완료", font=("Arial", 10, "bold"), fg="blue")
        self.status_label.pack()
        
        self.elapsed_label = tk.Label(left, text="", font=("Arial", 9), fg="gray")
        self.elapsed_label.pack()
        
        # 타이머 관련 변수
        self._timer_running = False
        self._timer_start = None
        self._current_step = ""
        
        # 로그
        self.log_area = scrolledtext.ScrolledText(left, height=5, width=50, font=("Arial", 9))
        self.log_area.pack(fill=tk.X, pady=3)
        self.log_area.config(state=tk.DISABLED)
        
        # 버튼 프레임 1
        btn_frame1 = tk.Frame(left)
        btn_frame1.pack(fill=tk.X, pady=3)
        tk.Button(btn_frame1, text="🚀 변환 시작", command=self.convert, 
                 font=("Arial", 10, "bold"), bg="#FF9800", fg="white", 
                 padx=8, pady=5, cursor="hand2").pack(side=tk.LEFT, padx=2)
        self.btn_save = tk.Button(btn_frame1, text="💾 문서 저장", command=self.save_document, 
                                  font=("Arial", 10, "bold"), bg="#4CAF50", fg="white",
                                  padx=8, pady=5, state=tk.DISABLED, cursor="hand2")
        self.btn_save.pack(side=tk.LEFT, padx=2)
        self.btn_preview = tk.Button(btn_frame1, text="👁 미리보기", command=self.preview, 
                                     font=("Arial", 10), state=tk.DISABLED, cursor="hand2")
        self.btn_preview.pack(side=tk.LEFT, padx=2)
        
        # 버튼 프레임 2
        btn_frame2 = tk.Frame(left)
        btn_frame2.pack(fill=tk.X, pady=3)
        self.btn_folder = tk.Button(btn_frame2, text="📂 폴더", command=self.open_folder, 
                                   font=("Arial", 10), state=tk.DISABLED, cursor="hand2")
        self.btn_folder.pack(side=tk.LEFT, padx=2)
        self.btn_file = tk.Button(btn_frame2, text="📄 파일", command=self.open_file, 
                                 font=("Arial", 10), state=tk.DISABLED, cursor="hand2")
        self.btn_file.pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame2, text="🔄 초기화", command=self.reset, 
                 font=("Arial", 10), cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(btn_frame2, text="❌ 종료", command=root.quit, 
                 font=("Arial", 10), cursor="hand2").pack(side=tk.LEFT, padx=2)
        
        # ===== 오른쪽 패널 (미리보기) =====
        right = tk.Frame(main, bg="white", relief=tk.SUNKEN, bd=1)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        
        tk.Label(right, text="📋 미리보기", font=("Arial", 12, "bold"), bg="white").pack(pady=10)
        self.preview_area = scrolledtext.ScrolledText(right, height=40, width=40, 
                                                     bg="#f9f9f9", wrap=tk.WORD, font=("Arial", 9))
        self.preview_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.preview_area.config(state=tk.DISABLED)
        
        # 드래그앤드롭 설정
        if DND_AVAILABLE:
            self.file_display.drop_target_register("DND_Files")
            self.file_display.dnd_bind('<<Drop>>', self.drop_handler)
            self.log("프로그램 시작\n1. 회의정보 입력\n2. 파일을 드래그앤드롭 또는 파일 선택 버튼 사용\n3. 변환 시작")
        else:
            self.log("프로그램 시작\n1. 회의정보 입력\n2. 파일 선택 버튼으로 음성파일 선택\n3. 변환 시작")
    
    def select_file(self):
        """파일 선택 대화상자 오픈"""
        file = filedialog.askopenfilename(
            title="m4a 음성 파일 선택",
            filetypes=[("M4A Files", "*.m4a"), ("All Files", "*.*")]
        )
        if file:
            self.set_file(file)
    
    def set_file(self, path):
        """파일 설정"""
        self.selected_file = path
        size = os.path.getsize(path) / (1024 * 1024)
        
        # UI 전환
        self.empty_state.pack_forget()
        self.file_state.pack(fill=tk.BOTH, expand=True)
        
        # 파일 정보 업데이트
        self.file_name_label.config(text=f"📄 {os.path.basename(path)}")
        self.file_size_label.config(text=f"크기: {size:.2f} MB")
        
        self.file_display.config(bg="#c8e6c9")
        
        self.log(f"파일 선택: {os.path.basename(path)}")
    
    def drop_handler(self, event):
        """드래그앤드롭 처리"""
        files = self.root.tk.splitlist(event.data)
        for file in files:
            # 경로에서 중괄호 제거 (Windows에서 필요)
            file = file.strip('{}').strip('"').strip("'")
            
            # m4a 파일만 처리
            if file.lower().endswith('.m4a'):
                if os.path.exists(file):
                    self.set_file(file)
                    return
            
        messagebox.showerror("오류", "m4a 파일만 지원합니다!")
    
    def log(self, msg):
        """로그 추가"""
        self.log_area.config(state=tk.NORMAL)
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.root.update()
    
    def update_progress(self, value, status_text=None):
        """진행률 업데이트 (바 + % + 상태 텍스트)"""
        self.progress.delete("all")
        w = self.progress.winfo_width()
        if w < 2: w = 650
        h = 25
        fw = int((value / 100) * w)
        
        # 배경 바
        self.progress.create_rectangle(0, 0, w, h, fill="#E0E0E0", outline="")
        
        # 진행 바 (색상 그라데이션)
        if value < 50:
            bar_color = "#FF9800"  # 주황색 (음성변환)
        elif value < 80:
            bar_color = "#9C27B0"  # 보라색 (AI 처리)
        else:
            bar_color = "#4CAF50"  # 초록색 (완료)
        
        if fw > 0:
            self.progress.create_rectangle(0, 0, fw, h, fill=bar_color, outline="")
        
        # 텍스트: % + 상태
        display = f"{value}%"
        if status_text:
            display += f"  {status_text}"
        self.progress.create_text(w/2, h/2, text=display, fill="white" if fw > w*0.3 else "black",
                                 font=("Arial", 9, "bold"))
        
        if status_text:
            self.status_label.config(text=f"상태: {status_text}")
        
        self.root.update_idletasks()
    
    def _start_timer(self, step_name):
        """경과 시간 타이머 시작"""
        import time
        self._timer_running = True
        self._timer_start = time.time()
        self._current_step = step_name
        self._update_timer()
    
    def _update_timer(self):
        """타이머 업데이트 (매초)"""
        if not self._timer_running:
            return
        import time
        elapsed = time.time() - self._timer_start
        mins = int(elapsed // 60)
        secs = int(elapsed % 60)
        self.elapsed_label.config(text=f"⏱ {self._current_step} 경과 시간: {mins}분 {secs}초", fg="#666")
        self.root.after(1000, self._update_timer)
    
    def _stop_timer(self):
        """타이머 정지"""
        self._timer_running = False
    
    def convert(self):
        """변환 시작"""
        if not self.selected_file:
            messagebox.showerror("오류", "음성 파일을 선택해주세요!")
            return
        threading.Thread(target=self.convert_thread, daemon=True).start()
    
    def open_ai_settings(self):
        """AI 설정 창 열기"""
        if AI_AVAILABLE:
            AISettingsWindow(self.root, callback=lambda cfg: self.update_ai_status())
        else:
            messagebox.showwarning("알림", "AI 모듈을 로드할 수 없습니다.\n필요 패키지: google-generativeai, openai")
    
    def update_ai_status(self):
        """AI 상태 표시 업데이트"""
        if not AI_AVAILABLE:
            self.ai_status_label.config(text="⚠ AI 모듈 미설치", fg="red")
            return
        config = load_config()
        provider = config.get("ai_provider", "gemini")
        keys_field = "gemini_api_keys" if provider == "gemini" else "openai_api_keys"
        key_field = "gemini_api_key" if provider == "gemini" else "openai_api_key"
        keys = config.get(keys_field, [])
        single = config.get(key_field, "")
        key_count = len(keys) if keys else (1 if single else 0)
        provider_name = "Gemini" if provider == "gemini" else "OpenAI"
        if key_count > 0:
            self.ai_status_label.config(text=f"✅ {provider_name} 연결됨 (키 {key_count}개)", fg="green")
        else:
            self.ai_status_label.config(text=f"⚠ {provider_name} API 키 필요", fg="orange")
    
    def convert_thread(self):
        """변환 스레드"""
        try:
            info = {
                '회의명': self.entries['회의명'].get() or '정기 회의',
                '장소': self.entries['장소'].get() or '미정',
                '일시': self.entries['일시'].get(),
                '작성자': self.entries['작성자'].get() or '미정',
                '참석자': self.entries['참석자'].get() or '미정',
                '업체이름': self.entries['업체이름'].get() or '미정',
            }
            
            out = self.entry_out.get()
            if not out.endswith('.docx'): out += '.docx'
            company = self.entries['업체이름'].get() or '미정'
            company = company.replace(' ', '_')
            auto_filename = f'회의록_{datetime.now().strftime("%Y%m%d")}_{company}.docx'
            out = auto_filename
            
            # ============================================
            # === 1단계: 음성 → 텍스트 변환 ===
            # ============================================
            self.update_progress(5, "🎤 음성 변환 준비 중...")
            self.log("━" * 40)
            self.log("1단계: 음성 → 텍스트 변환")
            
            from speech_to_text import SpeechToText
            self.update_progress(10, "🎤 Whisper 모델 로딩 중...")
            self.log("Whisper 모델 로딩 중...")
            converter = SpeechToText(model="base")
            self.update_progress(15, "🎤 모델 로딩 완료")
            self.log("✅ 모델 로딩 완료")
            
            # 타이머 시작
            self._start_timer("음성 변환")
            self.update_progress(20, "🎤 음성 변환 중... (대용량 파일은 10~30분 소요)")
            self.log("음성 변환 중... (대용량 파일은 10~30분 소요될 수 있습니다)")
            
            def on_stt_progress(percent, message):
                self.update_progress(percent, "🎤 " + message)
                self.log(message)
            
            self.transcript = converter.convert_m4a_to_text(self.selected_file, progress_callback=on_stt_progress)
            self._stop_timer()
            self.update_progress(50, "✅ 음성 변환 완료!")
            self.log("✅ 음성 변환 완료")
            
            # ============================================
            # === 2단계: AI 회의록 구성 ===
            # ============================================
            ai_content = None
            if self.ai_enabled.get() and AI_AVAILABLE:
                try:
                    config = load_config()
                    provider = config.get("ai_provider", "gemini")
                    
                    # 단일 키 + 다중 키 모두 확인
                    key_field = "gemini_api_key" if provider == "gemini" else "openai_api_key"
                    keys_field = "gemini_api_keys" if provider == "gemini" else "openai_api_keys"
                    single_key = config.get(key_field, "")
                    multi_keys = config.get(keys_field, [])
                    has_key = bool(single_key) or bool(multi_keys)
                    
                    if not has_key:
                        self.log("⚠ API 키가 설정되지 않았습니다. AI 설정을 확인해주세요.")
                        self.log("⚠ 원본 텍스트로 문서를 생성합니다.")
                    else:
                        self._start_timer("AI 회의록 구성")
                        self.update_progress(55, "🤖 AI 회의록 구성 준비 중...")
                        self.log("━" * 40)
                        self.log("2단계: AI 회의록 자동 구성")
                        
                        provider_name = "Gemini" if provider == "gemini" else "OpenAI"
                        self.update_progress(60, "🤖 " + provider_name + " AI로 회의록 구성 중...")
                        self.log("AI 제공자: " + provider_name)
                        self.log("AI가 회의 내용을 분석하고 있습니다...")
                        
                        ai_content = generate_ai_meeting_minutes(self.transcript, info, config)
                        self.ai_minutes = ai_content
                        self._stop_timer()
                        self.update_progress(85, "✅ AI 회의록 구성 완료!")
                        self.log("✅ AI 회의록 구성 완료")
                        
                except Exception as ai_err:
                    self._stop_timer()
                    self.log("⚠ AI 오류: " + str(ai_err))
                    self.log("⚠ 원본 텍스트로 문서를 생성합니다.")
                    ai_content = None
            elif self.ai_enabled.get() and not AI_AVAILABLE:
                self.log("⚠ AI 모듈 미설치 - 원본 텍스트로 진행")
            elif not self.ai_enabled.get():
                self.log("ℹ AI 비활성화 - 원본 텍스트로 진행")
            
            # ============================================
            # === 미리보기 표시 ===
            # ============================================
            self.update_progress(90, "📋 미리보기 준비 중...")
            
            self.doc_content = ai_content if ai_content else self.transcript
            self.meeting_info = info
            
            # 미리보기 업데이트
            preview = "【회의정보】\n"
            preview += "회의명: " + info['회의명'] + "\n"
            preview += "장소: " + info['장소'] + "\n"
            preview += "일시: " + info['일시'] + "\n"
            preview += "작성자: " + info['작성자'] + "\n"
            preview += "참석자: " + info['참석자'] + "\n"
            preview += "업체이름: " + info['업체이름'] + "\n"
            
            sep_double = "═" * 40
            sep_single = "─" * 40
            
            if ai_content:
                preview += "\n" + sep_double + "\n"
                preview += "【AI 생성 회의록】 ← 이 내용이 Word 문서에 반영됩니다\n"
                preview += sep_double + "\n"
                preview += ai_content
                preview += "\n\n" + sep_single + "\n"
                preview += "【음성 변환 원문 (참고용)】\n"
                preview += sep_single + "\n"
                preview += self.transcript
            else:
                preview += "\n【회의내용】\n"
                preview += self.transcript
                preview += "\n\n⚠ AI 요약이 적용되지 않았습니다. 원본 텍스트가 문서에 반영됩니다."
            
            self.preview_area.config(state=tk.NORMAL)
            self.preview_area.delete(1.0, tk.END)
            self.preview_area.insert(tk.END, preview)
            self.preview_area.config(state=tk.DISABLED)
            
            # ============================================
            # === 완료 ===
            # ============================================
            self.update_progress(100, "✅ 변환 완료! 미리보기 확인 후 저장해주세요")
            self.elapsed_label.config(text="")
            self.log("━" * 40)
            if ai_content:
                self.log("✅ AI 회의록 구성 완료! 오른쪽 미리보기를 확인하세요.")
                self.log("📝 AI가 구성한 회의록이 Word 문서에 반영됩니다.")
            else:
                self.log("✅ 변환 완료! (원본 텍스트 기준)")
            self.log("💾 '문서 저장' 버튼을 눌러 Word 파일로 저장하세요.")
            
            self.btn_save.config(state=tk.NORMAL)
            self.btn_preview.config(state=tk.NORMAL)
            
            messagebox.showinfo("변환 완료", "미리보기를 확인한 후 '💾 문서 저장' 버튼을 눌러주세요.")
            
        except Exception as e:
            self._stop_timer()
            self.update_progress(0, "❌ 오류 발생")
            self.status_label.config(text="상태: ❌ 오류", fg="red")
            self.elapsed_label.config(text="")
            self.log("❌ 오류: " + str(e))
            messagebox.showerror("오류", "오류 발생:\n" + str(e))
    
    def save_document(self):
        """회의록 Word 문서 저장"""
        if not self.doc_content or not self.meeting_info:
            messagebox.showwarning("알림", "먼저 변환을 실행해주세요.")
            return
        
        company = self.meeting_info.get('업체이름', '미정').replace(' ', '_')
        default_name = f'회의록_{datetime.now().strftime("%Y%m%d")}_{company}.docx'
        
        save_path = filedialog.asksaveasfilename(
            title="회의록 저장",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word 문서", "*.docx"), ("모든 파일", "*.*")]
        )
        if not save_path:
            return
        
        try:
            self.log("💾 Word 문서 저장 중...")
            self.status_label.config(text="상태: 문서 저장 중...", fg="orange")
            
            from document_generator import generate_meeting_minutes
            generate_meeting_minutes(self.meeting_info, self.doc_content, save_path)
            
            self.generated_file = os.path.abspath(save_path)
            self.status_label.config(text="상태: ✅ 저장 완료!", fg="green")
            self.log(f"✅ 저장 완료: {os.path.basename(save_path)}")
            
            self.btn_folder.config(state=tk.NORMAL)
            self.btn_file.config(state=tk.NORMAL)
            
            messagebox.showinfo("저장 완료", f"회의록이 저장되었습니다.\n{save_path}")
        except Exception as e:
            self.log(f"❌ 저장 오류: {str(e)}")
            messagebox.showerror("오류", f"저장 중 오류 발생:\n{str(e)}")
    
    def preview(self):
        """미리보기 팝업"""
        if not self.doc_content and not self.transcript:
            messagebox.showwarning("알림", "먼저 변환을 실행해주세요.")
            return
        
        info_text = "\n".join([
            f"회의명: {self.entries['회의명'].get()}",
            f"장소: {self.entries['장소'].get()}",
            f"일시: {self.entries['일시'].get()}",
            f"작성자: {self.entries['작성자'].get()}",
            f"참석자: {self.entries['참석자'].get()}",
            f"업체이름: {self.entries['업체이름'].get()}"
        ])
        content = f"【회의정보】\n{info_text}\n\n"
        if self.ai_minutes:
            content += f"【AI 생성 회의록】\n{self.ai_minutes}\n\n"
        content += f"【음성 변환 원문】\n{self.transcript}"
        PreviewWindow(self.root, "회의록 미리보기", content)
    
    def open_folder(self):
        """폴더 열기"""
        if self.generated_file:
            os.startfile(os.path.dirname(self.generated_file))
            self.log("폴더 열기")
    
    def open_file(self):
        """파일 열기"""
        if self.generated_file and os.path.exists(self.generated_file):
            os.startfile(self.generated_file)
            self.log("파일 열기")
    
    def reset(self):
        """초기화"""
        self._stop_timer()
        for key, default in [("회의명","정기 회의"),("장소","미정"),("일시",datetime.now().strftime("%Y.%m.%d")),("작성자","미정"),("참석자","미정"),("업체이름","미정")]:
            self.entries[key].delete(0, tk.END)
            self.entries[key].insert(0, default)
        
        self.file_state.pack_forget()
        self.empty_state.pack(fill=tk.BOTH, expand=True)
        self.file_display.config(bg="lightgray")
        
        self.entry_out.delete(0, tk.END)
        self.entry_out.insert(0, f'회의록_{datetime.now().strftime("%Y%m%d")}_미정.docx')
        
        self.update_progress(0, "")
        self.status_label.config(text="상태: 준비 완료", fg="blue")
        self.elapsed_label.config(text="")
        
        self.log_area.config(state=tk.NORMAL)
        self.log_area.delete(1.0, tk.END)
        self.log_area.config(state=tk.DISABLED)
        
        self.preview_area.config(state=tk.NORMAL)
        self.preview_area.delete(1.0, tk.END)
        self.preview_area.config(state=tk.DISABLED)
        
        self.btn_save.config(state=tk.DISABLED)
        self.btn_folder.config(state=tk.DISABLED)
        self.btn_file.config(state=tk.DISABLED)
        self.btn_preview.config(state=tk.DISABLED)
        
        self.selected_file = None
        self.transcript = None
        self.ai_minutes = None
        self.doc_content = None
        self.meeting_info = None
        self.generated_file = None
        
        self.log("프로그램 초기화됨")

if __name__ == '__main__':
    if DND_AVAILABLE:
        root = tkdnd.Tk()
    else:
        root = tk.Tk()
    app = MeetingApp(root)
    root.mainloop()
