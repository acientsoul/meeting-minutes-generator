# -*- coding: utf-8 -*-
"""
회의록 자동 생성 GUI 프로그램
PySimpleGUI를 사용한 사용자 친화적 인터페이스
"""

import PySimpleGUI as sg
import os
import sys
from datetime import datetime
from pathlib import Path
from speech_to_text import SpeechToText
from document_generator import generate_meeting_minutes


# PySimpleGUI 테마 설정
try:
    sg.theme('DarkBlue3')
except:
    pass
sg.set_options(font=('Arial', 10), element_padding=(8, 8))


class MeetingMinutesApp:
    def __init__(self):
        self.converter = None
        self.selected_file = None
    
    def create_window(self):
        """GUI 윈도우 생성"""
        layout = [
            # 제목
            [sg.Text('🎤 회의록 자동 생성 프로그램', 
                     font=('Arial', 16, 'bold'),
                     text_color='white',
                     background_color='#0066cc',
                     expand_x=True,
                     pad=(0, 10))],
            
            # 구분선
            [sg.Text('_' * 60, text_color='gray')],
            
            # 회의 정보 섹션
            [sg.Frame('📋 회의 정보', [
                [sg.Text('회의명:', size=(12, 1)), 
                 sg.InputText('정기 회의', key='회의명', expand_x=True)],
                
                [sg.Text('장소:', size=(12, 1)), 
                 sg.InputText('미정', key='장소', expand_x=True)],
                
                [sg.Text('일시:', size=(12, 1)), 
                 sg.InputText(datetime.now().strftime("%Y.%m.%d"), 
                             key='일시', expand_x=True)],
                
                [sg.Text('작성자:', size=(12, 1)), 
                 sg.InputText('미정', key='작성자', expand_x=True)],
                
                [sg.Text('참석자:', size=(12, 1)), 
                 sg.InputText('미정', key='참석자', expand_x=True)],
            ], expand_x=True)],
            
            # 파일 선택 섹션
            [sg.Frame('🎵 음성 파일', [
                [sg.Text('m4a 파일:', size=(12, 1)), 
                 sg.InputText(key='파일경로', expand_x=True, disabled=True),
                 sg.FileBrowse('선택', file_types=(('M4A Files', '*.m4a'),),
                              key='파일선택', enable_events=True, 
                              target='파일경로')],
                
                [sg.Text('', key='파일상태', text_color='lightblue', size=(50, 1))],
            ], expand_x=True)],
            
            # 출력 파일명 섹션
            [sg.Frame('💾 출력 설정', [
                [sg.Text('파일명:', size=(12, 1)), 
                 sg.InputText(f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx',
                             key='출력파일명', expand_x=True)],
            ], expand_x=True)],
            
            # 진행 상황
            [sg.Frame('⏳ 진행 상황', [
                [sg.ProgressBar(100, orientation='h', size=(45, 20), 
                               key='진행바', bar_color=('lightblue', 'gray'))],
                [sg.Text('준비 완료', key='상태메시지', text_color='lightgreen')],
            ], expand_x=True)],
            
            # 결과
            [sg.Frame('✅ 결과', [
                [sg.Multiline(size=(50, 5), 
                             key='결과', 
                             disabled=True,
                             background_color='#1a1a1a',
                             text_color='lightgreen',
                             expand_x=True,
                             expand_y=True)],
            ], expand_x=True, expand_y=True)],
            
            # 버튼들
            [sg.Button('🚀 변환 시작', key='변환시작', size=(12, 1), button_color=('white', 'green')),
             sg.Button('📂 폴더 열기', key='폴더열기', size=(12, 1), disabled=True),
             sg.Button('📄 파일 열기', key='파일열기', size=(12, 1), disabled=True),
             sg.Button('🔄 초기화', key='초기화', size=(12, 1)),
             sg.Button('❌ 종료', key='종료', size=(12, 1))],
        ]
        
        return sg.Window('회의록 자동 생성 - GUI 버전', 
                        layout, 
                        size=(700, 900),
                        finalize=True)
    
    def log_message(self, window, message, message_type='info'):
        """결과 창에 메시지 출력"""
        current_text = window['결과'].get()
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if message_type == 'error':
            prefix = "❌ "
        elif message_type == 'success':
            prefix = "✅ "
        elif message_type == 'info':
            prefix = "ℹ️  "
        else:
            prefix = "→ "
        
        new_text = f"{current_text}[{timestamp}] {prefix}{message}\n"
        window['결과'].update(new_text)
        window['결과'].see(len(new_text))  # 자동 스크롤
    
    def run(self):
        """애플리케이션 실행"""
        window = self.create_window()
        generated_file = None
        
        self.log_message(window, "프로그램이 시작되었습니다")
        self.log_message(window, "회의 정보를 입력하고 음성 파일을 선택해주세요", 'info')
        
        while True:
            event, values = window.read()
            
            if event == sg.WINDOW_CLOSED or event == '종료':
                break
            
            elif event == '파일선택':
                # 파일 선택 시 상태 표시
                file_path = values['파일경로']
                if file_path:
                    file_size = os.path.getsize(file_path) / (1024 * 1024)
                    window['파일상태'].update(
                        f"✓ {os.path.basename(file_path)} ({file_size:.1f}MB)")
                    self.selected_file = file_path
            
            elif event == '초기화':
                # 초기화
                window['회의명'].update('정기 회의')
                window['장소'].update('미정')
                window['일시'].update(datetime.now().strftime("%Y.%m.%d"))
                window['작성자'].update('미정')
                window['참석자'].update('미정')
                window['파일경로'].update('')
                window['파일상태'].update('')
                window['출력파일명'].update(
                    f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
                window['진행바'].update(0)
                window['상태메시지'].update('준비 완료')
                window['결과'].update('')
                window['폴더열기'].update(disabled=True)
                window['파일열기'].update(disabled=True)
                self.selected_file = None
                generated_file = None
                self.log_message(window, "초기화되었습니다", 'info')
            
            elif event == '변환시작':
                # 유효성 검사
                if not self.selected_file:
                    self.log_message(window, "음성 파일을 선택해주세요", 'error')
                    continue
                
                # 회의 정보 수집
                meeting_info = {
                    '회의명': values['회의명'] or '정기 회의',
                    '장소': values['장소'] or '미정',
                    '일시': values['일시'] or datetime.now().strftime("%Y.%m.%d"),
                    '작성자': values['작성자'] or '미정',
                    '참석자': values['참석자'] or '미정',
                }
                
                output_file = values['출력파일명'] or 'meeting_minutes.docx'
                if not output_file.endswith('.docx'):
                    output_file += '.docx'
                
                try:
                    window['변환시작'].update(disabled=True)
                    window['상태메시지'].update('진행 중...')
                    window.refresh()
                    
                    # 단계 1: 음성 변환
                    self.log_message(window, "음성 변환 중... (1단계)", 'info')
                    window['진행바'].update(30)
                    window.refresh()
                    
                    self.log_message(window, "Whisper 모델 로딩 중...")
                    converter = SpeechToText(model="base")
                    
                    self.log_message(window, "음성 파일 변환 중... (시간이 소요될 수 있습니다)")
                    transcript = converter.convert_m4a_to_text(self.selected_file)
                    
                    window['진행바'].update(60)
                    window.refresh()
                    
                    # 단계 2: Word 문서 생성
                    self.log_message(window, "Word 문서 생성 중... (2단계)", 'info')
                    generate_meeting_minutes(meeting_info, transcript, output_file)
                    
                    window['진행바'].update(100)
                    window['상태메시지'].update('✅ 완료!')
                    
                    # 결과 표시
                    generated_file = os.path.abspath(output_file)
                    self.log_message(window, f"생성 완료!", 'success')
                    self.log_message(window, f"파일: {generated_file}", 'info')
                    self.log_message(window, f"회의명: {meeting_info['회의명']}", 'info')
                    self.log_message(window, f"장소: {meeting_info['장소']}", 'info')
                    self.log_message(window, f"참석자: {meeting_info['참석자']}", 'info')
                    
                    # 버튼 활성화
                    window['폴더열기'].update(disabled=False)
                    window['파일열기'].update(disabled=False)
                    
                except Exception as e:
                    window['상태메시지'].update('❌ 오류 발생')
                    self.log_message(window, f"오류: {str(e)}", 'error')
                    import traceback
                    self.log_message(window, traceback.format_exc(), 'error')
                
                finally:
                    window['변환시작'].update(disabled=False)
            
            elif event == '폴더열기':
                if generated_file:
                    folder = os.path.dirname(generated_file)
                    if sys.platform == 'win32':
                        os.startfile(folder)
                    elif sys.platform == 'darwin':
                        os.system(f'open "{folder}"')
                    else:
                        os.system(f'xdg-open "{folder}"')
                    self.log_message(window, "폴더를 열었습니다", 'success')
            
            elif event == '파일열기':
                if generated_file and os.path.exists(generated_file):
                    if sys.platform == 'win32':
                        os.startfile(generated_file)
                    elif sys.platform == 'darwin':
                        os.system(f'open "{generated_file}"')
                    else:
                        os.system(f'xdg-open "{generated_file}"')
                    self.log_message(window, "파일을 열었습니다", 'success')
        
        window.close()


if __name__ == "__main__":
    app = MeetingMinutesApp()
    app.run()
