# -*- coding: utf-8 -*-
"""
회의록 자동 생성 GUI (간단 버전)
"""

import PySimpleGUI as sg
import os
import sys
from datetime import datetime

class MeetingApp:
    def __init__(self):
        self.generated_file = None
    
    def run(self):
        layout = [
            [sg.Text('🎤 회의록 자동 생성', font=('Arial', 14, 'bold'))],
            [sg.Text('-' * 50)],
            [sg.Text('회의명:', size=(8, 1)), sg.InputText('정기 회의', key='회의명', size=(35, 1))],
            [sg.Text('장소:', size=(8, 1)), sg.InputText('미정', key='장소', size=(35, 1))],
            [sg.Text('일시:', size=(8, 1)), sg.InputText(datetime.now().strftime("%Y.%m.%d"), key='일시', size=(35, 1))],
            [sg.Text('작성자:', size=(8, 1)), sg.InputText('미정', key='작성자', size=(35, 1))],
            [sg.Text('참석자:', size=(8, 1)), sg.InputText('미정', key='참석자', size=(35, 1))],
            [sg.Text('-' * 50)],
            [sg.Text('음성 파일:', size=(8, 1)), sg.InputText(key='파일', size=(35, 1), disabled=True),
             sg.FileBrowse('선택', file_types=(('M4A', '*.m4a'),), target='파일')],
            [sg.Text('파일상태:', key='파일상태', size=(50, 1), text_color='blue')],
            [sg.Text('-' * 50)],
            [sg.Text('출력파일:', size=(8, 1)), sg.InputText(f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx', key='출력', size=(35, 1))],
            [sg.Text('-' * 50)],
            [sg.ProgressBar(100, orientation='h', size=(40, 15), key='진행바')],
            [sg.Text('상태: 준비 완료', key='상태', size=(50, 1))],
            [sg.Multiline(size=(60, 6), key='로그', disabled=True)],
            [sg.Button('🚀 변환 시작'), sg.Button('📂 폴더 열기', disabled=True, key='폴더'),
             sg.Button('📄 파일 열기', disabled=True, key='파일_열기'),
             sg.Button('초기화'), sg.Button('종료')],
        ]
        
        window = sg.Window('회의록 자동 생성', layout)
        
        self.log(window, "프로그램 시작됨")
        self.log(window, "1. 회의 정보 입력")
        self.log(window, "2. 음성 파일 선택")
        self.log(window, "3. '변환 시작' 버튼 클릭")
        
        while True:
            event, values = window.read()
            
            if event in (sg.WINDOW_CLOSED, '종료'):
                break
            
            if event == '파일':
                if values['파일']:
                    size = os.path.getsize(values['파일']) / (1024 * 1024)
                    window['파일상태'].update(f"✓ {os.path.basename(values['파일'])} ({size:.1f}MB)")
            
            elif event == '초기화':
                window['회의명'].update('정기 회의')
                window['장소'].update('미정')
                window['일시'].update(datetime.now().strftime("%Y.%m.%d"))
                window['작성자'].update('미정')
                window['참석자'].update('미정')
                window['파일'].update('')
                window['파일상태'].update('')
                window['출력'].update(f'회의록_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx')
                window['진행바'].update(0)
                window['상태'].update('상태: 준비 완료')
                window['로그'].update('')
                window['폴더'].update(disabled=True)
                window['파일_열기'].update(disabled=True)
                self.generated_file = None
                self.log(window, "초기화됨")
            
            elif event == '🚀 변환 시작':
                if not values['파일']:
                    self.log(window, "❌ 음성 파일을 선택해주세요")
                    continue
                
                window['🚀 변환 시작'].update(disabled=True)
                window.refresh()
                
                try:
                    meeting_info = {
                        '회의명': values['회의명'] or '정기 회의',
                        '장소': values['장소'] or '미정',
                        '일시': values['일시'],
                        '작성자': values['작성자'] or '미정',
                        '참석자': values['참석자'] or '미정',
                    }
                    
                    output = values['출력'] if values['출력'].endswith('.docx') else values['출력'] + '.docx'
                    
                    self.log(window, "1단계: 음성 변환 중...")
                    window['진행바'].update(30)
                    window['상태'].update('상태: 음성 변환 중...')
                    window.refresh()
                    
                    from speech_to_text import SpeechToText
                    self.log(window, "Whisper 모델 로딩...")
                    converter = SpeechToText(model="base")
                    
                    self.log(window, "음성 파일 변환 중... (시간이 소요될 수 있습니다)")
                    transcript = converter.convert_m4a_to_text(values['파일'])
                    
                    window['진행바'].update(70)
                    window.refresh()
                    
                    self.log(window, "2단계: Word 문서 생성 중...")
                    from document_generator import generate_meeting_minutes
                    generate_meeting_minutes(meeting_info, transcript, output)
                    
                    window['진행바'].update(100)
                    window['상태'].update('상태: ✅ 완료!')
                    
                    self.generated_file = os.path.abspath(output)
                    self.log(window, "✅ 성공!")
                    self.log(window, f"파일: {self.generated_file}")
                    
                    window['폴더'].update(disabled=False)
                    window['파일_열기'].update(disabled=False)
                    
                except Exception as e:
                    window['상태'].update(f'상태: ❌ 오류 - {str(e)[:30]}')
                    self.log(window, f"❌ 오류: {str(e)}")
                
                finally:
                    window['🚀 변환 시작'].update(disabled=False)
            
            elif event == '📂 폴더':
                if self.generated_file:
                    folder = os.path.dirname(self.generated_file)
                    if sys.platform == 'win32':
                        os.startfile(folder)
                    self.log(window, "폴더 열기")
            
            elif event == '📄 파일_열기':
                if self.generated_file and os.path.exists(self.generated_file):
                    if sys.platform == 'win32':
                        os.startfile(self.generated_file)
                    self.log(window, "파일 열기")
        
        window.close()
    
    def log(self, window, msg):
        current = window['로그'].get()
        new_msg = f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n"
        window['로그'].update(current + new_msg)

if __name__ == '__main__':
    try:
        app = MeetingApp()
        app.run()
    except Exception as e:
        print(f"❌ 오류: {e}")
        import traceback
        traceback.print_exc()
        input("계속하려면 Enter를 누르세요...")
