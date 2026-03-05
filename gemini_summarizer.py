# -*- coding: utf-8 -*-
"""
Gemini AI를 이용한 회의록 축약 및 정리 모듈
"""

import os
from dotenv import load_dotenv
import google.generativeai as genai

# .env 파일 로드
load_dotenv()

class GeminiSummarizer:
    """Gemini AI를 이용한 회의록 축약 클래스"""
    
    def __init__(self):
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            raise ValueError("GEMINI_API_KEY가 설정되지 않았습니다. .env 파일을 확인하세요.")
        
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-pro')
    
    def summarize(self, transcript_text, meeting_info):
        """
        음성 변환 텍스트를 축약 및 정리
        
        Args:
            transcript_text (str): 음성 변환 원본 텍스트
            meeting_info (dict): 회의 정보
        
        Returns:
            str: 축약된 회의록
        """
        prompt = f"""다음은 회의 음성을 텍스트로 변환한 원본입니다.
이를 읽기 좋은 회의록 형태로 축약하고 정리해주세요.

【회의 정보】
- 회의명: {meeting_info.get('회의명', '정기 회의')}
- 장소: {meeting_info.get('장소', '미정')}
- 일시: {meeting_info.get('일시', '미정')}
- 작성자: {meeting_info.get('작성자', '미정')}
- 참석자: {meeting_info.get('참석자', '미정')}
- 업체이름: {meeting_info.get('업체이름', '미정')}

【원본 음성 텍스트】
{transcript_text}

【요청사항】
1. 주요 의제와 논의 사항을 명확하게 정리하세요.
2. 결정사항과 액션아이템을 구분해서 표시하세요.
3. 불필요한 반복 표현은 제거하세요.
4. 각 항목에 대한 담당자와 예상 일정을 포함하세요.
5. 최종 결과물은 체계적인 회의록 형식으로 작성해주세요.

【출력 형식】
## 주요 의제
- 의제1
- 의제2
...

## 주요 논의 사항
내용

## 결정사항
- 결정 1
- 결정 2
...

## 액션 아이템
| 담당자 | 내용 | 예상 완료일 |
|--------|------|----------|
| | | |

## 기타 사항
내용"""
        
        response = self.model.generate_content(prompt)
        return response.text
    
    def get_meeting_summary(self, transcript_text, meeting_info):
        """최종 회의록 생성"""
        try:
            summarized = self.summarize(transcript_text, meeting_info)
            return summarized
        except Exception as e:
            print(f"❌ Gemini API 오류: {str(e)}")
            # 실패 시 원본 텍스트 반환
            return transcript_text
