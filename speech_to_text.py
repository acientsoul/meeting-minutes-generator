# -*- coding: utf-8 -*-
"""
음성 파일(m4a)을 텍스트로 변환하는 모듈
OpenAI Whisper 사용
"""

import whisper
import os
import time


class SpeechToText:
    def __init__(self, model="base"):
        """
        Whisper 모델 초기화
        model 옵션: "tiny", "base", "small", "medium", "large"
        """
        print(f"Whisper 모델 로딩 중 ({model})...")
        self.model = whisper.load_model(model)
        print("✅ 모델 로딩 완료")

    def convert_m4a_to_text(self, file_path, progress_callback=None):
        """
        음성 파일을 텍스트로 변환
        
        Args:
            file_path: 음성 파일 경로
            progress_callback: 진행률 콜백 함수 (percent, message)
            
Returns:
            변환된 텍스트
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        
        print(f"\n🎤 음성 파일 변환 중: {os.path.basename(file_path)}")
        print("이 과정은 파일 길이에 따라 시간이 걸릴 수 있습니다...")
        
        try:
            if progress_callback:
                progress_callback(30, "음성 분석 중... (대용량 파일은 시간이 소요됩니다)")
            
            start_time = time.time()
            
            # Whisper를 사용하여 음성을 텍스트로 변환 (verbose=True로 세그먼트별 출력)
            result = self.model.transcribe(file_path, language="ko", verbose=False)
            
            elapsed = time.time() - start_time
            elapsed_min = int(elapsed // 60)
            elapsed_sec = int(elapsed % 60)
            
            # 세그먼트 정보로 진행률 표시
            segments = result.get("segments", [])
            total_segments = len(segments)
            
            transcript = result["text"]
            print(f"✅ 음성 변환 완료 ({elapsed_min}분 {elapsed_sec}초 소요, {total_segments}개 세그먼트)")
            
            if progress_callback:
                progress_callback(50, f"✅ 음성 변환 완료 ({elapsed_min}분 {elapsed_sec}초)")
            
            return transcript
            
        except Exception as e:
            print(f"❌ 오류 발생: {str(e)}")
            raise


def convert_speech_to_text(audio_file_path, model_size="base"):
    """
    편의 함수: m4a 파일을 텍스트로 변환
    
    Args:
        audio_file_path: m4a 파일 경로
        model_size: Whisper 모델 크기
        
    Returns:
        변환된 텍스트
    """
    converter = SpeechToText(model=model_size)
    return converter.convert_m4a_to_text(audio_file_path)
