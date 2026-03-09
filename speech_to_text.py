# -*- coding: utf-8 -*-
"""
음성 파일(m4a)을 텍스트로 변환하는 모듈
OpenAI Whisper 사용
- 대용량 파일은 ffmpeg로 임시 wav 파일로 분할 → 디스크 기반 처리
"""

import whisper
import numpy as np
import os
import time
import gc
import subprocess
import json
import tempfile

# 청크 분할 기준 (파일 크기 MB)
CHUNK_THRESHOLD_MB = 25
# 청크 길이 (초) - 2분 단위로 분할 (메모리 절약)
CHUNK_DURATION_SEC = 120
# 청크 간 오버랩 (초) - 문장 끊김 방지
CHUNK_OVERLAP_SEC = 2


class SpeechToText:
    def __init__(self, model="base"):
        """
        Whisper 모델 초기화
        model 옵션: "tiny", "base", "small", "medium", "large"
        """
        print(f"Whisper 모델 로딩 중 ({model})...")
        self.model = whisper.load_model(model)
        print("✅ 모델 로딩 완료")

    def _get_audio_duration(self, file_path):
        """ffprobe로 오디오 파일의 총 길이(초)를 반환"""
        # 방법 1: ffprobe -show_entries
        try:
            cmd = [
                "ffprobe", "-v", "error",
                "-show_entries", "format=duration",
                "-of", "default=noprint_wrappers=1:nokey=1",
                file_path
            ]
            process = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            stdout, _ = process.communicate(timeout=30)
            if stdout:
                dur_str = stdout.decode("utf-8", errors="replace").strip()
                if dur_str:
                    return float(dur_str)
        except Exception:
            pass
        
        # 방법 2: ffprobe JSON
        try:
            cmd = [
                "ffprobe", "-v", "quiet",
                "-print_format", "json",
                "-show_format",
                file_path
            ]
            process = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            stdout, _ = process.communicate(timeout=30)
            if stdout:
                output = stdout.decode("utf-8", errors="replace")
                info = json.loads(output)
                if "format" in info and "duration" in info["format"]:
                    return float(info["format"]["duration"])
        except Exception:
            pass
        
        # 방법 3: 파일 크기 기반 추정
        file_size = os.path.getsize(file_path)
        estimated_duration = file_size / (128 * 1024 / 8)
        print(f"⚠ ffprobe 실패, 파일 크기 기반 추정: {int(estimated_duration)}초")
        return estimated_duration

    def _extract_chunk_to_wav(self, file_path, start_sec, duration_sec, output_wav):
        """
        ffmpeg로 오디오 파일의 특정 구간을 wav 파일로 추출.
        디스크에 저장하므로 메모리를 거의 사용하지 않음.
        """
        cmd = [
            "ffmpeg", "-y",
            "-ss", str(start_sec),
            "-t", str(duration_sec),
            "-i", file_path,
            "-ac", "1",
            "-ar", "16000",
            "-v", "quiet",
            output_wav
        ]
        
        process = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
        )
        _, stderr = process.communicate(timeout=120)
        
        if process.returncode != 0:
            err_msg = stderr.decode('utf-8', errors='replace') if stderr else "알 수 없는 오류"
            raise RuntimeError(f"ffmpeg 오류: {err_msg}")

    def convert_m4a_to_text(self, file_path, progress_callback=None):
        """
        음성 파일을 텍스트로 변환
        대용량 파일(25MB 이상)은 청크 분할 처리
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")
        
        file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
        print(f"\n🎤 음성 파일 변환 중: {os.path.basename(file_path)} ({file_size_mb:.1f}MB)")
        
        try:
            start_time = time.time()
            
            if file_size_mb >= CHUNK_THRESHOLD_MB:
                transcript = self._transcribe_chunked(file_path, progress_callback)
            else:
                if progress_callback:
                    progress_callback(30, "음성 분석 중...")
                result = self.model.transcribe(file_path, language="ko", verbose=False)
                transcript = result["text"]
            
            elapsed = time.time() - start_time
            elapsed_min = int(elapsed // 60)
            elapsed_sec = int(elapsed % 60)
            
            print(f"✅ 음성 변환 완료 ({elapsed_min}분 {elapsed_sec}초 소요)")
            
            if progress_callback:
                progress_callback(50, f"✅ 음성 변환 완료 ({elapsed_min}분 {elapsed_sec}초)")
            
            return transcript
            
        except Exception as e:
            print(f"❌ 오류 발생: {str(e)}")
            raise

    def _transcribe_chunked(self, file_path, progress_callback=None):
        """
        대용량 파일을 청크별로 임시 wav 파일로 추출 후 변환.
        메모리에 전체 오디오를 올리지 않아 메모리 부족 방지.
        """
        if progress_callback:
            progress_callback(22, "🔄 오디오 파일 정보 확인 중...")
        
        # 1) 총 길이 확인
        total_duration = self._get_audio_duration(file_path)
        total_min = int(total_duration // 60)
        total_sec = int(total_duration % 60)
        print(f"📊 오디오 길이: {total_min}분 {total_sec}초")
        
        # 2) 청크 시간 구간 계산
        step = CHUNK_DURATION_SEC - CHUNK_OVERLAP_SEC
        chunks = []
        pos = 0.0
        while pos < total_duration:
            end = min(pos + CHUNK_DURATION_SEC, total_duration)
            chunks.append((pos, end - pos))
            if end >= total_duration:
                break
            pos += step
        
        num_chunks = len(chunks)
        print(f"📦 {num_chunks}개 청크로 분할 처리합니다 (각 {CHUNK_DURATION_SEC}초)")
        
        if progress_callback:
            progress_callback(25, f"📦 {num_chunks}개 청크로 분할 처리 시작...")
        
        # 3) 임시 디렉토리에서 청크별 변환
        all_texts = []
        tmp_dir = tempfile.mkdtemp(prefix="whisper_chunks_")
        
        try:
            for i, (start_sec, dur_sec) in enumerate(chunks):
                start_min = int(start_sec // 60)
                start_s = int(start_sec % 60)
                end_sec = start_sec + dur_sec
                end_min = int(end_sec // 60)
                end_s = int(end_sec % 60)
                
                msg = f"🎤 청크 {i+1}/{num_chunks} 변환 중 ({start_min}:{start_s:02d}~{end_min}:{end_s:02d})"
                print(msg)
                
                if progress_callback:
                    pct = 25 + int((i / num_chunks) * 23)
                    progress_callback(pct, msg)
                
                # ffmpeg로 해당 구간을 임시 wav 파일로 추출
                tmp_wav = os.path.join(tmp_dir, f"chunk_{i:04d}.wav")
                self._extract_chunk_to_wav(file_path, start_sec, dur_sec, tmp_wav)
                
                # Whisper로 변환 (파일 경로 전달 — Whisper가 자체 로딩)
                result = self.model.transcribe(tmp_wav, language="ko", verbose=False)
                text = result["text"].strip()
                
                if text:
                    all_texts.append(text)
                
                # 임시 파일 즉시 삭제 + 메모리 해제
                try:
                    os.remove(tmp_wav)
                except OSError:
                    pass
                del result
                gc.collect()
        finally:
            # 임시 디렉토리 정리
            try:
                import shutil
                shutil.rmtree(tmp_dir, ignore_errors=True)
            except Exception:
                pass
        
        transcript = " ".join(all_texts)
        return transcript


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
