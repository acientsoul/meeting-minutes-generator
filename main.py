# -*- coding: utf-8 -*-
"""
회의록 자동 생성 메인 프로그램
음성 파일(m4a) → 텍스트 변환 → Word 문서 생성
"""

import os
import sys
from datetime import datetime
from speech_to_text import convert_speech_to_text
from document_generator import generate_meeting_minutes


def get_user_input():
    """사용자로부터 회의 정보 입력받기"""
    print("\n" + "="*60)
    print("📝 회의록 자동 생성 프로그램")
    print("="*60)
    
    meeting_info = {}
    
    print("\n[회의 정보 입력]")
    print("(각 항목을 입력하고 Enter를 누르세요)\n")
    
    meeting_info['회의명'] = input("✓ 회의명: ").strip() or "정기 회의"
    meeting_info['장소'] = input("✓ 장소: ").strip() or "미정"
    meeting_info['일시'] = input("✓ 일시 (예: 2026.02.12): ").strip() or datetime.now().strftime("%Y.%m.%d")
    meeting_info['작성자'] = input("✓ 작성자: ").strip() or "미정"
    meeting_info['참석자'] = input("✓ 참석자 (쉼표로 구분): ").strip() or "미정"
    meeting_info['업체이름'] = input("✓ 업체이름: ").strip() or "미정"
    
    return meeting_info


def get_audio_file():
    """음성 파일 경로 입력받기"""
    while True:
        audio_path = input("\n🎤 m4a 음성 파일 경로를 입력하세요: ").strip()
        
        if not audio_path:
            print("❌ 파일 경로를 입력해주세요")
            continue
        
        # 경로 정규화
        audio_path = os.path.normpath(audio_path)
        
        if not os.path.exists(audio_path):
            print(f"❌ 파일을 찾을 수 없습니다: {audio_path}")
            continue
        
        if not audio_path.lower().endswith('.m4a'):
            print("❌ m4a 파일만 지원합니다")
            continue
        
        file_size_mb = os.path.getsize(audio_path) / (1024 * 1024)
        print(f"✅ 파일 선택됨: {os.path.basename(audio_path)} ({file_size_mb:.1f}MB)")
        
        return audio_path


def get_output_path(meeting_info=None):
    """출력 파일 경로 설정"""
    company = ""
    if meeting_info and meeting_info.get('업체이름'):
        company = meeting_info['업체이름'].replace(' ', '_')
    else:
        company = "미정"
    
    timestamp = datetime.now().strftime("%Y%m%d")
    default_name = f"회의록_{timestamp}_{company}.docx"
    
    custom_name = input(f"\n📄 출력 파일명 (기본값: {default_name}): ").strip()
    
    if not custom_name:
        output_path = default_name
    else:
        if not custom_name.endswith('.docx'):
            custom_name += '.docx'
        output_path = custom_name
    
    return output_path


def main():
    """메인 실행 함수"""
    try:
        # 사용자 입력
        meeting_info = get_user_input()
        audio_file = get_audio_file()
        output_file = get_output_path(meeting_info)
        
        print("\n" + "="*60)
        print("🔄 처리 시작")
        print("="*60)
        
        # 단계 1: 음성을 텍스트로 변환
        print("\n[단계 1/2] 음성을 텍스트로 변환 중...")
        transcript = convert_speech_to_text(audio_file, model_size="base")
        
        # 단계 2: Word 문서 생성
        print("\n[단계 2/2] Word 문서 생성 중...")
        output_path = generate_meeting_minutes(meeting_info, transcript, output_file)
        
        print("\n" + "="*60)
        print("✅ 완료!")
        print("="*60)
        print(f"\n📄 생성된 파일: {os.path.abspath(output_path)}")
        print(f"\n📋 회의록 정보:")
        print(f"  - 회의명: {meeting_info['회의명']}")
        print(f"  - 장소: {meeting_info['장소']}")
        print(f"  - 일시: {meeting_info['일시']}")
        print(f"  - 작성자: {meeting_info['작성자']}")
        print(f"  - 참석자: {meeting_info['참석자']}")
        print(f"  - 업체이름: {meeting_info['업체이름']}")
        
    except KeyboardInterrupt:
        print("\n\n⚠️  프로그램이 사용자에 의해 중단되었습니다")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
