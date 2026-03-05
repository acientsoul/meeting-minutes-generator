# -*- coding: utf-8 -*-
"""
AI 회의록 자동 생성 모듈
- Google Gemini API 또는 OpenAI GPT API를 사용하여
  음성 변환 텍스트를 구조화된 회의록으로 변환
"""

import json
import os

# ===== API 지원 =====
GEMINI_AVAILABLE = False
OPENAI_AVAILABLE = False

try:
    from google import genai
    GEMINI_AVAILABLE = True
except ImportError:
    pass

try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    pass


# ===== 설정 파일 경로 =====
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ai_config.json")


def load_config():
    """설정 파일 로드"""
    default = {
        "ai_provider": "gemini",  # "gemini" or "openai"
        "gemini_api_key": "",
        "gemini_api_keys": [],      # 다중 키 지원 (로테이션)
        "openai_api_key": "",
        "openai_api_keys": [],      # 다중 키 지원 (로테이션)
        "gemini_model": "gemini-2.0-flash",
        "openai_model": "gpt-4o-mini",
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                default.update(saved)
        except Exception:
            pass
    return default


def _get_all_keys(config, provider):
    """
    해당 provider의 모든 API 키 목록을 반환.
    단일 키(gemini_api_key) + 다중 키 목록(gemini_api_keys)을 합쳐서 중복 제거.
    """
    if provider == "gemini":
        single = config.get("gemini_api_key", "")
        multi = config.get("gemini_api_keys", [])
    else:
        single = config.get("openai_api_key", "")
        multi = config.get("openai_api_keys", [])
    
    keys = []
    if single:
        keys.append(single)
    for k in (multi or []):
        k = k.strip()
        if k and k not in keys:
            keys.append(k)
    return keys


def save_config(config):
    """설정 파일 저장"""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


MEETING_PROMPT = """당신은 전문 회의록 작성 비서입니다.
아래의 음성 변환 텍스트를 분석하여, 핵심 내용만 추출·축약하여 체계적인 회의록을 작성해주세요.

## 회의 기본 정보
- 회의명: {meeting_name}
- 장소: {location}
- 일시: {date}
- 작성자: {author}
- 참석자: {attendees}
- 업체이름: {company}

## 음성 변환 원문
{transcript}

## 핵심 작성 지침

### ★ 축약 원칙 (가장 중요):
- 원문의 20~30% 분량으로 축약하세요. 장황한 설명은 1~2줄 핵심 문장으로 압축하세요.
- 중복되거나 반복되는 내용은 한 번만 작성하세요.
- 잡담, 인사말, 상투적 표현, 불필요한 수식어는 모두 제거하세요.
- "~했습니다", "~라고 합니다" 등 구어체는 "~함", "~예정" 등 간결한 명사형/개조식으로 변환하세요.
- 구체적인 수치, 날짜, 고유명사는 반드시 보존하세요.

### 번호 체계 규칙 (반드시 준수):
- 대주제: "1. 제목" (숫자+점)
- 소주제: "1) 제목" (숫자+괄호)  
- 세부항목: "i. 내용" (로마자 소문자+점)
- 하위항목: "a) 내용" (알파벳+괄호)
- 참고/화살표: "→ 내용" (화살표)

### 포함할 섹션:
1. 기업/업체 소개 (해당하는 경우, 2~3줄로 간략히)
   1) 기업 개요
   2) 주요 기술/역량
2. 주요 논의 사항 (핵심 안건별 요약)
   1) 안건명
      i. 핵심 내용 (1~2줄)
      → 결론/합의 사항
3. 기타 논의 사항 (간략히)
4. 향후 To Do List
   1) 후속 조치 항목
      → 담당자/기한 (파악 가능한 경우)

### 주의사항:
- 한국어로 작성하고, 전문적이고 명확한 개조식 문체를 사용하세요.
- 마크다운 기호(#, -, *, ** 등)는 절대 사용하지 마세요. 위의 번호 체계만 사용하세요.
- 대주제 번호(1. 2. 3.)는 줄 맨 앞에 작성하세요.
- 들여쓰기는 하위 레벨로 갈수록 스페이스 4칸씩 증가시키세요.
- 각 항목은 최대 2줄을 넘기지 마세요. 간결함이 최우선입니다.
"""


def generate_with_gemini(transcript, meeting_info, config):
    """Google Gemini API로 회의록 생성 (다중 키 자동 로테이션)"""
    if not GEMINI_AVAILABLE:
        raise ImportError("google-genai 패키지가 설치되지 않았습니다.\npip install google-genai")
    
    keys = _get_all_keys(config, "gemini")
    if not keys:
        raise ValueError("Gemini API 키가 설정되지 않았습니다.")
    
    model_name = config.get("gemini_model", "gemini-2.0-flash")
    
    prompt = MEETING_PROMPT.format(
        meeting_name=meeting_info.get("회의명", ""),
        location=meeting_info.get("장소", ""),
        date=meeting_info.get("일시", ""),
        author=meeting_info.get("작성자", ""),
        attendees=meeting_info.get("참석자", ""),
        company=meeting_info.get("업체이름", ""),
        transcript=transcript,
    )
    
    last_error = None
    for idx, api_key in enumerate(keys):
        try:
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
            )
            # 성공한 키를 맨 앞으로 이동시켜 다음에 우선 사용
            if idx > 0:
                _rotate_key_to_front(config, "gemini", api_key)
            return response.text
        except Exception as e:
            last_error = e
            err_msg = str(e)
            # 할당량 초과(429), 인증 오류(401/403), 키 무효(400)면 다음 키로
            if any(code in err_msg for code in ["429", "RESOURCE_EXHAUSTED", "401", "403", "400", "API_KEY_INVALID"]):
                print(f"[키 로테이션] 키 #{idx+1} 실패 ({err_msg[:80]}...), 다음 키 시도")
                continue
            else:
                raise  # 다른 종류의 오류는 바로 raise
    
    raise Exception(f"모든 Gemini API 키({len(keys)}개)가 실패했습니다.\n마지막 오류: {last_error}")


def _rotate_key_to_front(config, provider, successful_key):
    """성공한 키를 목록 맨 앞으로 이동하고 설정 저장"""
    if provider == "gemini":
        key_field = "gemini_api_key"
        keys_field = "gemini_api_keys"
    else:
        key_field = "openai_api_key"
        keys_field = "openai_api_keys"
    
    config[key_field] = successful_key
    keys = config.get(keys_field, [])
    if successful_key in keys:
        keys.remove(successful_key)
    keys.insert(0, successful_key)
    config[keys_field] = keys
    try:
        save_config(config)
    except Exception:
        pass


def generate_with_openai(transcript, meeting_info, config):
    """OpenAI GPT API로 회의록 생성 (다중 키 자동 로테이션)"""
    if not OPENAI_AVAILABLE:
        raise ImportError("openai 패키지가 설치되지 않았습니다.\npip install openai")
    
    keys = _get_all_keys(config, "openai")
    if not keys:
        raise ValueError("OpenAI API 키가 설정되지 않았습니다.")
    
    prompt = MEETING_PROMPT.format(
        meeting_name=meeting_info.get("회의명", ""),
        location=meeting_info.get("장소", ""),
        date=meeting_info.get("일시", ""),
        author=meeting_info.get("작성자", ""),
        attendees=meeting_info.get("참석자", ""),
        company=meeting_info.get("업체이름", ""),
        transcript=transcript,
    )
    
    last_error = None
    for idx, api_key in enumerate(keys):
        try:
            client = openai.OpenAI(api_key=api_key)
            response = client.chat.completions.create(
                model=config.get("openai_model", "gpt-4o-mini"),
                messages=[
                    {"role": "system", "content": "당신은 전문 회의록 작성 비서입니다."},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.3,
            )
            if idx > 0:
                _rotate_key_to_front(config, "openai", api_key)
            return response.choices[0].message.content
        except Exception as e:
            last_error = e
            err_msg = str(e)
            if any(code in err_msg for code in ["429", "rate_limit", "401", "403", "400", "API_KEY_INVALID", "invalid_api_key"]):
                print(f"[키 로테이션] 키 #{idx+1} 실패 ({err_msg[:80]}...), 다음 키 시도")
                continue
            else:
                raise
    
    raise Exception(f"모든 OpenAI API 키({len(keys)}개)가 실패했습니다.\n마지막 오류: {last_error}")


def generate_ai_meeting_minutes(transcript, meeting_info, config=None):
    """
    AI를 사용하여 회의록 자동 생성
    
    Args:
        transcript: 음성 변환된 텍스트
        meeting_info: 회의 정보 딕셔너리
        config: AI 설정 (없으면 파일에서 로드)
    
    Returns:
        AI가 생성한 구조화된 회의록 텍스트
    """
    if config is None:
        config = load_config()
    
    provider = config.get("ai_provider", "gemini")
    
    if provider == "gemini":
        return generate_with_gemini(transcript, meeting_info, config)
    elif provider == "openai":
        return generate_with_openai(transcript, meeting_info, config)
    else:
        raise ValueError(f"지원하지 않는 AI 제공자: {provider}")


def get_available_providers():
    """사용 가능한 AI 제공자 목록 반환"""
    providers = []
    if GEMINI_AVAILABLE:
        providers.append("gemini")
    if OPENAI_AVAILABLE:
        providers.append("openai")
    return providers
