# 🎤 회의록 자동 생성 프로그램 (AI 지원)

음성 파일(m4a)을 텍스트로 변환하고, **AI(Gemini/OpenAI)가 핵심 내용을 축약·정리**하여 Word 회의록 문서를 자동 생성하는 GUI 프로그램입니다.

## 🎯 주요 기능

| 기능 | 설명 |
|------|------|
| 🎤 음성→텍스트 | OpenAI Whisper 기반 한국어 음성 인식 |
| 🤖 AI 회의록 요약 | Google Gemini / OpenAI GPT로 핵심 축약 |
| 📄 Word 문서 생성 | KH VATEC 회의록 양식 자동 포맷팅 |
| 🖥️ GUI | tkinter 기반 (드래그앤드롭, 미리보기, 진행률 표시) |
| 👁️ 미리보기 | 저장 전 AI 요약 결과 확인 가능 |

## 🚀 설치 및 실행

### 1. 저장소 클론

```bash
git clone https://github.com/acientsoul/meeting-minutes-generator.git
cd meeting-minutes-generator
```

### 2. Python 패키지 설치

```bash
pip install -r requirements.txt
```

### 3. ffmpeg 설치 (Whisper 필수)

**Windows:**
```bash
winget install Gyan.FFmpeg
```

**Mac:**
```bash
brew install ffmpeg
```

**Linux:**
```bash
sudo apt install ffmpeg
```

> ⚠️ ffmpeg 설치 후 터미널을 재시작하세요.

### 4. GUI 실행

```bash
python gui_stable.py
```

> 첫 실행 시 Whisper 모델(~140MB)을 자동 다운로드합니다.

## 🤖 AI 설정 (선택사항)

AI 요약 기능을 사용하려면 API 키가 필요합니다.

### Google Gemini (무료 할당량 제공)
1. https://aistudio.google.com/apikey 에서 API 키 발급
2. GUI의 `⚙ AI 설정` 클릭 → Gemini 선택 → API 키 입력

### OpenAI GPT
1. https://platform.openai.com/api-keys 에서 API 키 발급
2. GUI의 `⚙ AI 설정` 클릭 → OpenAI 선택 → API 키 입력

> AI를 사용하지 않아도 원본 텍스트로 문서 생성 가능합니다.

## 📖 사용 방법

1. **회의 정보 입력** — 회의명, 장소, 일시, 작성자, 참석자, 업체이름
2. **음성 파일 선택** — m4a 파일을 버튼으로 선택 또는 드래그앤드롭
3. **변환 시작** — Whisper 음성 변환 → AI 축약 (자동 진행)
4. **미리보기 확인** — 오른쪽 패널에서 AI 요약 결과 확인
5. **문서 저장** — `💾 문서 저장` 버튼으로 Word 파일 저장

## 📁 파일 구조

```
├── gui_stable.py              # GUI 메인 프로그램
├── ai_meeting_generator.py    # AI 회의록 요약 모듈 (Gemini/OpenAI)
├── speech_to_text.py          # Whisper 음성→텍스트 변환
├── document_generator.py      # Word 문서 생성 (KH VATEC 양식)
├── main.py                    # CLI 버전
├── requirements.txt           # Python 패키지 목록
└── README.md
```

## ⚙️ Whisper 모델 옵션

| 모델 | 크기 | 속도 | 정확도 |
|------|------|------|--------|
| tiny | ~40MB | 매우 빠름 | 낮음 |
| **base** | ~140MB | **빠름 (기본값)** | **보통** |
| small | ~400MB | 보통 | 높음 |
| medium | ~1.4GB | 느림 | 매우 높음 |
| large | ~2.9GB | 매우 느림 | 최고 |

## 🐛 문제 해결

| 증상 | 해결 방법 |
|------|-----------|
| `ffmpeg not found` | ffmpeg 설치 후 터미널 재시작 |
| 음성 변환이 오래 걸림 | 40MB 파일 기준 CPU에서 10~30분 소요 (정상) |
| AI 요약 안됨 | `⚙ AI 설정`에서 API 키 확인 |
| `python` 명령 안됨 (Windows) | `python3` 또는 전체 경로 사용 |

## 📄 라이선스

이 프로그램은 개인 및 기업 용도로 자유롭게 사용 가능합니다.

---

**Version**: 2.0 | **Updated**: 2026.03.05
