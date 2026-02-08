# Docling

Gemini API를 활용한 문서 텍스트 & 이미지 추출기

## 지원 형식

- PDF (.pdf)
- PowerPoint (.pptx)
- PowerPoint 구형 (.ppt) - LibreOffice 필요

## 설치

```bash
pip3 install -r requirements.txt
```

## 설정

1. `.env.sample`을 복사하여 `.env` 파일 생성
```bash
cp .env.sample .env
```

2. `.env` 파일에 본인의 Gemini API Key 입력
```
GEMINI_API_KEY=your_api_key_here
```

API Key는 [Google AI Studio](https://aistudio.google.com/apikey)에서 발급받을 수 있습니다.

## 사용법

```bash
python3 run.py <파일경로>
```

### 예시

```bash
python3 run.py document.pdf
python3 run.py presentation.pptx
python3 run.py legacy.ppt
```

## 출력 구조

```
<파일명>_output/
├── texts/
│   ├── page_001.md
│   ├── page_002.md
│   └── full_text.md
└── images/
    ├── page001_img01.png
    └── ...
```

## 주요 기능

- Gemini OCR을 활용한 텍스트 추출
- 표는 마크다운 표 형식으로 변환
- 제목/소제목 계층 구조 유지
- 임베디드 이미지 자동 추출
- API 오류 시 자동 재시도 (exponential backoff)
