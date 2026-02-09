# Docling

Claude API를 활용한 통합 문서 텍스트 & 이미지 추출기

## 지원 형식

| 형식 | 처리 방식 | API 필요 |
|------|-----------|----------|
| PDF (.pdf) | PyMuPDF 이미지 + EasyOCR (무료) | X |
| PowerPoint (.pptx) | python-pptx 직접 추출 | X |
| PowerPoint 구형 (.ppt) | LibreOffice → PPTX 변환 | X |
| Word (.docx) | python-docx 직접 추출 | X |
| Word 구형 (.doc) | LibreOffice → DOCX 변환 | X |
| 한글 (.hwp) | LibreOffice → DOCX/PDF 변환 | X |
| Excel (.xlsx) | openpyxl 직접 추출 | X |
| Excel 구형 (.xls) | LibreOffice → XLSX 변환 | X |

## 설치

```bash
pip3 install -r requirements.txt
```

### 추가 요구사항

- **PPT/DOC/XLS/HWP 구형 형식**: LibreOffice 필요
  ```bash
  brew install --cask libreoffice
  ```

- **PDF OCR**: EasyOCR 자동 설치 (최초 실행 시 모델 다운로드)

## 사용법

```bash
python3 doc_extract.py <파일경로> [출력디렉토리]
```

### 예시

```bash
python3 doc_extract.py document.pdf
python3 doc_extract.py presentation.pptx
python3 doc_extract.py report.docx
python3 doc_extract.py data.xlsx
python3 doc_extract.py 한글문서.hwp
python3 doc_extract.py legacy.ppt
python3 doc_extract.py legacy.doc
python3 doc_extract.py legacy.xls
```

## 출력 구조

```
<파일명>_extracted/
├── texts.json    # 텍스트 (JSON 구조화)
├── images.json   # 이미지 메타데이터 + ref 정보
└── images/       # 추출된 이미지 파일들
```

### 형식별 texts.json 구조

- **PPT/PPTX**: `slides[]` - 슬라이드별 텍스트/표/이미지
- **DOC/DOCX/HWP**: `sections[]` - 헤딩 기반 섹션별 텍스트/표/이미지
- **XLS/XLSX**: `sheets[]` - 시트별 데이터(헤더+행), merged cells
- **PDF**: `pages[]` - 페이지별 OCR 텍스트/이미지

## 주요 기능

- PDF: EasyOCR을 활용한 무료 텍스트 추출 (한글, 영어 지원)
- PPT/PPTX: 슬라이드별 텍스트, 표, 이미지 추출
- DOC/DOCX: 섹션별 단락, 표, 인라인 이미지 추출
- XLS/XLSX: 시트별 데이터, 헤더 자동 감지, 이미지 추출
- HWP: LibreOffice 변환 후 DOCX/PDF 파이프라인 활용
- 표는 구조화된 JSON 형식으로 변환
- 이미지에 주변 텍스트 기반 ref/title 정보 포함
- 완전 무료: API 키 불필요
