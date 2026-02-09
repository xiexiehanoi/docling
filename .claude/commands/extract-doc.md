Extract text and images from documents.

Target file: $ARGUMENTS

## Step 1: Input Validation

First, verify:
- File exists (`ls -la "$ARGUMENTS"`)
- Extension is supported (.ppt, .pptx, .pdf, .doc, .docx, .hwp, .xls, .xlsx)

If file doesn't exist or format is unsupported, notify user immediately.
If it's a directory, find all supported files inside and extract each one.

## Step 2: Run Extraction

```
python3 /Users/gfw/Desktop/project/docling/doc_extract.py "$ARGUMENTS"
```

## Step 3: Verify Results

After extraction completes:
1. Read `texts.json` and summarize structure (item count, 2 text samples, image count & ref info)
2. Display `images/` folder file list

## Step 4: Result Summary Template

Report results in this format:

```
### Extraction Results
- **File**: <filename>
- **Format**: <extension>
- **Text**: <slides/pages/sections/sheets count> items
- **Images**: <image count> extracted
- **Output**: <output directory>

### Text Samples
<Preview of first 2 items>

### Image List
<Image filenames + ref info>
```

## Error Handling Guide

| Error Message | Solution |
|---------------|----------|
| `LibreOffice 필요` | Guide to run `brew install --cask libreoffice` |
| `python-pptx 필요` | Guide to run `pip install python-pptx` |
| `python-docx 필요` | Guide to run `pip install python-docx` |
| `openpyxl 필요` | Guide to run `pip install openpyxl` |
| `PyMuPDF 필요` | Guide to run `pip install PyMuPDF` |
| `EasyOCR 필요` (PDF only) | Guide to run `pip install easyocr` |
| `변환 실패` (PPT/DOC/XLS/HWP) | Check LibreOffice installation, verify file integrity |

## Supported Formats

| Format | Processing Method | API Required |
|--------|-------------------|--------------|
| .ppt | LibreOffice → PPTX → python-pptx | X |
| .pptx | python-pptx direct | X |
| .doc | LibreOffice → DOCX → python-docx | X |
| .docx | python-docx direct | X |
| .hwp | LibreOffice → DOCX/PDF | X |
| .xls | LibreOffice → XLSX → openpyxl | X |
| .xlsx | openpyxl direct | X |
| .pdf | PyMuPDF images + EasyOCR text (FREE) | X |

## Output Structure

```
<filename>_extracted/
├── texts.json    # Text + image metadata unified JSON
└── images/       # Extracted image files
```

### Image Filename Rules
Image filenames are auto-generated based on surrounding text/table headers:
- Images in tables: `{column_header} 이미지_슬라이드{N}.png` (e.g., `개선전 이미지_슬라이드3.png`)
- Near text: `{nearby_text} 이미지_슬라이드{N}.png`
- DOCX: `{ref} 이미지_섹션{N}.png`
- XLSX: `{ref} 이미지_시트{N}.png`
- PDF: `{ref} 이미지_페이지{N}.png`

### shape_id Rules
Each content item's shape_id is also descriptive:
- Text: `{content_summary}_슬라이드{N}`
- Image: same as image filename (without extension)
- Table: `표_{header_summary}_슬라이드{N}`

### Format-specific texts.json Structure

**PPT/PPTX**: `slides[]` - per-slide content (text, table, group, image)
**DOC/DOCX/HWP**: `sections[]` - heading-based content (text, table, image_ref)
**XLS/XLSX**: `sheets[]` - per-sheet data (headers + rows), merged_cells
**PDF**: `pages[]` - per-page content (OCR text, image_ref)

Note: texts.json does not include position information.
