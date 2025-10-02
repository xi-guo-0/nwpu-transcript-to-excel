# NWPU Transcript to Excel

Convert NWPU transcript PDFs (Chinese and English) into the official Excel upload template.

## Features
- Parses both Chinese and English transcript PDFs using `pdfplumber`.
- Preserves the template workbook’s styles and metadata (writes only `sheet1.xml` and `sharedStrings.xml`).
- Outputs separate Excel files for CN/EN as needed.

## Install
From source (recommended while iterating locally):

```
python -m venv .venv
. .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -e .
```

This installs a console command `nwpu-transcript`.

> If you can’t or don’t want to install, see “Run from source” below.

## Usage
```
nwpu-transcript \
  --template 课程分学期模版.xlsx \
  --chinese path/to/your_chinese_transcript.pdf \
  --english path/to/your_english_transcript.pdf
```

- Provide one or both of `--chinese` / `--english`. Each output is independent.
- Defaults:
  - `--template`: `课程分学期模版.xlsx` in the current directory
  - `--output-chinese`: `transcript_chinese.xlsx`
  - `--output-english`: `transcript_english.xlsx`

Examples:
- Only Chinese transcript:
  ```
  nwpu-transcript --template 课程分学期模版.xlsx --chinese 本科成绩单.pdf
  ```
- Only English transcript:
  ```
  nwpu-transcript --template 课程分学期模版.xlsx --english 本科成绩单_en.pdf
  ```

## Run from source
You can run without installing:

```
python -m nwpu_transcript.cli --template 课程分学期模版.xlsx --chinese your.pdf
```

## Requirements
- Python 3.9+
- `pdfplumber`

Install dependency manually if not installing the package:

```
pip install pdfplumber
```
