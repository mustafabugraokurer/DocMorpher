# DocMorpher

DocMorpher — Convert TIF, PDF, UDF, DOCX, XLSX, PPTX and more into various formats like PDF, Word, and images using a smart and user-friendly GUI. TIF, PDF, DOCX, XLSX, PPTX ve UDF dosyalarını farklı formatlara dönüştürebilen, kullanıcı dostu ve kapsamlı bir belge dönüştürme aracı.

## Requirements

Install Python 3.8 or later and the following packages:

```
pip install pytesseract pdf2docx ocrmypdf comtypes pillow python-docx pymupdf numpy
```

These libraries provide OCR and Office conversion features used by the GUI.

## External Tools

- **Tesseract OCR** – Required for OCR functionality. Download from [https://github.com/tesseract-ocr/tesseract](https://github.com/tesseract-ocr/tesseract) and make sure `pytesseract.pytesseract.tesseract_cmd` in `gui_converter.py` points to the `tesseract` executable.
- **Microsoft Office** – Needed for DOCX/XLSX/PPTX to PDF conversions. The script uses the COM interface via `comtypes`, so Office must be installed on Windows.

## Running

1. Ensure Tesseract and Microsoft Office are installed and configured.
2. Install the required Python packages listed above.
3. Run the application with:

```
python gui_converter.py
```

Select the source file, choose an output directory and target format, then click **Dönüştür** to start the conversion.
