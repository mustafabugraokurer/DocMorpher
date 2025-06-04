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

## Türkçe Kullanım

DocMorpher ile TIF, PDF, UDF, DOCX, XLSX ve PPTX dosyalarını kolayca PDF, Word
ve görsel formatlara dönüştürebilirsiniz.

### Gereksinimler

Python 3.8 veya daha yeni bir sürüm ve şu paketlerin kurulması gerekir:

```
pip install pytesseract pdf2docx ocrmypdf comtypes pillow python-docx pymupdf numpy
```

### Harici Araçlar

- **Tesseract OCR** – OCR işlemleri için gereklidir. `gui_converter.py` içinde
`pytesseract.pytesseract.tesseract_cmd` değişkeninin `tesseract` yürütülebilirine
işaret ettiğinden emin olun.
- **Microsoft Office** – DOCX/XLSX/PPTX dosyalarını PDF'e çevirmek için
gereklidir.

### Çalıştırma

1. Tesseract ve Microsoft Office'in kurulu ve yapılandırılmış olduğundan emin olun.
2. Yukarıda listelenen Python paketlerini kurun.
3. Uygulamayı şu komutla başlatın:

```
python gui_converter.py
```

Kaynak dosyayı seçin, çıktı klasörünü ve hedef formatı belirleyin; ardından
**Dönüştür** düğmesine tıklayarak dönüştürmeyi başlatın.
