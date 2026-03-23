
# 📝 OCR-PDF-to-Word

<div align="center">

A **Python-powered OCR system** that transforms **scanned PDFs into fully editable Word documents**, preserving page structure and formatting. Ideal for digitizing documents, reports, and notes with high accuracy and efficiency.  

</div>

---

## 🚀 Project Overview

This OCR system works in a **step-by-step pipeline**:

1. **PDF to Images** : Each page of the PDF is converted into a high-resolution image using `pdf2image`.
2. **Image Preprocessing**: Images are converted to grayscale and binarized (black & white) to improve OCR accuracy.
3. **Text Extraction**: `pytesseract` reads the processed images and extracts text efficiently, handling multiple pages seamlessly.
4. **Word Document Creation**: Extracted text is added page-wise into a Word document (`.docx`) with proper page breaks.
5. **Performance Tracking**: Total processing time is displayed to track efficiency for large documents.

The pipeline ensures **reliable text extraction** from scanned PDFs with varying fonts, layouts, and quality.

---

## 🛠️ Technologies & Tools

`Python` , `pytesseract` , `pdf2image` , `PIL (Pillow)` , `python-docx` , `Image Preprocessing` , `PDF Handling` , `Text Extraction`

**External Requirements:**
- **Tesseract OCR**: Required for text recognition.  
- **Poppler**: Required by `pdf2image` to convert PDFs to images.  

---

## ⚡ Setup Instructions

### Clone the repository

```bash
git clone https://github.com/Sheraz-babar/OCR-PDF-to-Word.git
cd OCR-PDF-to-Word
```

### Install Required Libraries

```bash
pip install pdf2image pillow pytesseract python-docx
```

## 🔧 External Setup
This script requires two external tools to be installed on your system:

### 1. Tesseract OCR
#### Windows:

1. Download the installer from: GitHub Tesseract Releases https://github.com/UB-Mannheim/tesseract/wiki
2. Run the installer and note the installation path `(default: C:\Program Files\Tesseract-OCR\tesseract.exe`)
3. Update the tesseract_path variable in the script if you installed it elsewhere

#### macOS:
```bash
brew install tesseract
```

### 2. Poppler
Poppler is required for converting PDF pages to images.

Windows:

1. Download poppler from: poppler releases https://github.com/oschwartz10612/poppler-windows/releases/
2. Extract the zip file to a location like `C:\poppler`  
3. Add `C:\poppler\bin` to your system PATH environment variable:

   - Open System Properties → Environment Variables  
   - Edit the Path variable and add `C:\poppler\bin`  
   - Restart your terminal


#### macOS:
```bash
brew install poppler
```

## 📝 Usage

1. Update the file paths in the script:

   - **pdf_file** : Path to your input PDF file  
   - **output_docx** : Desired path for the output Word document  
   - **tesseract_path**: Path to your Tesseract executable (Windows only)

2. Run the script:
```bash
python code.py
```

## 📤 Output
The script generates:
1. A `.docx` file with extracted text
2. Each page is clearly marked with a page header
3. Progress messages in the console
4. Total processing time displayed upon completion

Example console output:
``` text
Processing page 1/10...
Page 1 done.

Processing page 2/10...
Page 2 done.

All pages processed! Text saved to 'output.docx'
Total time taken: 2 minutes 15 seconds
```

## ⚙️ Customization
You can modify the following parameters in the script:
1. **DPI** : Change `dpi = 300` to adjust image quality (higher = better OCR but slower)
2. **Threshold** : Modify 150 in `lambda x: 0 if x < 150 else 255` to adjust black/white sensitivity
3. **Tesseract Config** : Change `--psm 6` to use different page segmentation modes:
   - `--psm 3` : Fully automatic page segmentation (default)
   - `-- psm 4` : Assume a single column of text
   - `-- psm 6` : Assume a single uniform block of text
