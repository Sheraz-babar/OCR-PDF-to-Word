# 📝 OCR-PDF-to-Word

<div align="center">

A **Python-powered OCR system** that transforms **scanned PDFs into fully editable Word documents**, preserving page structure and formatting. Ideal for digitizing documents, reports, and notes with high accuracy and efficiency.  

</div>

---

## 🚀 Project Overview

This OCR system works in a **step-by-step pipeline**:

1. **PDF to Images** – Each page of the PDF is converted into a high-resolution image using `pdf2image`.
2. **Image Preprocessing** – Images are converted to grayscale and binarized (black & white) to improve OCR accuracy.
3. **Text Extraction** – `pytesseract` reads the processed images and extracts text efficiently, handling multiple pages seamlessly.
4. **Word Document Creation** – Extracted text is added page-wise into a Word document (`.docx`) with proper page breaks.
5. **Performance Tracking** – Total processing time is displayed to track efficiency for large documents.

The pipeline ensures **reliable text extraction** from scanned PDFs with varying fonts, layouts, and quality.

---

## 🛠️ Technologies & Tools

**Python, pytesseract, pdf2image, PIL (Pillow), python-docx, Image Preprocessing, PDF Handling, Text Extraction**  

**External Requirements:**
- **Tesseract OCR** – Required for text recognition.  
- **Poppler** – Required by `pdf2image` to convert PDFs to images.  

---

## ⚡ Setup Instructions

### 1️⃣ Clone the repository
```bash
git clone https://github.com/yourusername/OCR-PDF-to-Word.git
cd OCR-PDF-to-Word
