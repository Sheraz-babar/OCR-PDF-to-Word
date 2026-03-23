from pdf2image import convert_from_path
from PIL import Image, ImageOps
import pytesseract
from docx import Document
import time

pdf_file = "Dunham_Clinical.pdf"
output_docx = "Dunham_Clinical.docx"
tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pytesseract.pytesseract.tesseract_cmd = tesseract_path

# Start timer
start_time = time.time()

# Convert PDF to images
pages = convert_from_path(pdf_file, dpi=300)
total_pages = len(pages)

# Create Word document
doc = Document()

for i, page in enumerate(pages):
    print(f"Processing page {i+1}/{total_pages}...")

    # Preprocess image for OCR
    gray = ImageOps.grayscale(page)
    bw = gray.point(lambda x: 0 if x < 150 else 255, '1')

    # OCR
    text = pytesseract.image_to_string(bw, config='--psm 6')

    # Add to Word doc
    doc.add_paragraph(f"--- Page {i+1} ---")
    doc.add_paragraph(text)
    doc.add_page_break()

    print(f"Page {i+1} done.\n")

# Save Word file
doc.save(output_docx)

# End timer
end_time = time.time()
total_time = end_time - start_time

minutes = int(total_time // 60)
seconds = int(total_time % 60)

print(f"All pages processed! Text saved to '{output_docx}'")
print(f"Total time taken: {minutes} minutes {seconds} seconds")