import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

# Path to Tesseract executable (if needed)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def pdf_photo_to_text(pdf_path):
    # Open the PDF file
    doc = fitz.open(pdf_path)
    text = ""

    # Loop through each page in the PDF
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        
        # Extract images from each page
        for image_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            
            # Convert image bytes to a PIL image
            image = Image.open(io.BytesIO(image_bytes))

            # Apply OCR to the image
            page_text = pytesseract.image_to_string(image)

            # Append the OCR result to the text
            text += f"Page {page_num + 1}, Image {image_index + 1}:\n"
            text += page_text + "\n\n"
    
    return text

# Usage example
pdf_path = "your_photo_pdf.pdf"  # Replace with the path to your photo PDF
extracted_text = pdf_photo_to_text(pdf_path)

# Save the extracted text to a text file
with open("output_text.txt", "w") as text_file:
    text_file.write(extracted_text)

print("Text extraction from photo PDF completed!")
