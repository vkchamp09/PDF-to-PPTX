import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
import os

# ---------- SETTINGS ----------
pdf_path = "ROLEX ANU .pdf"     # Your PDF file
output_ppt = "ROLEX.pptx"  # Output PPT file
dpi = 300  # Higher DPI for clarity

# ---------- PDF TO PPT ----------
doc = fitz.open(pdf_path)
prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]  # Blank layout

# Match slide ratio to PDF first page to avoid grey borders
first_page = doc[0]
rect = first_page.rect
pdf_width, pdf_height = rect.width, rect.height

# Set PowerPoint slide size in inches (keeping PDF ratio)
slide_width_in = 10
slide_height_in = 10 * (pdf_height / pdf_width)
prs.slide_width = Inches(slide_width_in)
prs.slide_height = Inches(slide_height_in)

for page_num in range(len(doc)):
    page = doc[page_num]
    
    # Render PDF page to image
    pix = page.get_pixmap(dpi=dpi)
    img_path = f"page_{page_num + 1}.png"
    pix.save(img_path)
    
    # Add new blank slide
    slide = prs.slides.add_slide(blank_slide_layout)
    
    # Insert the page image to fully fit the slide
    slide.shapes.add_picture(img_path, Inches(0), Inches(0),
                             width=prs.slide_width, height=prs.slide_height)
    
    # Remove the temporary image
    os.remove(img_path)

# Save PPT
prs.save(output_ppt)
print(f"✅ PDF successfully converted to {output_ppt} without grey borders!")
