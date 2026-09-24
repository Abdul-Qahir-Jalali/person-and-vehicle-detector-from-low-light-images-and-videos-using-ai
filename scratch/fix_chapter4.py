import sys
import re
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def fix_chapter4(input_path, output_path):
    doc = Document(input_path)
    
    in_chapter_4 = False
    
    for p in doc.paragraphs:
        txt = p.text.strip().lower()
        
        if "4.1 machine learning model" in txt or "chapter 4" in txt:
            in_chapter_4 = True
            
        if "5.1 home page overview" in txt or "chapter 5" in txt:
            in_chapter_4 = False
            
        if in_chapter_4:
            # Check for the specific 16,000 Hz line
            if "16,000" in txt or "hz" in txt or "resampled" in txt:
                # We need to replace this paragraph and insert the new steps
                # Wait, python-docx can just replace the text of this paragraph, and insert before/after
                
                # Clear the current paragraph
                p.text = "Image Resizing: Resizing image frames to standard dimensions (e.g., 640x640 pixels) to match the YOLO network's input layer requirements."
                
                # Insert the remaining steps after this paragraph
                # But insert_paragraph_before on the NEXT paragraph is easier
                # Let's just do it directly on this paragraph's element by inserting new paragraphs after it
                new_p1 = p.insert_paragraph_before("Pixel Normalization: Scaling pixel values between 0 and 1 to ensure consistent model convergence.")
                new_p2 = p.insert_paragraph_before("Tensor Conversion: Converting the processed image array into a format compatible with the PyTorch deep learning backend.")
                
                # Wait, insert_paragraph_before inserts BEFORE.
                # So if p.text is the first step, and I insert before it, the order is:
                # 1. Pixel Normalization
                # 2. Tensor Conversion
                # 3. Image Resizing
                # Let's fix the order!
                
                p.text = "Tensor Conversion: Converting the processed image array into a format compatible with the PyTorch deep learning backend."
                p.insert_paragraph_before("Pixel Normalization: Scaling pixel values between 0 and 1 to ensure consistent model convergence.")
                p.insert_paragraph_before("Image Resizing: Resizing image frames to standard dimensions (e.g., 640x640 pixels) to match the YOLO network's input layer requirements.")

    # General replacements across the entire document just to be absolutely certain
    replacements = [
        (r'(?i)waveform', 'visual features'),
        (r'(?i)\bHz\b', 'pixels'),
        (r'(?i)vocals', 'visual features'),
        (r'(?i)audio samples', 'image frames'),
        (r'(?i)sound waves', 'video frames')
    ]

    from docx.document import Document as _Document
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.table import CT_Tbl
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph

    def iter_block_items(parent):
        if isinstance(parent, _Document):
            parent_elm = parent.element.body
        elif isinstance(parent, _Cell):
            parent_elm = parent._tc
        else:
            return

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def process_text_replacements(parent):
        for block in iter_block_items(parent):
            if isinstance(block, Paragraph):
                original_text = block.text
                new_text = original_text
                for old, new in replacements:
                    new_text = re.sub(old, new, new_text)
                
                if new_text != original_text and len(block.runs) > 0:
                    block.runs[0].text = new_text
                    for i in range(1, len(block.runs)):
                        block.runs[i].text = ""
                        
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        process_text_replacements(cell)

    process_text_replacements(doc)

    doc.save(output_path)
    print("Chapter 4 Fixed and Saved")

if __name__ == "__main__":
    fix_chapter4(sys.argv[1], sys.argv[2])
