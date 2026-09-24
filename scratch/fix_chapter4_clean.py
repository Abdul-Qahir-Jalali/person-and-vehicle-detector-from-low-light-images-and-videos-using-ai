import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def fix_chapter4_clean(input_path, output_path):
    doc = Document(input_path)
    
    start_idx = -1
    end_idx = -1
    
    paragraphs = list(doc.paragraphs)
    
    for i, p in enumerate(paragraphs):
        txt = p.text.strip().lower()
        
        # Locate Section 4.2.2 heading
        if "4.2.2 feature preparation" in txt and "\t" not in txt:
            start_idx = i
            continue
            
        if start_idx != -1 and end_idx == -1:
            # The next section in the original doc was 4.3 System Requirements or something? 
            # Or 4.3 Model Training. Let's just look for any "4.3" or "4.x"
            # In standard FYP format, it's probably "4.3" or another subsection.
            if txt.startswith("4.") and txt != "4.2.2 feature preparation":
                end_idx = i
                break
                
    if start_idx != -1 and end_idx != -1:
        # We will delete everything strictly between start_idx and end_idx.
        for i in range(start_idx + 1, end_idx):
            p = paragraphs[i]
            p_element = p._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)
                
        # The content to insert
        new_content = [
            "Final feature preparation involves normalizing the image dimensions and pixel values to match the input specifications required by the YOLO network.",
            "",
            "Image Resizing: Resizing image frames to standard dimensions (e.g., 640x640 pixels) to match the YOLO network's input layer requirements.",
            "",
            "Pixel Normalization: Scaling pixel values between 0 and 1 to ensure consistent model convergence.",
            "",
            "Tensor Conversion: Converting the processed image array into a format compatible with the PyTorch deep learning backend.",
            "",
            "This step ensures consistency across all inputs, whether they originate from high-resolution cameras or low-quality video streams.",
            ""
        ]
        
        insert_target = paragraphs[end_idx]
        for text in new_content:
            insert_target.insert_paragraph_before(text)
                
    doc.save(output_path)
    print("Chapter 4 Section 4.2.2 Cleaned and Saved")

if __name__ == "__main__":
    fix_chapter4_clean(sys.argv[1], sys.argv[2])
