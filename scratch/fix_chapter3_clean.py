import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def fix_chapter3_clean_v2(input_path, output_path):
    doc = Document(input_path)
    
    start_idx = -1
    end_idx = -1
    
    paragraphs = list(doc.paragraphs)
    
    for i, p in enumerate(paragraphs):
        txt = p.text.strip().lower()
        
        if "3.4.2 detection result flow" in txt and "\t" not in txt:
            start_idx = i
            
        if "chapter 4" == txt and "\t" not in txt:
            end_idx = i
            break
            
    if start_idx != -1 and end_idx != -1:
        h3_style = paragraphs[start_idx].style
        
        for i in range(start_idx, end_idx):
            p = paragraphs[i]
            p_element = p._element
            parent = p_element.getparent()
            if parent is not None:
                parent.remove(p_element)
            
        new_content = [
            ("3.4.2 Detection Result Flow", h3_style),
            ("Once inference is complete, the results are formatted into a structured response containing object classes, confidence scores, and bounding box coordinates. The frontend parses this response to dynamically draw overlays on the media element presented to the user.", None),
            ("", None),
            ("3.4.3 Sequence Diagram", h3_style),
            ("[Placeholder: Sequence diagram showing the chronological interaction: User uploads media -> Flask Web App receives -> OpenCV Preprocessing Module enhances -> YOLO Deep Learning Model detects objects -> Flask returns Bounding Boxes and Confidence Scores to the interface.]", None),
            ("", None),
            ("3.4.4 Activity Diagram", h3_style),
            ("[Placeholder: Activity diagram showing the step-by-step decision flow of the object detection process, from media upload to bounding box rendering.]", None),
            ("", None),
            ("3.4.5 System Overview Diagram", h3_style),
            ("[Placeholder: High-level system overview diagram illustrating the client-server architecture, OpenCV enhancement module, and YOLO detection model.]", None),
            ("", None),
            ("3.4.6 System Architecture Diagram", h3_style),
            ("[Placeholder: Detailed system architecture diagram highlighting the specific technologies (Flask, OpenCV, YOLO) in their respective layers.]", None),
            ("", None),
            ("3.4.7 Use Case Diagram", h3_style),
            ("[Placeholder: Use case diagram showing the user interactions, including uploading an image/video, viewing the enhancement progress, and interpreting the final bounding boxes.]", None),
            ("", None)
        ]
        
        insert_target = paragraphs[end_idx]
        for text, style in new_content:
            new_p = insert_target.insert_paragraph_before(text)
            if style:
                new_p.style = style
                
    doc.save(output_path)
    print("Chapter 3 Cleaned and Saved")

if __name__ == "__main__":
    fix_chapter3_clean_v2(sys.argv[1], sys.argv[2])
