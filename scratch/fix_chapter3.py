import sys
import re
try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def fix_diagrams(input_path, output_path):
    doc = Document(input_path)
    
    in_chapter_3 = False
    
    for p in doc.paragraphs:
        txt = p.text.strip().lower()
        
        # Track if we are in Chapter 3
        if "3.1 architecture overview" in txt or "chapter 3" in txt or "system architecture" in txt:
            in_chapter_3 = True
            
        if "4.1 machine learning model" in txt or "chapter 4" in txt:
            in_chapter_3 = False
            
        if in_chapter_3:
            # Check if this paragraph contains an image (drawing)
            has_drawing = False
            for run in p.runs:
                # Look for <w:drawing> in the run's XML
                if run._element.xpath('.//w:drawing'):
                    has_drawing = True
                    # Remove the drawing from the XML
                    for drawing in run._element.xpath('.//w:drawing'):
                        drawing.getparent().remove(drawing)
            
            # Now we need to insert the text-based diagrams for Chapter 3.
            # I can just append them at the relevant sections.
            # Instead of appending exactly where the image was, I can look for the section headers.
            
            # Data Flow / Activity Diagram
            if "3.4.1 data flow for object detection" in txt or "3.4.4 activity diagram" in txt:
                # Clear placeholder text
                p.text = "Data Flow & Activity Diagram (Text-Based Flow):"
                # Add steps
                steps = [
                    "Step 1: User uploads Image/Video via Flask Web Interface.",
                    "Step 2: Flask Web App receives media.",
                    "Step 3: Send media to OpenCV Preprocessing Module.",
                    "Step 4: OpenCV applies Dehazing, Low-Light Enhancement, and Noise Removal.",
                    "Step 5: Preprocessed frames are passed to the YOLO Deep Learning Model.",
                    "Step 6: YOLO performs inference to detect objects (Vehicles, Pedestrians, etc.).",
                    "Step 7: Generate Bounding Boxes, Class Labels, and Confidence Scores.",
                    "Step 8: Display the processed Image/Video with Bounding Box overlays on the web interface."
                ]
                for step in steps:
                    new_p = p.insert_paragraph_before(step)
                    
            elif "3.4.3 sequence diagram" in txt or "3.4.6 system architecture diagram" in txt:
                p.text = "Sequence & Architecture Components:"
                lines = [
                    "• Preprocessing: OpenCV Enhancement Module",
                    "• Detection Core: YOLO Object Detection Model",
                    "• Input Interface: Enhanced Image Frames",
                    "• Output Interface: Bounding Boxes & Confidence Scores"
                ]
                for line in lines:
                    new_p = p.insert_paragraph_before(line)
                    
            elif "3.4.7 use case diagram" in txt:
                p.text = "Use Case States:"
                lines = [
                    "• Initial State: Upload Image/Video Screen",
                    "• Final State: Detection Result Screen (Bounding Boxes)"
                ]
                for line in lines:
                    new_p = p.insert_paragraph_before(line)
                    
            # Also clear any remaining placeholders I added earlier
            if "[diagram placeholder:" in txt:
                p.text = ""

    # Ensure absolute eradication in tables as well (just in case)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    txt = p.text.strip().lower()
                    if in_chapter_3:
                        for run in p.runs:
                            if run._element.xpath('.//w:drawing'):
                                for drawing in run._element.xpath('.//w:drawing'):
                                    drawing.getparent().remove(drawing)

    # Let's also do a general text replacement across all elements just to be 100% sure
    # using the exact mapping they gave for sequence and architecture
    replacements = [
        (r'(?i)ElevenLabs API', 'OpenCV Enhancement Module'),
        (r'(?i)Wav2Vec2-Large-XLSR-53', 'YOLO Object Detection Model'),
        (r'(?i)Wav2Vec2', 'YOLO Object Detection Model'),
        (r'(?i)Audio Features', 'Enhanced Image Frames'),
        (r'(?i)Prediction \(Real/Fake\)', 'Bounding Boxes & Confidence Scores'),
        (r'(?i)Prediction Real/Fake', 'Bounding Boxes & Confidence Scores'),
        (r'(?i)Upload Audio Screen', 'Upload Image/Video Screen'),
        (r'(?i)Upload Urdu audio', 'Upload Image/Video'),
        (r'(?i)Result Screen \(Real/Fake\)', 'Detection Result Screen (Bounding Boxes)'),
        (r'(?i)Real/Fake', 'Bounding Boxes'),
        (r'(?i)Real or Fake', 'Detected Objects'),
        (r'(?i)Deepfake', 'Object Detection'),
        (r'(?i)Audio', 'Image/Video'),
        (r'(?i)Urdu', 'Adverse Weather')
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
    print("Chapter 3 Diagrams Replaced and Saved")

if __name__ == "__main__":
    fix_diagrams(sys.argv[1], sys.argv[2])
