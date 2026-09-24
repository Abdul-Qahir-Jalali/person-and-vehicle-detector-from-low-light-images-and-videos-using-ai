import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def fix_chapter56_ui(input_path, output_path):
    doc = Document(input_path)
    
    paragraphs = list(doc.paragraphs)
    
    # We want to iterate through paragraphs and find the exact strings.
    # Note that we only want to inject them ONCE.
    found_51 = False
    found_57 = False
    found_611 = False
    
    for i, p in enumerate(paragraphs):
        txt = p.text.strip().lower()
        
        # 1. Section 5.1 (Home Page Overview)
        if "screenshot example:" in txt and not found_51:
            # Let's verify we are around section 5.1
            # We can just check previous 10 paragraphs for 5.1
            is_51 = False
            for j in range(max(0, i-10), i):
                if "5.1 home page overview" in paragraphs[j].text.lower():
                    is_51 = True
                    break
            
            if is_51:
                # Insert below this paragraph. We do this by inserting BEFORE the next paragraph.
                if i + 1 < len(paragraphs):
                    paragraphs[i+1].insert_paragraph_before("[Placeholder: Screenshot of the Web App Home Page displaying the title 'Robust Object Detection System' and a central drag-and-drop area for uploading adverse weather images or videos.]")
                    found_51 = True
                    
        # 2. Section 5.7 (Interpreting Detection Results)
        if "screenshot examples:" in txt and not found_57:
            is_57 = False
            for j in range(max(0, i-10), i):
                if "5.7" in paragraphs[j].text.lower() or "interpreting detection" in paragraphs[j].text.lower():
                    is_57 = True
                    break
                    
            if is_57:
                if i + 1 < len(paragraphs):
                    paragraphs[i+1].insert_paragraph_before("[Placeholder: Screenshot showing the detection results interface. An image of a foggy road is displayed with brightly colored bounding boxes drawn around detected vehicles and pedestrians, with confidence scores like 92% listed next to them.]")
                    found_57 = True
                    
        # 3. Section 6.1.1 (Detection Performance)
        if "screenshot:" in txt and not found_611:
            is_611 = False
            for j in range(max(0, i-10), i):
                if "6.1.1" in paragraphs[j].text.lower() or "detection performance" in paragraphs[j].text.lower() or "6.1" in paragraphs[j].text.lower():
                    is_611 = True
                    break
                    
            if is_611:
                if i + 1 < len(paragraphs):
                    paragraphs[i+1].insert_paragraph_before("[Placeholder: A classification report table showing precision, recall, and f1-score for object classes such as 'Vehicles', 'Pedestrians', and 'Traffic Signs', reflecting high performance (e.g., 0.92) across all metrics.]")
                    found_611 = True

    # Check if we missed any due to slight formatting differences
    if not found_51 or not found_57 or not found_611:
        print(f"Warning: Did not find all sections. 5.1: {found_51}, 5.7: {found_57}, 6.1.1: {found_611}")
        
    doc.save(output_path)
    print("Chapter 5 and 6 UI Placeholders Injected and Saved")

if __name__ == "__main__":
    fix_chapter56_ui(sys.argv[1], sys.argv[2])
