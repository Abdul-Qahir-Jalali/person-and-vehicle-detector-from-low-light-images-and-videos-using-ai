import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def fix_issue5(input_path, output_path):
    doc = Document(input_path)
    
    in_abbreviations = False
    
    # We will traverse through paragraphs
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

    paragraphs = []
    def extract_all_paragraphs(parent):
        for block in iter_block_items(parent):
            if isinstance(block, Paragraph):
                paragraphs.append(block)
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        extract_all_paragraphs(cell)
                        
    extract_all_paragraphs(doc)
    
    abbreviations_list = [
        "AI\t\tArtificial Intelligence",
        "ML\t\tMachine Learning",
        "DL\t\tDeep Learning",
        "CNN\t\tConvolutional Neural Network",
        "YOLO\t\tYou Only Look Once",
        "mAP\t\tMean Average Precision",
        "OpenCV\t\tOpen Source Computer Vision Library",
        "API\t\tApplication Programming Interface",
        "UI\t\tUser Interface"
    ]
    
    header_found = False
    header_paragraph = None
    
    for i, p in enumerate(paragraphs):
        txt = p.text.strip().lower()
        
        if "abbreviations" == txt:
            in_abbreviations = True
            header_found = True
            header_paragraph = p
            continue
            
        if in_abbreviations:
            # If we hit an empty line, maybe skip, but if we hit another main heading, break?
            # Usually Abbreviations is the last thing. Let's just clear ALL paragraphs after it
            # that have text.
            if txt:
                # Clear existing text
                for run in p.runs:
                    run.text = ""
                    
    # Now insert the new list right after the header
    if header_found:
        # We need to insert after header_paragraph, but inserting before the paragraph AFTER header is easier.
        idx = paragraphs.index(header_paragraph)
        next_idx = idx + 1
        
        # Let's find the first non-empty style paragraph we can copy style from, or just use normal.
        style_to_use = None
        for j in range(next_idx, len(paragraphs)):
            if paragraphs[j].style:
                style_to_use = paragraphs[j].style
                break
                
        # Insert exactly at next_idx
        # But wait, insert_paragraph_before on the next paragraph is better
        insertion_point = None
        if next_idx < len(paragraphs):
            insertion_point = paragraphs[next_idx]
        else:
            insertion_point = doc.add_paragraph()
            
        for abbrev in abbreviations_list:
            new_p = insertion_point.insert_paragraph_before(abbrev)
            if style_to_use:
                new_p.style = style_to_use
    
    doc.save(output_path)
    print("Issue 5 Fixed and Saved")

if __name__ == "__main__":
    fix_issue5(sys.argv[1], sys.argv[2])
