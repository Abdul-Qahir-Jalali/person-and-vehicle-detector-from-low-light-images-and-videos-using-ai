import sys
try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.text.paragraph import Paragraph
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def rebuild_toc(input_path, output_path):
    doc = Document(input_path)
    
    # 1. Identify all headings in the body (after the TOC block)
    headings = []
    body_start = 0
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "Chapter 1" and p.style.name == "Heading 1":
            body_start = i
            break
            
    anchors = []
    for i in range(body_start, len(doc.paragraphs)):
        p = doc.paragraphs[i]
        style_name = p.style.name.lower()
        if "heading 1" in style_name or "heading 2" in style_name or "heading 3" in style_name:
            lvl = int(style_name.split()[-1])
            headings.append((i, lvl, p.text.strip()))
            
            text_lower = p.text.strip().lower()
            if text_lower == "chapter 1": anchors.append((i, 1))
            elif text_lower == "chapter 2": anchors.append((i, 5))
            elif text_lower == "chapter 3": anchors.append((i, 12))
            elif text_lower == "chapter 4": anchors.append((i, 24))
            elif text_lower == "chapter 5": anchors.append((i, 31))
            elif text_lower == "chapter 6": anchors.append((i, 37))
            elif text_lower == "chapter 7": anchors.append((i, 42))
            elif text_lower == "references": anchors.append((i, 46))
            elif text_lower == "abbreviations": anchors.append((i, 47))

    if not anchors:
        anchors = [(body_start, 1), (len(doc.paragraphs)-1, 50)]

    def get_page(para_idx):
        for j in range(len(anchors)-1):
            start_idx, start_pg = anchors[j]
            end_idx, end_pg = anchors[j+1]
            if start_idx <= para_idx <= end_idx:
                if end_idx == start_idx: return start_pg
                ratio = (para_idx - start_idx) / (end_idx - start_idx)
                return int(start_pg + ratio * (end_pg - start_pg))
        if para_idx >= anchors[-1][0]:
            return anchors[-1][1]
        return anchors[0][1]

    # 2. Find and delete the existing TOC paragraphs
    toc_start = -1
    toc_end = -1
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip().lower() == "table of contents" and p.style.name == "Heading 1":
            toc_start = i
        if toc_start != -1 and i > toc_start:
            if "toc" in p.style.name.lower():
                toc_end = i
            elif p.style.name.lower() == "heading 1": 
                break
                
    for i in range(toc_start + 1, toc_end + 1):
        p_element = doc.paragraphs[i]._element
        parent = p_element.getparent()
        if parent is not None:
            parent.remove(p_element)

    # 3. Insert new TOC paragraphs
    target_p = doc.paragraphs[toc_start] # "Table of Contents" heading
    
    pre_toc = [
        ("UNdertaking", 1, "v"),
        ("Acknowledgements", 1, "vi"),
        ("Table of Contents", 1, "vii"),
        ("List of figures", 1, "x")
    ]
    all_toc_entries = pre_toc + [(txt, lvl, str(get_page(idx))) for idx, lvl, txt in headings]
    
    # We will insert paragraphs directly into the XML tree after target_p
    current_element = target_p._element
    for txt, lvl, pg in all_toc_entries:
        if not txt: continue
        new_p = OxmlElement("w:p")
        # Give it the right style
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), f"toc{lvl}") # The internal XML name for style is usually toc1, toc2 without space
        pPr.append(pStyle)
        
        # Add tab stops (optional, Word handles it via style usually, but safe to just add a run with tab)
        run = OxmlElement("w:r")
        t_text = OxmlElement("w:t")
        t_text.text = txt
        run.append(t_text)
        
        run_tab = OxmlElement("w:r")
        tab = OxmlElement("w:tab")
        run_tab.append(tab)
        
        run_pg = OxmlElement("w:r")
        t_pg = OxmlElement("w:t")
        t_pg.text = pg
        run_pg.append(t_pg)
        
        new_p.append(pPr)
        new_p.append(run)
        new_p.append(run_tab)
        new_p.append(run_pg)
        
        current_element.addnext(new_p)
        current_element = new_p

    doc.save(output_path)
    print("Rebuilt TOC successfully.")

if __name__ == "__main__":
    rebuild_toc(sys.argv[1], sys.argv[2])
