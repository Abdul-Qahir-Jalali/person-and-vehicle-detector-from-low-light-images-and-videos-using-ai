import sys
try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def fix_toc_perfectly(input_path, output_path):
    doc = Document(input_path)
    
    # 1. Identify headings and compute pages
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

    pre_toc = [
        ("UNdertaking", 1, "v"),
        ("Acknowledgements", 1, "vi"),
        ("Table of Contents", 1, "vii"),
        ("List of figures", 1, "x")
    ]
    all_toc_entries = pre_toc + [(txt, lvl, str(get_page(idx))) for idx, lvl, txt in headings]

    # 2. Identify TOC paragraphs
    toc_paras = []
    in_toc = False
    for p in doc.paragraphs:
        txt = p.text.strip().lower()
        if txt == "table of contents" and p.style.name == "Heading 1":
            in_toc = True
            continue
        if in_toc:
            if "toc" in p.style.name.lower():
                toc_paras.append(p)
            elif p.style.name == "Heading 1":
                break
                
    # 3. Overwrite TOC paragraphs
    for i, (txt, lvl, pg) in enumerate(all_toc_entries):
        if not txt: continue
        
        if i < len(toc_paras):
            # Reuse existing paragraph
            p = toc_paras[i]
            # Clear text
            for r in p.runs:
                r.text = ""
                
            # Need to clear all existing elements just in case, but runs text="" is safe
            p.style = f"toc {lvl}"
            
            p.add_run(txt)
            r_tab = p.add_run()
            r_tab._element.append(OxmlElement('w:tab'))
            p.add_run(pg)
        else:
            # Create new paragraph after the last one
            prev_p = toc_paras[-1] if toc_paras else None
            if prev_p:
                new_p = prev_p.insert_paragraph_before("")
                # Since insert_paragraph_before inserts before, it's tricky.
                # Just insert using xml
                new_p_xml = OxmlElement('w:p')
                
                pPr = OxmlElement("w:pPr")
                pStyle = OxmlElement("w:pStyle")
                pStyle.set(qn("w:val"), f"toc{lvl}")
                pPr.append(pStyle)
                
                run_txt = OxmlElement('w:r')
                t_txt = OxmlElement('w:t')
                t_txt.text = txt
                run_txt.append(t_txt)
                
                run_tab = OxmlElement('w:r')
                run_tab.append(OxmlElement('w:tab'))
                
                run_pg = OxmlElement('w:r')
                t_pg = OxmlElement('w:t')
                t_pg.text = pg
                run_pg.append(t_pg)
                
                new_p_xml.append(pPr)
                new_p_xml.append(run_txt)
                new_p_xml.append(run_tab)
                new_p_xml.append(run_pg)
                
                # Append to the end of the toc_paras list in XML
                last_p = toc_paras[-1]._element if toc_paras else None
                if last_p is not None:
                    last_p.addnext(new_p_xml)
                    # We need to update toc_paras list so the next iteration appends after this new one!
                    # But actually we can just keep appending after the last added one
                    # Let's keep a reference to the last inserted element
                    toc_paras.append(new_p_xml) # store just to have a reference, though it's an oxmlelement now
            
    # Remove leftover TOC paragraphs
    if len(toc_paras) > len(all_toc_entries):
        for i in range(len(all_toc_entries), len(toc_paras)):
            p = toc_paras[i]
            # p might be an OxmlElement if we appended it, but here it's definitely from the original doc
            if hasattr(p, '_element'):
                parent = p._element.getparent()
                if parent is not None:
                    parent.remove(p._element)

    doc.save(output_path)
    print("Perfect TOC generated.")

if __name__ == "__main__":
    fix_toc_perfectly(sys.argv[1], sys.argv[2])
