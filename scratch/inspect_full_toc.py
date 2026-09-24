import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def inspect_full_toc(path):
    doc = Document(path)
    
    in_toc = False
    for i, p in enumerate(doc.paragraphs[:150]):
        txt = p.text.strip().lower()
        
        if "table of contents" == txt or "table of contents" in txt:
            in_toc = True
            
        if in_toc:
            print(f"[{i}] PARA (style: {p.style.name}): '{p.text}'")
            if "chapter 3" in txt:
                break

if __name__ == "__main__":
    inspect_full_toc(sys.argv[1])
