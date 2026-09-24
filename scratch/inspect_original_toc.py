import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def inspect_original_toc(path):
    doc = Document(path)
    
    in_toc = False
    for i, p in enumerate(doc.paragraphs[:150]):
        txt = p.text.strip().lower()
        if "table of contents" == txt:
            in_toc = True
            
        if in_toc and "toc" in p.style.name.lower():
            runs_repr = [repr(r.text) for r in p.runs]
            print(f"[{i}] {p.style.name} | Text: {repr(p.text)} | Runs: {runs_repr}")
            if "chapter 3" in txt:
                break

if __name__ == "__main__":
    inspect_original_toc(sys.argv[1])
