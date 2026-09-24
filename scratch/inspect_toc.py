import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def inspect_toc(path):
    doc = Document(path)
    
    for i, p in enumerate(doc.paragraphs[:200]):
        if "3.4 data flow" in p.text.lower() or "step 1" in p.text.lower() or "chapter 4" in p.text.lower() or "sequence & architecture" in p.text.lower():
            print(f"[{i}] {p.text}")

if __name__ == "__main__":
    inspect_toc(sys.argv[1])
