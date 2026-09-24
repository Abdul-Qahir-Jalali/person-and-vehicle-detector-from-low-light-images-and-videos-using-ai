import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def extract_headings(path):
    doc = Document(path)
    
    for i, p in enumerate(doc.paragraphs):
        style_name = p.style.name.lower()
        if "heading" in style_name and "toc" not in style_name:
            print(f"[{i}] {style_name}: {p.text}")

if __name__ == "__main__":
    extract_headings(sys.argv[1])
