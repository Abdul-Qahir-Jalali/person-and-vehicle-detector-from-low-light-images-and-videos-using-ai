import sys
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def verify(path):
    doc = Document(path)
    forbidden_words = ["deepfake", "audio", "wav2vec", "elevenlabs", "urdu"]
    found = []
    
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

    def check_parent(parent):
        for block in iter_block_items(parent):
            if isinstance(block, Paragraph):
                txt = block.text.lower()
                for w in forbidden_words:
                    if w in txt:
                        found.append((w, txt))
            elif isinstance(block, Table):
                for row in block.rows:
                    for cell in row.cells:
                        check_parent(cell)

    check_parent(doc)
    
    if not found:
        print("VERIFIED: Clean")
    else:
        print("FAILED. Found words:")
        for w, txt in found:
            print(f"Word '{w}' in '{txt}'")

if __name__ == "__main__":
    verify(sys.argv[1])
