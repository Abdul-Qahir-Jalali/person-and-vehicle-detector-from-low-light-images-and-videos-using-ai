import sys
import json
try:
    from docx import Document
except ImportError:
    print("Please run with: uv run --with python-docx script.py")
    sys.exit(1)

def dump_structure(input_path, output_path):
    doc = Document(input_path)
    structure = []
    
    # We will traverse all body elements (paragraphs and tables) in order
    # to preserve the exact document flow.
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
            raise ValueError("something's not right")

        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            if block.text.strip():
                structure.append({"type": "p", "text": block.text.strip()})
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if p.text.strip():
                            structure.append({"type": "cell_p", "text": p.text.strip()})
                            
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(structure, f, indent=2)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(1)
    dump_structure(sys.argv[1], sys.argv[2])
