import sys
import xml.etree.ElementTree as ET

def get_text(node):
    text = []
    for child in node:
        if child.tag.endswith('}t'):
            if child.text:
                text.append(child.text)
        else:
            text.extend(get_text(child))
    return text

def parse_docx_xml(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()
    # Find all paragraphs
    text_content = []
    for p in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        p_text = "".join(get_text(p))
        if p_text.strip():
            text_content.append(p_text)
    return "\n".join(text_content)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(1)
    path = sys.argv[1]
    sys.stdout.reconfigure(encoding='utf-8')
    print(parse_docx_xml(path))
