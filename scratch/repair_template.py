from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path

source=Path(r'E:\hammad project\project details\doccumentation\abdul qahir (53), raja sharyar muneer (83) fyp doccumentation.docx')
clean=Path(r'E:\hammad project\project details\doccumentation\Robust Object Detection System Documentation (Perfect LOF).docx')
target=Path(r'E:\hammad project\scratch\legacy_template_repaired.docx')
with ZipFile(source) as a, ZipFile(clean) as b, ZipFile(target,'w',ZIP_DEFLATED) as out:
    have=set(b.namelist())
    for item in a.infolist():
        blob=b.read(item.filename) if item.filename in have and item.filename.startswith('word/media/') else a.read(item.filename)
        out.writestr(item,blob)
print(target)
