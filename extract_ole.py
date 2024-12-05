import olefile
from docx import Document
import os

def extract_ole_objects_from_docx(docx_path, output_dir):
    doc = Document(docx_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    ole_objects = []

    for rel in doc.part.rels.values():
        if "oleObject" in rel.reltype:
            ole_data = rel.target_part.blob
            ole_objects.append((rel.target_part.partname, ole_data))

    for i, (partname, ole_data) in enumerate(ole_objects):
        ole_path = os.path.join(output_dir, f"{partname.stem}_{i+1}.bin")
        with open(ole_path, "wb") as f:
            f.write(ole_data)
        extract_from_ole(ole_path, output_dir, i+1, partname.stem)

def extract_ole_objects_from_doc(doc_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    if olefile.isOleFile(doc_path):
        ole = olefile.OleFileIO(doc_path)
        for entry in ole.listdir():
            if entry[0] == 'WordDocument':
                ole_objects = ole.openstream(entry).read()
                ole_path = os.path.join(output_dir, f"{entry[0]}.bin")
                with open(ole_path, "wb") as f:
                    f.write(ole_objects)
                extract_from_ole(ole_path, output_dir, 1, entry[0])

def extract_from_ole(ole_path, output_dir, index, partname):
    if olefile.isOleFile(ole_path):
        ole = olefile.OleFileIO(ole_path)
        for entry in ole.listdir():
            match entry[0]:
                case 'WordDocument':
                    docx_path = os.path.join(output_dir, f"{partname}_nested_doc_{index}.docx")
                    with open(docx_path, "wb") as f:
                        f.write(ole.openstream(entry).read())
                    print(f"Extracted nested DOCX document {index} with name {partname}")
                case 'CONTENTS':
                    pdf_path = os.path.join(output_dir, f"{partname}_nested_doc_{index}.pdf")
                    with open(pdf_path, "wb") as f:
                        f.write(ole.openstream(entry).read())
                    print(f"Extracted nested PDF document {index} with name {partname}")
                case 'Workbook':  # Excel
                    xls_path = os.path.join(output_dir, f"{partname}_nested_doc_{index}.xls")
                    with open(xls_path, "wb") as f:
                        f.write(ole.openstream(entry).read())
                    print(f"Extracted nested XLS document {index} with name {partname}")
                case 'PowerPoint Document':
                    ppt_path = os.path.join(output_dir, f"{partname}_nested_doc_{index}.ppt")
                    with open(ppt_path, "wb") as f:
                        f.write(ole.openstream(entry).read())
                    print(f"Extracted nested PPT document {index} with name {partname}")

def main():
    file_path = "data/TS-Copy WBS Plan to Plan.docx"
    output_dir = "data/extracted_ole_objects"
    
    if file_path.endswith('.docx'):
        extract_ole_objects_from_docx(file_path, output_dir)
    elif file_path.endswith('.doc'):
        extract_ole_objects_from_doc(file_path, output_dir)
    else:
        print("Unsupported file format")