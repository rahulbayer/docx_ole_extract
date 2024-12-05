import olefile
from docx import Document
import os

def extract_ole_objects(docx_path, output_dir):
    doc = Document(docx_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    ole_objects = []

    for rel in doc.part.rels.values():
        if "oleObject" in rel.reltype:
            ole_data = rel.target_part.blob
            ole_objects.append(ole_data)
    for i, ole_data in enumerate(ole_objects):
        ole_path = os.path.join(output_dir, f"{rel.target_part.partname}_{i+1}.bin")
        with open(ole_path, "wb") as f:
            f.write(ole_data)
        extract_from_ole(ole_path, output_dir, i+1)

def extract_from_ole(ole_path, output_dir, index):
    if olefile.isOleFile(ole_path):
        ole = olefile.OleFileIO(ole_path)
        for entry in ole.listdir():
            if entry[0] == 'WordDocument':
                docx_path = os.path.join(output_dir, f"nested_doc_{index}.docx")
                with open(docx_path, "wb") as f:
                    f.write(ole.openstream(entry).read())
                print(f"Extracted nested DOCX document {index}")
            elif entry[0] == 'CONTENTS':
                pdf_path = os.path.join(output_dir, f"nested_doc_{index}.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(ole.openstream(entry).read())
                print(f"Extracted nested PDF document {index}")

def main():
    docx_path = "data/TS-Copy WBS Plan to Plan.docx"
    output_dir = "data/extracted_ole_objects"
    extract_ole_objects(docx_path, output_dir)
