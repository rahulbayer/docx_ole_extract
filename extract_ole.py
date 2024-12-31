import os
import pyzipper
import olefile
import glob


def extract_embedded_files_from_docx(docx_path, output_dir):
    """
    Extract embedded files from a .docx file and save them to the specified output directory.

    Parameters:
    docx_path (str): Path to the .docx file.
    output_dir (str): Directory where the extracted files will be saved.
    """
    # Create the output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Open the .docx file as a zip file
    with pyzipper.AESZipFile(docx_path, 'r') as docx_zip:
        # List all the files in the docx
        for file_info in docx_zip.infolist():
            # Check if the file is an embedded object
            if file_info.filename.startswith('word/embeddings/'):
                # Extract the embedded file
                embedded_file_name = os.path.basename(file_info.filename)
                output_file_path = os.path.join(output_dir, embedded_file_name)
                with open(output_file_path, 'wb') as f:
                    f.write(docx_zip.read(file_info.filename))
                print(f"Extracted {embedded_file_name} from {docx_path}")
            elif file_info.filename.endswith('.zip'):
                # Extract zip files directly
                zip_file_name = os.path.basename(file_info.filename)
                output_zip_path = os.path.join(output_dir, zip_file_name)
                with open(output_zip_path, 'wb') as f:
                    f.write(docx_zip.read(file_info.filename))
                print(f"Extracted zip file {zip_file_name} from {docx_path}")
                # Extract files from the zip archive
                extract_files_from_zip(output_zip_path, output_dir)

def extract_embedded_files_from_doc(doc_path, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    extract_original_file_name(doc_path) 
    if olefile.isOleFile(doc_path):
        ole = olefile.OleFileIO(doc_path)
        
        # Print all entries for debugging
        print("Entries in OLE file:")
        for entry in ole.listdir():
            
               if 'Package' in entry:
                   package_path = '/'.join(entry)
                   print (f"file Path{package_path}")
        # Extract from the Package entry
        try:
            
            if ole.exists(package_path):
                embedded_data = ole.openstream(package_path).read()
                name = package_path.split('/'[-1])
                print(f"original name {name}")
                print(f"Raw data from {package_path}: {embedded_data[:40]}")  # Print first 40 bytes for inspection
                
                # Check for valid Excel file signatures
                if embedded_data[:8] == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':  # For .xls
                    output_file_path = os.path.join(output_dir, "extracted_package.xls")
                    with open(output_file_path, "wb") as f:
                        f.write(embedded_data)
                    print(f"Extracted valid .xls file from Package: {output_file_path}")

                elif embedded_data[:4] == b'\x50\x4B\x03\x04':  # For .xlsx
                    output_file_path = os.path.join(output_dir, "extracted_package.xlsx")
                    with open(output_file_path, "wb") as f:
                        f.write(embedded_data)
                    print(f"Extracted valid .xlsx file from Package: {output_file_path}")

                else:
                    print("The data in /ObjectPool/_1584722410/Package does not appear to be a valid Excel file.")
            else:
                print("No valid Package found in the OLE file.")

        except Exception as e:
            print(f"Failed to extract from Package: {e}")

# Ensure to call this function in your main extraction routine
def extract_original_file_name(doc_path):
    if olefile.isOleFile(doc_path):
        ole = olefile.OleFileIO(doc_path)

        # Check for DocumentSummaryInformation
        if ole.exists('\x05DocumentSummaryInformation'):
            summary_info = ole.openstream('\x05DocumentSummaryInformation').read()
            print("Document Summary Information found.")
            # Extracting properties
            doc_properties = olefile.OleFileIO(doc_path).get_metadata()
            if doc_properties and doc_properties.title:
                print(f"Title: {doc_properties.title}")
            if doc_properties and doc_properties.subject:
                print(f"Subject: {doc_properties.subject}")
            if doc_properties and doc_properties.author:
                print(f"Author: {doc_properties.author}")
            if doc_properties and doc_properties.keywords:
                print(f"Keywords: {doc_properties.keywords}")

        # Check for SummaryInformation
        if ole.exists('\x05SummaryInformation'):
            summary_info = ole.openstream('\x05SummaryInformation').read()
            print("Summary Information found.")
            # Extracting properties
            summary_properties = olefile.OleFileIO(doc_path).get_metadata()
            if summary_properties and summary_properties.title:
                print(f"Title: {summary_properties.title}")
            if summary_properties and summary_properties.subject:
                print(f"Subject: {summary_properties.subject}")
            if summary_properties and summary_properties.author:
                print(f"Author: {summary_properties.author}")
            if summary_properties and summary_properties.keywords:
                print(f"Keywords: {summary_properties.keywords}")

        # Optionally, check for other entries that may contain file name information
        # For example, check for an entry that might contain the original file name
        for entry in ole.listdir():
            print(f"Entry: {entry}")

        ole.close()
    else:
        print("The specified file is not a valid OLE file.")

def extract_files_from_zip(zip_path, output_dir):
    """
    Extract files from a zip archive and save them to the specified output directory.

    Parameters:
    zip_path (str): Path to the zip file.
    output_dir (str): Directory where the extracted files will be saved.
    """
    with pyzipper.AESZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(output_dir)
        print(f"Extracted all files from {zip_path} to {output_dir}")

def determine_file_extension(entry_name):
    """
    Determine the file extension based on the entry name.

    Parameters:
    entry_name (str): The name of the entry.

    Returns:
    str: The corresponding file extension or None if unsupported.
    """
    print(entry_name)
    if 'Word.Document' in entry_name:
        if 'Word.Document.8' in entry_name or 'Word.Document.97' in entry_name:
            return '.doc'  # Older Word format
        else:
            return '.docx'  # Newer Word format
    elif 'CONTENTS' in entry_name:  # PDF
        return '.pdf'
    elif 'Workbook' in entry_name:  # Excel (both .xls and .xlsx)
        return '.xls'  # You may want to handle .xlsx separately based on other checks
    elif 'Excel.Document' in entry_name:  # Excel .xlsx
        return '.xlsx'
    elif 'Excel.Sheet' in entry_name:  # Excel .xlsx
        return '.xlsx'
    elif 'Data' in entry_name:  # Excel .xlsx
        return '.xlsx'
    elif 'PowerPoint Document' in entry_name:
        return '.ppt'
    elif 'Ole10Native' in entry_name:  # Zip
        return '.zip'
    else:
        return None
def extract_from_bin(bin_file_path, output_dir):
    """
    Extract actual files from a .bin file.

    Parameters:
    bin_file_path (str): Path to the .bin file.
    output_dir (str): Directory where the extracted files will be saved.
    """
    if olefile.isOleFile(bin_file_path):
        ole = olefile.OleFileIO(bin_file_path)
        
        # Iterate through the entries in the OLE file
        for entry in ole.listdir():
            # Skip unsupported entry types
            if entry[0] in ['CompObj', 'ObjInfo', 'DocumentSummaryInformation', 'SummaryInformation']:
                print(f"Skipping unsupported entry type: {entry[0]}")
                continue
            
            # Determine the correct file extension based on the entry type
            extension = determine_file_extension(entry[0])
            if extension:
                try:
                    embedded_data = ole.openstream(entry).read()
                    output_file_path = os.path.join(output_dir, f"{os.path.basename(bin_file_path)}{extension}")
                    with open(output_file_path, "wb") as f:
                        f.write(embedded_data)
                    print(f"Extracted {entry[0]} from {bin_file_path} as {extension}")
                    extract_files_from_zip(output_file_path, output_dir)
                except Exception as e:
                    print(f"Failed to extract {entry[0]} from {bin_file_path}: {e}")
def delete_unneccessary_files(folder_path, file_extensions):
    """
    Deletes all files of specific types from a specified folder.

    :param folder_path: Path to the folder from which to delete files.
    :param file_extensions: A list of file extensions of the files to delete (e.g., ['.txt', '.log']).
    """
    for file_extension in file_extensions:
        # Create the search pattern for the specified file type
        search_pattern = os.path.join(folder_path, f'*{file_extension}')
        
        # Use glob to find all files matching the pattern
        files_to_delete = glob.glob(search_pattern)

        # Delete each file found
        for file_path in files_to_delete:
            try:
                os.remove(file_path)
                print(f'Deleted: {file_path}')
            except Exception as e:
                print(f'Error deleting {file_path}: {e}')
                              
       
                                               

               
                                                        
       
                                
            
                                
                                         
                              
                                                       

def main():
    """
    Main function to execute the extraction process based on the file type.
    """
    # Prompt the user for the path to the Word document
    file_path = input("Enter the path of the Word document (.docx or .doc): ")
    output_dir = "data/extracted_embedded_files"

    # Check if the specified file exists
    if not os.path.exists(file_path):
        print("The specified file does not exist.")
        return

    # Determine the file type and call the appropriate extraction function
    if file_path.endswith('.docx'):
        extract_embedded_files_from_docx(file_path, output_dir)
    elif file_path.endswith('.doc'):
        extract_embedded_files_from_doc(file_path, output_dir)
    else:
        print("Unsupported file format")

    # After extraction, check for any .bin files and extract from them
    bin_files = [f for f in os.listdir(output_dir) if f.endswith('.bin')]
    if bin_files:
        print("Found .bin files. Extracting actual files from .bin...")
        for bin_file in bin_files:
            extract_from_bin(os.path.join(output_dir, bin_file), output_dir)
       

    delete_unneccessary_files(output_dir,[ '.zip'])                           
                                                                                                                                   
                                   

if __name__ == "__main__":
    main()