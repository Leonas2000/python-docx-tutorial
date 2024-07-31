import zipfile
import os
import shutil

# Path of your word template document
docx_file = 'Word Example.docx'
# How the output folder should be called and where to save it
extract_folder = 'extracted_output'

if os.path.exists(extract_folder):
    shutil.rmtree(extract_folder)

os.makedirs(extract_folder, exist_ok=True)

with zipfile.ZipFile(docx_file, 'r') as zip_ref:
    zip_ref.extractall(extract_folder)
