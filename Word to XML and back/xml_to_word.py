import zipfile
import os

# Folder Path containing all the xml files
extract_folder = 'Simple Word'

# How the new word document should be called and where it is saved.
new_docx_file = 'reziped.docx'

with zipfile.ZipFile(new_docx_file, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
    for foldername, subfolders, filenames in os.walk(extract_folder):
        for filename in filenames:
            file_path = os.path.join(foldername, filename)
            archive_name = os.path.relpath(file_path, extract_folder)
            docx_zip.write(file_path, arcname=archive_name)

print(f"{new_docx_file} created successfully.")
