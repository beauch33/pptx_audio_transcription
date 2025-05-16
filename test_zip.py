import zipfile
import os

def unzip_pptx(pptx_path):
    """
    Unzips a .pptx PowerPoint file into a folder for extraction.
    Ignores CRC errors and logs the corrupted files.
    """
    if pptx_path.endswith('.pptx') and os.path.exists(pptx_path):
        print(f"Unzipping '{pptx_path}'...")
        extract_folder = pptx_path.replace('.pptx', '')

        # Create folder if not exists
        os.makedirs(extract_folder, exist_ok=True)

        corrupted_files = []

        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            for member in zip_ref.namelist():
                try:
                    zip_ref.extract(member, extract_folder)
                except zipfile.BadZipFile as e:
                    print(f"Corrupted file found: {member}")
                    corrupted_files.append(member)

        if corrupted_files:
            print("\nThe following files could not be extracted due to CRC errors:")
            for file in corrupted_files:
                print(f"- {file}")
        else:
            print("All files extracted successfully.")

        print(f"Unzipped to: {extract_folder}")
        return extract_folder
    else:
        print(f"File not found or not a .pptx file: {pptx_path}")
        return None
