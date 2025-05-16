import zipfile
import os
import whisper
import shutil

# === UNZIP FUNCTION === #
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

# === AUTO DETECT FUNCTION === #
def auto_detect_pptx(folder_path):
    """
    If the user provides a folder instead of a file, this function will:
    1. Look for .pptx files in that folder.
    2. If one file is found, return its path.
    3. If multiple files are found, prompt the user to choose.
    """
    if not os.path.isdir(folder_path):
        print(f"Provided path is not a folder: {folder_path}")
        return None
    
    pptx_files = [f for f in os.listdir(folder_path) if f.endswith('.pptx')]
    
    if len(pptx_files) == 0:
        print(f"No .pptx files found in: {folder_path}")
        return None
    elif len(pptx_files) == 1:
        print(f"Found .pptx file: {pptx_files[0]}")
        return os.path.join(folder_path, pptx_files[0])
    else:
        print("Multiple .pptx files found:")
        for idx, file in enumerate(pptx_files, 1):
            print(f"{idx}. {file}")
        choice = int(input("Select the file number to process: ")) - 1
        if 0 <= choice < len(pptx_files):
            return os.path.join(folder_path, pptx_files[choice])
        else:
            print("Invalid choice.")
            return None


# === EXTRACT AUDIO FUNCTION === #
def extract_audio_from_pptx(unzipped_folder):
    """
    Extracts audio files from the 'ppt/media' folder of an unzipped PowerPoint.
    """
    if not os.path.exists(unzipped_folder):
        print(f"Path not found: {unzipped_folder}")
        return []

    # Prepare the media folder for output
    media_path = 'media'
    os.makedirs(media_path, exist_ok=True)

    # Locate the media folder inside the unzipped directory
    target_folder = os.path.join(unzipped_folder, 'ppt', 'media')

    media_files = []
    if os.path.exists(target_folder):
        print("Found media files... extracting.")
        for root, _, files in os.walk(target_folder):
            for file in files:
                if file.endswith(('.m4a', '.mp3', '.wav')):
                    source = os.path.join(root, file)
                    destination = os.path.join(media_path, file)
                    shutil.copy2(source, destination)
                    media_files.append(destination)
        print(f"Audio files extracted: {media_files}")
    else:
        print("No media folder found inside the provided path. Please check the structure.")

    return media_files


# === WHISPER TRANSCRIPTION FUNCTION === #
def transcribe_audio_files(audio_files):
    """
    Transcribes audio files using Whisper and saves .txt and .vtt files.
    """
    model = whisper.load_model("base")
    
    for file in audio_files:
        print(f"\n🔎 Transcribing {file}...")

        # 🔥 Convert to absolute path
        abs_path = os.path.abspath(file)
        print(f"🔎 Absolute path resolved: {abs_path}")

        # 🔎 Verify if the file exists
        if os.path.exists(abs_path):
            print(f"✅ File exists: {abs_path}")
        else:
            print(f"❌ File not found: {abs_path}")
            print("🔍 Listing contents of media folder:")
            media_folder = os.path.dirname(abs_path)
            for item in os.listdir(media_folder):
                print(f"  - {item}")
            continue

        # 🔎 Try transcribing
        try:
            result = model.transcribe(abs_path)
        except FileNotFoundError as e:
            print(f"Failed to transcribe {file}: {e}")
            continue

        # Generate text and VTT filenames
        text_path = abs_path.replace('.m4a', '.txt')
        vtt_path = abs_path.replace('.m4a', '.vtt')

        # Save transcription as .txt
        with open(text_path, 'w') as f:
            f.write(result['text'])

        # Save transcription as .vtt
        with open(vtt_path, 'w') as f:
            f.write("WEBVTT\n\n")
            for segment in result['segments']:
                start = segment['start']
                end = segment['end']
                text = segment['text']
                f.write(f"{format_time(start)} --> {format_time(end)}\n{text}\n\n")

        print(f"Transcription complete: {text_path}, {vtt_path}")


# === TIME FORMAT FUNCTION === #
def format_time(seconds):
    """
    Formats time in seconds to HH:MM:SS,ms format for VTT files.
    """
    ms = int((seconds % 1) * 1000)
    seconds = int(seconds)
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{hours:02}:{minutes:02}:{seconds:02}.{ms:03}"


# === MAIN FUNCTION === #
def main():
    path = input("Enter the path to your PowerPoint (.pptx) file or its folder: ")

    # If a folder is provided, try to auto-detect the .pptx file
    if os.path.isdir(path):
        pptx_file = auto_detect_pptx(path)
    else:
        pptx_file = path
    
    # Proceed if we have a valid .pptx file
    if pptx_file:
        unzipped_path = unzip_pptx(pptx_file)
        if unzipped_path:
            audio_files = extract_audio_from_pptx(unzipped_path)
            if audio_files:
                transcribe_audio_files(audio_files)
            else:
                print("No audio files found.")
    else:
        print("Could not find a valid .pptx file to process.")


if __name__ == "__main__":
    main()
