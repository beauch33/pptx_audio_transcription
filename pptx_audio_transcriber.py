import os
import zipfile
from pathlib import Path
import whisper
import shutil
import subprocess
import ffmpeg

# Initialize Whisper model
model = whisper.load_model("base")

# Paths
INPUT_FOLDER = 'C:/Users/beauc/Workspace/ppt_Tst'
OUTPUT_FOLDER = 'C:/Users/beauc/Workspace/transcription/output'


def extract_audio_from_pptx(pptx_path, output_folder):
    """ Extracts audio files from a PowerPoint presentation """
    temp_folder = Path(output_folder) / 'temp'
    if temp_folder.exists():
        shutil.rmtree(temp_folder)
    temp_folder.mkdir(parents=True, exist_ok=True)

    corrupted_files = []

    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        for file_info in zip_ref.infolist():
            if file_info.filename.startswith('ppt/media/') and file_info.filename.endswith(('.m4a', '.mp3', '.wav')):
                try:
                    zip_ref.extract(file_info, temp_folder)
                except zipfile.BadZipFile:
                    print(f"⚠️ Corrupted file skipped: {file_info.filename}")
                    corrupted_files.append(file_info.filename)

    if corrupted_files:
        print("\nThe following audio files were corrupted and could not be extracted:")
        for f in corrupted_files:
            print(f"- {f}")

    media_folder = temp_folder / 'ppt' / 'media'
    if not media_folder.exists():
        print("No media folder found.")
        return []

    audio_files = sorted(media_folder.glob('*.m4a'))
    return audio_files


def is_valid_audio(file_path):
    """ Check if an audio file is playable with ffmpeg """
    try:
        result = subprocess.run(
            ['ffmpeg', '-v', 'error', '-i', str(file_path), '-f', 'null', '-'],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        return result.returncode == 0
    except Exception as e:
        print(f"Error validating audio: {e}")
        return False


def get_audio_duration(file_path):
    """ Get the duration of an audio file using ffmpeg """
    try:
        probe = ffmpeg.probe(str(file_path))
        duration = float(probe['format']['duration'])
        return duration
    except Exception as e:
        print(f"⚠️ Could not retrieve duration for {file_path}: {e}")
        return 0.0


def is_silent_transcription(text):
    """ Check if the transcription is empty or near-empty """
    return len(text.strip()) == 0


def format_time(seconds):
    """ Format seconds to VTT timestamp """
    millis = int((seconds % 1) * 1000)
    seconds = int(seconds)
    minutes = seconds // 60
    hours = minutes // 60
    return f"{hours:02}:{minutes % 60:02}:{seconds % 60:02}.{millis:03}"


def transcribe_audio(audio_file, output_folder, slide_number):
    """ Transcribes audio using Whisper and saves to .txt and .vtt """
    
    # Validate audio before transcribing
    if not is_valid_audio(audio_file):
        print(f"⚠️ Skipping corrupted or unreadable audio file: {audio_file}")
        return
    
    # Check the duration before transcribing
    duration = get_audio_duration(audio_file)
    if duration < 1.0:
        print(f"⚠️ Skipping {audio_file} — duration too short: {duration} seconds.")
        return
    
    print(f"🔎 Transcribing {audio_file}...")

    # Perform transcription
    result = model.transcribe(str(audio_file))
    transcribed_text = result['text']

    # Check if transcription is empty
    if is_silent_transcription(transcribed_text):
        print(f"⚠️ Transcription was empty for: {audio_file}")
        return
    
    # File naming
    original_filename = audio_file.parent.parent.parent.name
    output_filename_txt = f"{original_filename}_slide_{slide_number}.txt"
    output_filename_vtt = f"{original_filename}_slide_{slide_number}.vtt"
    text_file = os.path.join(output_folder, output_filename_txt)
    vtt_file = os.path.join(output_folder, output_filename_vtt)

    # Write to .txt
    with open(text_file, 'w', encoding='utf-8') as f:
        f.write(transcribed_text)
    
    # Write to .vtt (WebVTT format)
    with open(vtt_file, 'w', encoding='utf-8') as f:
        f.write("WEBVTT\n\n")
        for segment in result['segments']:
            start = format_time(segment['start'])
            end = format_time(segment['end'])
            text = segment['text']
            f.write(f"{start} --> {end}\n{text}\n\n")

    print(f"✅ Transcription saved as {text_file} and {vtt_file}")


def main():
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    pptx_files = [f for f in os.listdir(INPUT_FOLDER) if f.endswith('.pptx')]

    if not pptx_files:
        print("No .pptx files found.")
        return

    for pptx_file in pptx_files:
        print(f"Processing {pptx_file}...")
        pptx_path = os.path.join(INPUT_FOLDER, pptx_file)
        
        # Create output folder for this PPTX
        original_filename = Path(pptx_file).stem
        pptx_output_folder = os.path.join(OUTPUT_FOLDER, original_filename)
        if not os.path.exists(pptx_output_folder):
            os.makedirs(pptx_output_folder)

        audio_files = extract_audio_from_pptx(pptx_path, pptx_output_folder)
        if not audio_files:
            print(f"No audio found in {pptx_file}.")
            continue

        for index, audio_file in enumerate(audio_files, 1):
            transcribe_audio(audio_file, pptx_output_folder, index)
        
        # Clean up temporary files
        temp_folder = os.path.join(pptx_output_folder, 'temp')
        if os.path.exists(temp_folder):
            shutil.rmtree(temp_folder)
        
        print(f"✅ Completed processing for {pptx_file}.")

    print("🚀 All files processed.")


if __name__ == "__main__":
    main()
