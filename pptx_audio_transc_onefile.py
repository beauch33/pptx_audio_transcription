import os
import zipfile
from pathlib import Path
import whisper
import shutil
import subprocess
import ffmpeg

# Initialize Whisper model
model = whisper.load_model("base")

# Check if we are running inside Docker
if os.path.exists('/app'):
    print("📦 Running inside Docker!")
    INPUT_FOLDER = '/app/ppt_Tst'
    OUTPUT_FOLDER = '/app/output'
else:
    print("💻 Running locally!")
    INPUT_FOLDER = 'C:/Users/beauc/Workspace/ppt_Tst'
    OUTPUT_FOLDER = 'C:/Users/beauc/Workspace/transcription/output'

# Make sure the directories exist
if not os.path.exists(INPUT_FOLDER):
    raise FileNotFoundError(f"❌ Input folder not found: {INPUT_FOLDER}")
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)


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


def transcribe_and_merge(audio_files, output_folder, original_filename):
    """ Transcribes all audio files and merges them into one .txt and .vtt """
    txt_path = os.path.join(output_folder, f"{original_filename}_transcription.txt")
    vtt_path = os.path.join(output_folder, f"{original_filename}_transcription.vtt")

    current_time = 0.0

    with open(txt_path, 'w', encoding='utf-8') as txt_file, open(vtt_path, 'w', encoding='utf-8') as vtt_file:
        vtt_file.write("WEBVTT\n\n")
        
        for index, audio_file in enumerate(audio_files, 1):
            if not is_valid_audio(audio_file):
                print(f"⚠️ Skipping corrupted or unreadable audio file: {audio_file}")
                continue

            duration = get_audio_duration(audio_file)
            if duration < 1.0:
                print(f"⚠️ Skipping {audio_file} — duration too short: {duration} seconds.")
                continue

            print(f"🔎 Transcribing {audio_file}...")
            result = model.transcribe(str(audio_file))
            transcribed_text = result['text']

            if is_silent_transcription(transcribed_text):
                print(f"⚠️ Transcription was empty for: {audio_file}")
                continue

            # Write to .txt
            txt_file.write(f"Slide {index}:\n{transcribed_text}\n\n")

            # Write to .vtt
            for segment in result['segments']:
                start = format_time(current_time + segment['start'])
                end = format_time(current_time + segment['end'])
                vtt_file.write(f"{start} --> {end}\n{segment['text']}\n\n")
            
            # Accumulate the current time to keep VTT in sync
            current_time += duration

    print(f"✅ Transcription saved as {txt_path} and {vtt_path}")


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

        transcribe_and_merge(audio_files, pptx_output_folder, original_filename)

        # Clean up temporary files
        temp_folder = os.path.join(pptx_output_folder, 'temp')
        if os.path.exists(temp_folder):
            shutil.rmtree(temp_folder)
        
        print(f"✅ Completed processing for {pptx_file}.")

    print("🚀 All files processed.")


if __name__ == "__main__":
    main()
