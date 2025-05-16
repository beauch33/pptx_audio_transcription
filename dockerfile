# Use the official lightweight Python image
FROM python:3.11-slim

# Set the working directory
WORKDIR /app

# Copy the local code to the container
COPY . /app

# Install FFmpeg for audio processing and Git for cloning
RUN apt-get clean && apt-get update -o Acquire::ForceIPv4=true && \
    apt-get install -y ffmpeg git --fix-missing && \
    rm -rf /var/lib/apt/lists/*

# Install pip dependencies
RUN pip install --upgrade pip

# Install the Python wrapper for ffmpeg
RUN pip install ffmpeg-python

# Install Whisper directly from GitHub with retries
RUN pip install --default-timeout=100 --retries=5 --no-cache-dir git+https://github.com/openai/whisper.git

# Expose port (optional if you want to use Flask later)
EXPOSE 5000

# Run the script when the container starts
CMD ["python", "pptx_audio_transc_onefile.py"]
