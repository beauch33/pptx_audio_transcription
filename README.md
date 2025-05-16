PowerPoint Transcription with Whisper and Docker

This project allows you to extract audio from .pptx PowerPoint presentations, transcribe them using OpenAI's Whisper, and save the results as both .txt and .vtt files. It is fully containerized with Docker for portability and ease of deployment.

Project Structure

transcription/
├── Dockerfile                  - Docker image definition
├── pptx_audio_transc_onefile.py - Main transcription script
├── requirements.txt            - Python dependencies
├── run_transcriber.bat         - Windows batch file for easy execution
├── ppt_Tst/                    - Folder where you place your PowerPoint files
└── output/                     - Folder where transcriptions are saved

Setup Instructions

Build the Docker Image

Open your terminal and navigate to the project folder:

cd /path/to/transcription

Run the Docker build command:

docker build --no-cache -t pptx-transc-onefile .

Prepare Your Folders

Place all PowerPoint files (.pptx) inside:

/path/to/transcription/ppt_Tst

The script will automatically find all .pptx files in this folder and process them.

Transcriptions will be saved in:

/path/to/transcription/output

Running the Script

To run the Docker container and start transcribing:

docker run --rm -v /path/to/transcription/ppt_Tst:/app/ppt_Tst -v /path/to/transcription/output:/app/output pptx-transc-onefile

You should see transcriptions appear in the output/ folder as:

output/
├── presentation1_transcription.txt
├── presentation1_transcription.vtt
├── presentation2_transcription.txt
├── presentation2_transcription.vtt
└── ...

To Transfer to Work:

Save your Docker image to a file:

docker save -o pptx_transc_onefile.tar pptx-transc-onefile

Copy the pptx_transc_onefile.tar to a USB or cloud storage.

At work, load it:

docker load -i pptx_transc_onefile.tar

Run the Docker container as usual.

Optional: .bat File for One-Click Run

Create a run_transcriber.bat with the following:

@echo off
docker run --rm -v D:/PowerPoint_Audio/Input:/app/ppt_Tst -v D:/PowerPoint_Audio/Output:/app/output pptx-transc-onefile
pause

Double-click the .bat file and it will process all .pptx files automatically.

Next Steps

Test the Docker image at work.

Verify output files are created properly.

Confirm .bat script works for one-click execution.

Troubleshooting

FileNotFoundError → Make sure the path to ppt_Tst is correct and mounted properly.

Network Errors → If Docker can't connect, check VPN or firewall settings.

No Transcriptions in Output → Ensure the .pptx contains audio files.

For any issues, check the Docker logs:

docker logs [container_id]

Author: Your Name
Date: May 2025

