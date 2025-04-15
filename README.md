# Whisper Push-to-Talk

A Windows application that provides push-to-talk voice transcription using OpenAI's Whisper speech recognition model.

## Features

- Hold left Alt key to record speech, release to transcribe
- Automatically types transcribed text where your cursor is positioned
- Detects appropriate text fields for typing
- Runs silently in the background with system tray icon
- Maintains logs of all transcriptions
- Prevents multiple instances from running

## Requirements

- Windows 10/11
- Python 3.8+
- FFmpeg (included or in PATH)

## Installation

1. Clone this repository:

   ```
   git clone https://github.com/yourusername/PushToTalk.git
   cd PushToTalk
   ```

2. Install dependencies:

   ```
   pip install -r requirements.txt
   ```

3. Download FFmpeg:
   - Download from https://www.gyan.dev/ffmpeg/builds/ (get the "essentials" build)
   - Extract ffmpeg.exe from the bin folder to the same directory as the script

## Usage

1. Run the application:

   ```
   pythonw whisper_push_to_talk.pyw
   ```

2. Look for the microphone icon in your system tray (may be hidden in the "Show hidden icons" section)

3. To transcribe speech:

   - Place your cursor where you want the transcribed text to appear
   - Hold the left Alt key
   - Speak clearly
   - Release the Alt key
   - The transcribed text will be typed automatically

4. To exit:
   - Press ESC, or
   - Right-click the system tray icon and select "Exit"

## Auto-start on Windows Login

Create a shortcut to the script in your Windows Startup folder:

```
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
```

## Logs

The application generates the following local log files (not included in the repository):

- Transcriptions: `transcripts.log` - Records all transcribed speech with timestamps
- Application output: `output.log` - Contains general application status messages
- Errors: `error.log` - Records any errors for troubleshooting

All logs are stored in the same directory as the script.

## License

MIT License
