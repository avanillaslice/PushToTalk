import whisper
import pyaudio
import wave
import keyboard
import os
import time
import sys
import traceback
import shutil
import socket
import subprocess
from datetime import datetime
from pathlib import Path
import pystray
from PIL import Image, ImageDraw
from threading import Thread

# For Windows-specific text field detection
if sys.platform == 'win32':
    import ctypes
    
    # Use a simpler approach that's more permissive
    # This will allow typing in most applications
    GetForegroundWindow = ctypes.windll.user32.GetForegroundWindow
    
    # List of window class names that should NOT receive keyboard input
    non_text_classes = [
        'Shell_TrayWnd',      # Taskbar
        'Progman',            # Desktop
        'WorkerW',            # Desktop
        'Button',             # Buttons
        'ToolbarWindow32',    # Toolbars
        '#32770',             # Dialog boxes without text fields
        'CabinetWClass',      # Windows Explorer
        'ExploreWClass'       # Windows Explorer (browse mode)
    ]
    
    def is_text_field_focused():
        """Simple check that allows typing most places except obvious non-text areas"""
        try:
            # Get foreground window
            fg_window = GetForegroundWindow()
            if not fg_window:
                return True  # When in doubt, allow typing
                
            # Get class name of the foreground window
            class_name = ctypes.create_unicode_buffer(256)
            ctypes.windll.user32.GetClassNameW(fg_window, class_name, 256)
            class_str = class_name.value
            
            # Debug: Log the window class name
            print(f"[DEBUG] Current window class: {class_str}")
            
            # If the window class is in our non-text list, don't type
            if class_str in non_text_classes:
                print(f"[DEBUG] Blocked typing in non-text window: {class_str}")
                return False
            
            # Additional check for Explorer windows that might have different classes
            window_text = ctypes.create_unicode_buffer(512)
            ctypes.windll.user32.GetWindowTextW(fg_window, window_text, 512)
            window_title = window_text.value
            
            # If the window title contains "Explorer" or specific folder paths
            if "Explorer" in window_title or "This PC" in window_title or ":\\" in window_title:
                print(f"[DEBUG] Blocked typing in Explorer-like window: {window_title}")
                return False
                
            # Allow typing for all other windows
            return True
        except Exception as e:
            # If there's any error, default to allowing typing
            print(f"[DEBUG] Error in text field detection: {str(e)}")
            return True
else:
    # On non-Windows platforms, always allow typing
    def is_text_field_focused():
        return True

# Hide console windows for subprocess calls (this prevents ffmpeg console window)
# Windows-specific flag to hide console window for subprocesses
if sys.platform == 'win32':
    # Set subprocess creation flags to hide window
    CREATE_NO_WINDOW = 0x08000000
    # Store original Popen for later use with our custom flags
    original_popen = subprocess.Popen
    
    # Create a custom Popen that always uses CREATE_NO_WINDOW
    class SilentPopen(subprocess.Popen):
        def __init__(self, *args, **kwargs):
            if 'startupinfo' not in kwargs:
                kwargs['startupinfo'] = subprocess.STARTUPINFO()
                kwargs['startupinfo'].dwFlags |= subprocess.STARTF_USESHOWWINDOW
                kwargs['startupinfo'].wShowWindow = 0  # SW_HIDE
            if 'creationflags' not in kwargs:
                kwargs['creationflags'] = CREATE_NO_WINDOW
            super().__init__(*args, **kwargs)
    
    # Replace the original Popen with our silent version
    subprocess.Popen = SilentPopen

# Single instance lock
LOCK_SOCKET = None
def ensure_single_instance():
    """Ensure only one instance of the application is running using a socket lock"""
    global LOCK_SOCKET
    LOCK_SOCKET = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    
    try:
        # Try to bind to a port that only one instance can use
        LOCK_SOCKET.bind(('localhost', 47482))
        print("No other instance detected, continuing startup")
        return True
    except socket.error:
        print("Another instance is already running. Exiting.")
        return False

# Check if another instance is running first
if not ensure_single_instance():
    sys.exit(0)

# Check dependencies
try:
    import whisper
    import pyaudio
    import keyboard
    import pystray
    from PIL import Image, ImageDraw
except ImportError as e:
    print(f"Missing dependency: {e}")
    print("Please install required packages: pip install openai-whisper pyaudio keyboard pystray pillow")
    sys.exit(1)

# Setup paths - use script location instead of hardcoded path
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(SCRIPT_DIR, "error.log")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "recording.wav")
TRANSCRIPT_LOG = os.path.join(SCRIPT_DIR, "transcripts.log")
STDOUT_LOG = os.path.join(SCRIPT_DIR, "output.log")
FFMPEG_PATH = os.path.join(SCRIPT_DIR, "ffmpeg.exe")

# Create directory if it doesn't exist
os.makedirs(SCRIPT_DIR, exist_ok=True)

# Redirect stdout and stderr to log files for silent operation
sys.stdout = open(STDOUT_LOG, "a", encoding="utf-8")
sys.stderr = open(LOG_FILE, "a", encoding="utf-8")

# Log startup
print(f"[{datetime.now()}] Application started")

# Check for ffmpeg
ffmpeg_found = False
if os.path.exists(FFMPEG_PATH):
    os.environ["PATH"] += os.pathsep + SCRIPT_DIR
    ffmpeg_found = True
else:
    # Check if ffmpeg is in PATH
    if shutil.which("ffmpeg"):
        ffmpeg_found = True
        
if not ffmpeg_found:
    print("Warning: ffmpeg not found. Whisper may not work correctly.")
    log_error("ffmpeg not found in PATH or script directory")

# Error logging helper
def log_error(err_msg):
    with open(LOG_FILE, "a") as log:
        log.write(f"[{datetime.now()}] {err_msg}\n")

# Audio recording config
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 16000
MIN_FRAMES = 10  # Minimum number of frames for a valid recording

print("Loading Whisper model... (this may take a moment)")

# Create system tray icon
def create_image():
    # Create a simple mic icon
    width = 64
    height = 64
    color1 = (255, 255, 255, 0)
    color2 = (0, 128, 255, 255)
    
    image = Image.new('RGBA', (width, height), color1)
    dc = ImageDraw.Draw(image)
    
    # Draw a simple microphone icon
    dc.rounded_rectangle((20, 10, 44, 40), 10, fill=color2)
    dc.rectangle((27, 40, 37, 50), fill=color2)
    dc.ellipse((22, 50, 42, 60), fill=color2)
    
    return image

def on_exit(icon):
    icon.stop()
    global running
    running = False

# Create the system tray menu and icon
def setup_tray():
    icon = pystray.Icon("push_to_talk", create_image(), "Push-to-Talk (left Alt)")
    icon.menu = pystray.Menu(
        pystray.MenuItem("Exit", on_exit)
    )
    icon.run()
    return icon

# Load model
try:
    model = whisper.load_model("base")
    print("Model loaded successfully!")
except Exception:
    error_msg = "Failed to load Whisper model:\n" + traceback.format_exc()
    log_error(error_msg)
    print(f"Error: {error_msg}")
    sys.exit(1)

def record_audio(filename):
    """Record audio while left Alt is pressed, with proper resource cleanup"""
    frames = []
    
    try:
        audio = pyaudio.PyAudio()
        stream = audio.open(format=FORMAT, channels=CHANNELS,
                            rate=RATE, input=True,
                            frames_per_buffer=CHUNK)
        
        print("[Recording] Release left Alt key to stop recording...")
        
        try:
            while keyboard.is_pressed('alt'):
                data = stream.read(CHUNK)
                frames.append(data)
                
            print("[Processing] Recording complete, transcribing...")
        finally:
            # Ensure resources are cleaned up
            stream.stop_stream()
            stream.close()
            audio.terminate()
        
        # Check if recording has enough content
        if len(frames) < MIN_FRAMES:
            print("Recording too short. Please hold left Alt longer while speaking.")
            return False
            
        with wave.open(filename, 'wb') as wf:
            wf.setnchannels(CHANNELS)
            wf.setsampwidth(audio.get_sample_size(FORMAT))
            wf.setframerate(RATE)
            wf.writeframes(b''.join(frames))
            
        return True
            
    except Exception:
        error_msg = "Recording error:\n" + traceback.format_exc()
        log_error(error_msg)
        print(f"Error: {error_msg}")
        return False

def transcribe(filename):
    """Transcribe audio file and log the results, then type the transcribed text"""
    try:
        # Check if file exists and has content
        if not os.path.exists(filename) or os.path.getsize(filename) < 1000:
            print("Recording file is missing or too small to transcribe.")
            return
        
        # Set Windows-specific flags to hide console windows during transcription
        if sys.platform == 'win32':
            # Redirect stderr for whisper during transcription to prevent console window
            original_stderr = sys.stderr
            sys.stderr = open(os.devnull, 'w')
            
        # Run transcription with verbose=False to minimize output
        result = model.transcribe(filename, verbose=False)
        
        # Restore stderr if we changed it
        if sys.platform == 'win32':
            sys.stderr.close()
            sys.stderr = original_stderr
            
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        text = result["text"].strip()
        
        if not text:
            print("No speech detected in the recording.")
            return
            
        output = f"[{timestamp}] You said: {text}"
        print(output)
        
        with open(TRANSCRIPT_LOG, "a", encoding="utf-8") as f:
            f.write(output + "\n")
        
        # Check if cursor is in a text field before typing
        if is_text_field_focused():
            print(f"Typing out: {text}")
            # Small delay before typing to allow user to position cursor
            time.sleep(0.5)
            keyboard.write(text)
        else:
            print("No text field detected. Text not typed.")
            
    except Exception:
        error_msg = "Transcription error:\n" + traceback.format_exc()
        log_error(error_msg)
        print(f"Error: {error_msg}")

def cleanup():
    """Clean up resources before exiting"""
    try:
        if os.path.exists(OUTPUT_FILE):
            os.remove(OUTPUT_FILE)
        print(f"[{datetime.now()}] Application shutting down...")
    except Exception:
        log_error("Cleanup error:\n" + traceback.format_exc())

# Start the system tray in a separate thread
tray_thread = Thread(target=setup_tray, daemon=True)
tray_thread.start()

# Print startup message to log
print("\nPush-to-Talk Transcription")
print("==========================")
print("Hold left Alt to record audio")
print("Press ESC to exit the program")
print("Transcriptions will be saved to:", TRANSCRIPT_LOG)
print("==========================\n")

# Main loop
running = True
while running:
    try:
        if keyboard.is_pressed('esc'):
            running = False
            cleanup()
            break
            
        if keyboard.is_pressed('alt'):
            if record_audio(OUTPUT_FILE):
                transcribe(OUTPUT_FILE)
            time.sleep(0.5)
        else:
            time.sleep(0.1)
    except Exception:
        error_msg = "Main loop crash:\n" + traceback.format_exc()
        log_error(error_msg)
        print(f"Error: {error_msg}")
        time.sleep(1)
