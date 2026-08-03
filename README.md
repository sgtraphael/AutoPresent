# AutoPresent

PowerPoint Auto Presenter with Text-to-Speech.

AutoPresent reads the speaker notes of each slide aloud and automatically advances to the next slide. It supports both the built-in Windows SAPI voices and higher-quality offline neural voices via Piper TTS.

## Features

- Automatically reads PowerPoint speaker notes
- Auto-advances slides after notes are finished
- Pause / Resume / Stop controls
- Start from any slide
- Two TTS engines:
  - **SAPI (Windows)** – built-in voices
  - **Piper (Neural)** – offline neural voices (Lessac High, Amy Medium)
- Adjustable speech rate and volume
- Works as a standalone `.exe` (no Python required for end users)

## Requirements

- Windows 10/11
- Microsoft PowerPoint installed
- Python 3.10+ (only needed if running from source)

## Installation (from source)

```bash
git clone https://github.com/sgtraphael/AutoPresent.git
cd AutoPresent
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt