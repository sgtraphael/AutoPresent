# AutoPresent

PowerPoint Auto Presenter with Text-to-Speech.

AutoPresent reads the speaker notes of each slide aloud and can automatically advance to the next slide. It supports:

- Built-in Windows **SAPI** voices
- Offline neural voices via **Piper**
- Cloud neural voices via **Azure Speech**

## Features

- Automatically reads PowerPoint speaker notes
- Auto-advance slides after notes are finished
- Optional **manual advance mode** (disable auto-advance for Q&A)
- Pause / Resume / Stop controls
- Start from any slide
- Jump to a specific slide
- Optional always-on-top subtitle window
- Sentence-by-sentence subtitle updates
- Adjustable speech rate and volume
- Works as a standalone `.exe` (no Python required for end users)

## TTS Engines

- **SAPI (Windows)** – built-in local voices
- **Piper (Neural)** – offline higher-quality voices
- **Azure Speech** – cloud neural voices (English, Chinese, Filipino/Tagalog, and many more)

### Azure notes

- Azure requires internet access and a valid Speech key/region
- Language filter is shown only when Azure is selected
- Defaults to English + Jenny when available
- Pause / Next / Prev / Jump may wait until the current short chunk finishes  
  (cloud TTS limitation; SAPI and Piper can stop more immediately)

## Requirements

- Windows 10/11
- Microsoft PowerPoint installed
- Python 3.10+ (only needed if running from source)
- For Azure: an Azure Speech resource key + region

## Installation (from source)

```bash
git clone https://github.com/sgtraphael/AutoPresent.git
cd AutoPresent
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### Optional Azure setup

Create a `.env` file in the project root:

```env
AZURE_SPEECH_KEY=your_azure_key
AZURE_SPEECH_REGION=eastasia
```

## Run from source

```bash
python app.py
```

## Using the packaged app

1. Download the latest release (`AutoPresent.exe.zip`)
2. Extract it
3. (Optional) Add Piper voice model files for offline neural voices
4. (Optional) Create a `.env` file next to `AutoPresent.exe` for Azure:

```env
AZURE_SPEECH_KEY=your_azure_key
AZURE_SPEECH_REGION=eastasia
```

5. Run `AutoPresent.exe`

## Piper voices (optional)

If you want offline neural voices:

1. Download the Piper voice models
2. Place the `.onnx` and `.onnx.json` files in the same folder as `AutoPresent.exe`,  
   or in `%LOCALAPPDATA%\piper\voices\`

## Usage

1. Open a PowerPoint `.pptx` file
2. Choose a TTS engine
3. (Azure only) Choose language and voice
4. Optionally enable **Show Subtitles**
5. Optionally disable **Auto-advance to next slide** if you want manual control
6. Click **Start**

## Controls

- **Start** – begin presentation
- **Pause / Resume** – pause or continue speaking
- **Stop** – stop presentation
- **Prev / Next** – move between slides
- **Jump** – go to a specific slide number

## Notes

- Azure key is loaded from `.env` and should not be hardcoded
- For company use, keep the real Azure key outside the public repository
- SAPI/Piper are best when instant pause/skip is important
- Azure is best when natural cloud voice quality is more important