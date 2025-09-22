# PDF to Speech (Windows TTS)

A simple Python script to convert the text from a PDF into a **single audio file** using **Windows built-in voices**.

---

## Features

- Extracts all text from a PDF.
- Converts it to speech using Windows TTS (SAPI).
- Generates **one continuous WAV file**.
- Choose from available desktop voices: `David`, `Zira`, `Hazel`.

---

## Requirements

- Windows 10/11  
- Python 3.x  
- Python packages: `PyPDF2`, `comtypes`  

Install dependencies:

```bash
pip install PyPDF2 comtypes
```

## Usage

1. Place your PDF in the repo folder and update the path in the script:

```bash
PDF_PATH = "input.pdf"
OUTPUT_WAV = "output.wav"
VOICE_NAME = "David"
```

2. Run the script:

```bash
python main.py
```

3. The script will generate `output.wav` with the full spoken text.
