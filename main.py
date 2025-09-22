import comtypes.client
from PyPDF2 import PdfReader

# ---------- Configuration ----------
PDF_PATH = "input.pdf"  # Path to your PDF
OUTPUT_WAV = "output.wav"  # Output audio file
VOICE_NAME = "David"  # Choose from "David", "Zira", and "Hazel"


# ---------- Step 1: Extract text from PDF ----------
def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text().replace("\n", " ") + "\n"
    return text


# ---------- Step 2: Set up Speech Synthesizer ----------
def speak_text_to_file(text, output_path, voice_name):
    comtypes.CoInitialize()
    speaker = comtypes.client.CreateObject("SAPI.SpVoice")

    for v in speaker.GetVoices():
        if voice_name in v.GetDescription():
            speaker.Voice = v
            break

    stream = comtypes.client.CreateObject("SAPI.SpFileStream")
    stream.Open(output_path, 3, False)  # SSFMCreateForWrite = 3
    speaker.AudioOutputStream = stream
    speaker.Speak(text)
    stream.Close()
    comtypes.CoUninitialize()


if __name__ == "__main__":
    text = extract_text_from_pdf(PDF_PATH)
    print("PDF text extracted, generating audio...")
    speak_text_to_file(text, OUTPUT_WAV, VOICE_NAME)
    print(f"Audio saved to {OUTPUT_WAV}")
