"""
🧠 Darija Translator Backend - Hugging Face Spaces (Gradio SDK - FREE)
======================================================================
Provides Whisper transcription (Arabic/Darija speech → text)
with a REST API endpoint for the GitHub Pages frontend.

Deploy: https://huggingface.co/spaces → New Space → Gradio SDK → Upload files
"""

import os
import tempfile
import time
import json
import gradio as gr
from faster_whisper import WhisperModel

# Load Whisper model at startup
print("Loading Whisper model...", flush=True)
whisper_model = WhisperModel("base", device="cpu", compute_type="int8")
print("✅ Whisper ready!", flush=True)


def transcribe_audio(audio_path, language="ar"):
    """Transcribe audio file using faster-whisper."""
    if audio_path is None:
        return json.dumps({"text": "", "language": language, "error": "No audio"})

    try:
        if os.path.getsize(audio_path) < 1000:
            return json.dumps({"text": "", "language": language})

        segments, info = whisper_model.transcribe(
            audio_path,
            language=language,
            beam_size=5,
            best_of=5,
            temperature=0.0,
            vad_filter=True,
            vad_parameters=dict(min_silence_duration_ms=500, speech_pad_ms=300)
        )

        full_text = " ".join(seg.text for seg in segments).strip()

        return json.dumps({
            "text": full_text,
            "language": info.language,
            "confidence": round(info.language_probability, 3)
        })

    except Exception as e:
        return json.dumps({"text": "", "error": str(e)})


# Gradio Interface
with gr.Blocks(title="Darija Translator - Whisper Backend") as demo:
    gr.Markdown("## 🧠 Darija Whisper Transcription API")
    gr.Markdown("Upload audio or record from mic to transcribe Moroccan Darija / Arabic.")

    with gr.Row():
        audio_input = gr.Audio(
            label="Audio Input",
            type="filepath",
            sources=["upload", "microphone"]
        )
        language_input = gr.Dropdown(
            choices=["ar", "en", "fr"],
            value="ar",
            label="Language"
        )

    output = gr.Textbox(label="Result (JSON)", lines=4)
    transcribe_btn = gr.Button("Transcribe", variant="primary")
    transcribe_btn.click(
        fn=transcribe_audio,
        inputs=[audio_input, language_input],
        outputs=output,
        api_name="transcribe"
    )

    gr.Markdown("---")
    gr.Markdown("### API Usage from JavaScript")
    gr.Markdown("""
    ```javascript
    // Call from your frontend:
    const response = await fetch('https://subrati-darija-backend.hf.space/api/transcribe', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ data: [audioBlob, 'ar'] })
    });
    const result = await response.json();
    const transcription = JSON.parse(result.data[0]);
    ```
    """)

if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=7860)
