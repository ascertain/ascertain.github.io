"""
🧠 Darija Translator Backend - Hugging Face Spaces (Gradio SDK + ZeroGPU - FREE)
=================================================================================
Provides Whisper transcription (Arabic/Darija speech → text)
with a REST API endpoint for the GitHub Pages frontend.

Deploy: https://huggingface.co/spaces → New Space → Gradio SDK → Zero GPU → Upload files
"""

import os
import tempfile
import time
import json
import spaces
import gradio as gr
from faster_whisper import WhisperModel

# Model loaded on-demand inside GPU-decorated function
whisper_model = None


def get_model():
    global whisper_model
    if whisper_model is None:
        import torch
        device = "cuda" if torch.cuda.is_available() else "cpu"
        compute = "float16" if device == "cuda" else "int8"
        print(f"Loading Whisper model on {device}...", flush=True)
        whisper_model = WhisperModel("base", device=device, compute_type=compute)
        print("✅ Whisper ready!", flush=True)
    return whisper_model


@spaces.GPU
def transcribe_audio(audio_path, language="ar"):
    """Transcribe audio file using faster-whisper with Zero GPU."""
    if audio_path is None:
        return json.dumps({"text": "", "language": language, "error": "No audio"})

    try:
        if os.path.getsize(audio_path) < 1000:
            return json.dumps({"text": "", "language": language})

        segments, info = get_model().transcribe(
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
