"""
🧠 Darija Translator Backend - Hugging Face Spaces
====================================================
This runs on Hugging Face Spaces (FREE) and provides:
- Whisper transcription (Arabic/Darija speech → text)
- CORS-enabled API for the GitHub Pages frontend

Deploy: https://huggingface.co/spaces → New Space → Docker → Upload these files
"""

import os
import time
import tempfile
from flask import Flask, request, jsonify
from flask_cors import CORS
from faster_whisper import WhisperModel

app = Flask(__name__)
CORS(app)  # Allow cross-origin requests from GitHub Pages

# Load Whisper model
print("Loading Whisper model...", flush=True)
whisper_model = WhisperModel("base", device="cpu", compute_type="int8")
print("✅ Whisper ready!", flush=True)


@app.route('/')
def index():
    return jsonify({
        "service": "Darija Translator Backend",
        "status": "running",
        "endpoints": ["/transcribe"]
    })


@app.route('/transcribe', methods=['POST'])
def transcribe():
    """Transcribe audio using Whisper AI."""
    if 'audio' not in request.files:
        return jsonify({'error': 'No audio file'}), 400
    
    audio_file = request.files['audio']
    language = request.form.get('language', 'ar')
    
    temp_path = os.path.join(tempfile.gettempdir(), f'whisper_{int(time.time())}.webm')
    
    try:
        audio_file.save(temp_path)
        
        if os.path.getsize(temp_path) < 1000:
            return jsonify({'text': '', 'language': language})
        
        segments, info = whisper_model.transcribe(
            temp_path,
            language=language,
            beam_size=5,
            best_of=5,
            temperature=0.0,
            vad_filter=True,
            vad_parameters=dict(min_silence_duration_ms=500, speech_pad_ms=300)
        )
        
        full_text = " ".join(seg.text for seg in segments).strip()
        
        return jsonify({
            'text': full_text,
            'language': info.language,
            'confidence': info.language_probability
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
    finally:
        try:
            os.remove(temp_path)
        except:
            pass


@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=7860)
