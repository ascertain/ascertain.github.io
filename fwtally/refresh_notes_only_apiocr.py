import io
import json
import re
import time
from pathlib import Path

import fitz
import requests
from deep_translator import GoogleTranslator

BASE_DIR = Path(__file__).resolve().parent
PDF_PATH = BASE_DIR / "Tally Notes.pdf"
CHECKPOINT_PATH = BASE_DIR / "tally_translation_checkpoint.json"
OUT_TXT = BASE_DIR / "Notes_English_Refreshed.txt"


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def looks_garbled(text: str) -> bool:
    if not text:
        return True
    patterns = [r"\[kk", r"\bog\b", r"dgyk", r"lEcf", r"O;fDr", r"vk;"]
    hits = sum(1 for p in patterns if re.search(p, text))
    return hits >= 2


def ocr_space_image_bytes(image_bytes: bytes) -> str:
    url = "https://api.ocr.space/parse/image"
    payload = {
        "apikey": "helloworld",
        "language": "hin",
        "isOverlayRequired": False,
        "OCREngine": 2,
        "detectOrientation": True,
        "scale": True,
    }
    files = {"filename": ("page.png", image_bytes, "image/png")}
    r = requests.post(url, data=payload, files=files, timeout=120)
    r.raise_for_status()
    data = r.json()
    if data.get("IsErroredOnProcessing"):
        msg = data.get("ErrorMessage") or data.get("ErrorDetails") or "OCR processing error"
        if isinstance(msg, list):
            msg = " | ".join(str(x) for x in msg)
        raise RuntimeError(str(msg))
    parsed = data.get("ParsedResults") or []
    out = []
    for item in parsed:
        out.append(item.get("ParsedText", ""))
    return clean_text("\n".join(out))


def extract_notes_text() -> str:
    doc = fitz.open(PDF_PATH)
    page_texts = []
    total = len(doc)

    for i, page in enumerate(doc, start=1):
        native = clean_text(page.get_text("text"))
        need_ocr = (len(native) < 120) or looks_garbled(native)

        if need_ocr:
            print(f"Page {i}/{total}: OCR API", flush=True)
            pix = page.get_pixmap(dpi=300, alpha=False)
            png_bytes = pix.tobytes("png")
            ocr_txt = ""
            for attempt in range(3):
                try:
                    ocr_txt = ocr_space_image_bytes(png_bytes)
                    break
                except Exception:
                    time.sleep(1.8 * (attempt + 1))
            best = ocr_txt if len(ocr_txt) > len(native) else native
            page_texts.append(best if best else f"[Unclear text on page {i}]")
            time.sleep(1.0)
        else:
            print(f"Page {i}/{total}: native text", flush=True)
            page_texts.append(native)

    doc.close()
    return clean_text("\n\n".join(page_texts))


def chunks_for_translation(text: str, max_chars: int = 3500):
    paras = [p.strip() for p in text.split("\n\n") if p.strip()]
    chunk = ""
    for para in paras:
        candidate = para if not chunk else f"{chunk}\n\n{para}"
        if len(candidate) <= max_chars:
            chunk = candidate
        else:
            if chunk:
                yield chunk
            chunk = para
    if chunk:
        yield chunk


def translate_to_english(text: str) -> str:
    translator = GoogleTranslator(source="auto", target="en")
    chunks = list(chunks_for_translation(text))
    out = []
    for i, c in enumerate(chunks, start=1):
        print(f"Translating chunk {i}/{len(chunks)}", flush=True)
        translated = None
        for attempt in range(3):
            try:
                translated = translator.translate(c)
                break
            except Exception:
                time.sleep(1.2 * (attempt + 1))
        if translated is None:
            translated = "[Translation failed for one chunk. Original text retained below]\n" + c
        out.append(translated)
        time.sleep(0.1)
    return clean_text("\n\n".join(out))


def update_checkpoint(notes_en: str):
    data = {}
    if CHECKPOINT_PATH.exists():
        data = json.loads(CHECKPOINT_PATH.read_text(encoding="utf-8"))
    data["Notes"] = notes_en
    CHECKPOINT_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    raw = extract_notes_text()
    notes_en = translate_to_english(raw)
    OUT_TXT.write_text(notes_en, encoding="utf-8")
    update_checkpoint(notes_en)
    print(f"Updated: {CHECKPOINT_PATH}")
    print(f"Saved: {OUT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
