import io
import json
import re
import time
from pathlib import Path

import fitz
import numpy as np
from deep_translator import GoogleTranslator
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

BASE_DIR = Path(__file__).resolve().parent
PDF_PATH = BASE_DIR / "Tally Notes.pdf"
CHECKPOINT_PATH = BASE_DIR / "tally_translation_checkpoint.json"
OUTPUT_TXT = BASE_DIR / "Notes_English_Refreshed.txt"


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def looks_garbled_legacy_hindi(text: str) -> bool:
    if not text:
        return True
    patterns = [r"\[kk", r"\bog\b", r"\bvk\b", r"dgyk", r"lEcf", r"O;fDr"]
    pattern_hits = sum(1 for p in patterns if re.search(p, text))
    symbol_ratio = sum(1 for ch in text if ch in "[];`~") / max(len(text), 1)
    return pattern_hits >= 2 or symbol_ratio > 0.015


def preprocess_image_for_ocr(pil_img: Image.Image) -> Image.Image:
    gray = pil_img.convert("L")
    gray = ImageOps.autocontrast(gray)
    gray = ImageEnhance.Contrast(gray).enhance(1.7)
    gray = ImageEnhance.Sharpness(gray).enhance(1.8)
    gray = gray.filter(ImageFilter.MedianFilter(size=3))
    return gray


def ocr_page(page, reader) -> str:
    pix = page.get_pixmap(dpi=320, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    img = preprocess_image_for_ocr(img)
    arr = np.array(img)
    lines = reader.readtext(arr, detail=0, paragraph=True)
    return clean_text("\n".join(lines))


def extract_notes_text() -> str:
    if not PDF_PATH.exists():
        raise FileNotFoundError(f"Missing file: {PDF_PATH}")

    doc = fitz.open(PDF_PATH)
    total = len(doc)
    print(f"Total pages: {total}")

    reader = None
    page_texts = []

    for i, page in enumerate(doc, start=1):
        native = clean_text(page.get_text("text"))
        use_ocr = (not native) or (len(native) < 120) or looks_garbled_legacy_hindi(native)

        if use_ocr:
            if reader is None:
                import easyocr

                print("Initializing OCR model...")
                reader = easyocr.Reader(["hi", "en"], gpu=False)
            print(f"Page {i}/{total}: OCR")
            try:
                ocr_txt = ocr_page(page, reader)
            except Exception:
                ocr_txt = ""
            best = ocr_txt if len(ocr_txt) > len(native) else native
            page_texts.append(best if best else f"[Unclear text on page {i}]")
        else:
            print(f"Page {i}/{total}: native text")
            page_texts.append(native)

    doc.close()
    return clean_text("\n\n".join(page_texts))


def chunks_for_translation(text: str, max_chars: int = 4500):
    paras = [p.strip() for p in text.split("\n\n") if p.strip()]
    chunk = ""
    for para in paras:
        candidate = para if not chunk else f"{chunk}\n\n{para}"
        if len(candidate) <= max_chars:
            chunk = candidate
        else:
            if chunk:
                yield chunk
            if len(para) <= max_chars:
                chunk = para
            else:
                start = 0
                while start < len(para):
                    end = min(start + max_chars, len(para))
                    yield para[start:end]
                    start = end
                chunk = ""
    if chunk:
        yield chunk


def translate_to_english(text: str) -> str:
    translator = GoogleTranslator(source="auto", target="en")
    parts = []
    chunk_list = list(chunks_for_translation(text))
    total = len(chunk_list)
    for idx, chunk in enumerate(chunk_list, start=1):
        print(f"Translating chunk {idx}/{total}")
        translated = None
        for attempt in range(3):
            try:
                translated = translator.translate(chunk)
                break
            except Exception:
                time.sleep(1.2 * (attempt + 1))
        if translated is None:
            translated = "[Translation failed for one chunk. Original text retained below]\n" + chunk
        parts.append(translated)
        time.sleep(0.05)
    return clean_text("\n\n".join(parts))


def update_checkpoint(notes_english: str):
    if CHECKPOINT_PATH.exists():
        data = json.loads(CHECKPOINT_PATH.read_text(encoding="utf-8"))
    else:
        data = {}
    data["Notes"] = notes_english
    CHECKPOINT_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    notes_raw = extract_notes_text()
    notes_english = translate_to_english(notes_raw)
    OUTPUT_TXT.write_text(notes_english, encoding="utf-8")
    update_checkpoint(notes_english)
    print(f"Updated checkpoint: {CHECKPOINT_PATH}")
    print(f"Saved notes output: {OUTPUT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
