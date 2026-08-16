import argparse
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


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def apply_manual_legacy_fixes(text: str) -> str:
    # Targeted fixes for common legacy-encoded Hindi lines found in Tally Notes.
    fixes = [
        (
            r"og \[kkrs tks fdfl O;fDr ;k laLFkk ds uke ls cuk, tkrs gS\] os Personal\s*Account ¼O;fDrxr \[kkrs½ dgykrs gSA tSls&",
            "Those accounts that are created in the name of a person or an institution are called Personal Accounts (individual accounts), such as:",
        ),
        (
            r"og \[kkrs tks fdfl oLrq ;k lEifRr vkfn ls lEcfU\/kr gksrs gSa Real\s*Account ¼okLrfod \[kkrs½ dgykrs gSaA tSls&",
            "Those accounts related to assets, property, or things are called Real Accounts, such as:",
        ),
        (
            r"og \[kkrs tks ykHk&gkfu\] vk;&O;;\] vkSj Ø;&foØ; ls lEcfU\/kr gksrs gSa\s*Nominal Account ¼ukeek= ds \[kkrs½ dgykrs gSaA tSls&",
            "Those accounts related to profit-loss, income-expense, and purchase-sales are called Nominal Accounts, such as:",
        ),
        (
            r"¼ ikus okyk ½",
            "(the receiver)",
        ),
        (
            r"¼ nsus okyk ½",
            "(the giver)",
        ),
        (
            r"¼ vkus okyk ½",
            "(what comes in)",
        ),
        (
            r"¼ tkus okyk ½",
            "(what goes out)",
        ),
        (
            r"¼ gkfu\] \[kPkkZ ½",
            "(loss and expenses)",
        ),
        (
            r"¼ ykHk\] vk; ½",
            "(profit and income)",
        ),
        (
            r"lHkh izdkj ds iw¡th \[kkrs ¼ tc O;kikj vkjEHk djrs gS ½\s*Capital Account ds vUrxZr \[kksys tkrs gSA tSls&",
            "All capital-related accounts opened at the start of business are grouped under Capital Account, such as:",
        ),
        (
            r"O;kikj \}kjk fy, x, _\.k O;kikj dk nkf;Ro gSA bl izdkj\s*ds \[kkrs Loans Liabilities ds vUrxZr cuk, tkrs gSA tSls&",
            "Loan amounts taken by the business are liabilities, and such accounts are grouped under Loans Liabilities, such as:",
        ),
        (
            r"Pkkyw nkf;Ro ds \[kkrs Current Liabilities ds vUrxZr \[kksys\s*tkrs gSA tSls&vfxze fy;k gqvk _\.k\] fofo/k ysunkjA",
            "Current obligation accounts are opened under Current Liabilities, such as short-term borrowings and sundry creditors.",
        ),
        (
            r"jke us jkgqy ls 5000 dk Computer m/kkj \[kjhnk A",
            "Ram purchased a computer worth 5000 on credit from Rahul.",
        ),
        (
            r"jke us vius edku dks fxjoh j\[k dj 10000 dk jkgqy ls yksu fy;kA",
            "Ram mortgaged his house and took a loan of 10000 from Rahul.",
        ),
    ]

    out = text
    for pattern, replacement in fixes:
        out = re.sub(pattern, replacement, out, flags=re.IGNORECASE)
    return out


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


def extract_page_text(page: fitz.Page, force_ocr: bool = False) -> tuple[str, str]:
    native = clean_text(page.get_text("text"))
    need_ocr = force_ocr or (len(native) < 120) or looks_garbled(native)

    if need_ocr:
        pix = page.get_pixmap(dpi=300, alpha=False)
        png_bytes = pix.tobytes("png")
        ocr_txt = ""
        for attempt in range(3):
            try:
                ocr_txt = ocr_space_image_bytes(png_bytes)
                break
            except Exception:
                time.sleep(1.8 * (attempt + 1))
        if force_ocr and ocr_txt:
            best = ocr_txt
            source = "ocr"
        else:
            best = ocr_txt if len(ocr_txt) > len(native) else native
            source = "ocr" if len(ocr_txt) >= len(native) else "native"
        return (best if best else "[Unclear page text]", source)

    return native, "native"


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
    text = apply_manual_legacy_fixes(text)
    translator = GoogleTranslator(source="auto", target="en")
    chunks = list(chunks_for_translation(text))
    out = []
    for c in chunks:
        translated = None
        for attempt in range(3):
            try:
                translated = translator.translate(c)
                break
            except Exception:
                time.sleep(1.2 * (attempt + 1))
        if translated is None:
            translated = "[Translation failed for chunk]\n" + c
        out.append(translated)
        time.sleep(0.1)
    return clean_text("\n\n".join(out))


def parse_pages_arg(pages: str, total_pages: int) -> list[int]:
    result = []
    for part in pages.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start = int(a)
            end = int(b)
            for p in range(start, end + 1):
                if 1 <= p <= total_pages:
                    result.append(p)
        else:
            p = int(part)
            if 1 <= p <= total_pages:
                result.append(p)
    result = sorted(set(result))
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Process selected pages from Tally Notes for review.")
    parser.add_argument("--pages", default="1-3", help="Pages to process, e.g. 1 or 1-3 or 1,3,5-7")
    parser.add_argument(
        "--force-ocr-pages",
        default="",
        help="Pages to force OCR, e.g. 2 or 2-4 or 2,4,8",
    )
    parser.add_argument("--update-checkpoint", action="store_true", help="Update Notes in checkpoint with processed pages only")
    args = parser.parse_args()

    doc = fitz.open(PDF_PATH)
    selected = parse_pages_arg(args.pages, len(doc))
    forced = parse_pages_arg(args.force_ocr_pages, len(doc)) if args.force_ocr_pages.strip() else []
    forced_set = set(forced)
    if not selected:
        print("No valid pages selected")
        return 1

    review_parts = []
    review_meta = []

    for page_no in selected:
        page = doc[page_no - 1]
        print(f"Processing page {page_no}", flush=True)
        raw_text, source = extract_page_text(page, force_ocr=page_no in forced_set)
        en_text = translate_to_english(raw_text)

        review_meta.append({"page": page_no, "source": source, "raw_len": len(raw_text), "en_len": len(en_text)})
        review_parts.append(f"\n\n===== PAGE {page_no} ({source}) =====\n")
        review_parts.append(en_text)

    doc.close()

    out_review = BASE_DIR / f"Notes_review_pages_{args.pages.replace(',', '_').replace('-', 'to')}.txt"
    out_json = BASE_DIR / f"Notes_review_pages_{args.pages.replace(',', '_').replace('-', 'to')}.json"
    out_review.write_text("\n".join(review_parts), encoding="utf-8")
    out_json.write_text(json.dumps(review_meta, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.update_checkpoint:
        data = {}
        if CHECKPOINT_PATH.exists():
            data = json.loads(CHECKPOINT_PATH.read_text(encoding="utf-8"))
        existing = str(data.get("Notes", "")).strip()
        merged = existing + "\n\n" + "\n".join(review_parts)
        data["Notes"] = clean_text(merged)
        CHECKPOINT_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Created review text: {out_review}")
    print(f"Created review meta: {out_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
