import argparse
import io
import json
import re
import sys
import time
import traceback
from pathlib import Path
from typing import Iterable, List

import fitz  # PyMuPDF
import numpy as np
from deep_translator import GoogleTranslator
from docx import Document
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer
from xml.sax.saxutils import escape


SECTIONS = [
    ("Syllabus", "Tally Syllabus.pdf", "pdf"),
    ("Notes", "Tally Notes.pdf", "pdf"),
    ("Tally Workbook Case Study - 1", "Tally Workbook Case Study - 1.pdf", "pdf"),
    ("Tally Workbook Case Study - 2", "Tally Workbook Case Study - 2.pdf", "pdf"),
    ("Tally Workbook Case Study - 3", "Tally Workbook Case Study - 3.pdf", "pdf"),
    ("Tally Workbook Case Study - 4", "Tally Workbook Case Study - 4.pdf", "pdf"),
    ("Tally Workbook Case Study - 5", "Tally Workbook Case Study - 5.pdf", "pdf"),
    ("Tally Workbook Case Study - 6", "Tally Workbook Case Study - 6.pdf", "pdf"),
    ("Tally.ERP 9 Keyboard Shortcuts", "Tally.ERP  9 Keyboard Shortcuts.pdf", "pdf"),
    ("Configuration", "configuration.docx", "docx"),
    ("TDS", "tds.docx", "docx"),
]


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = text.replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def has_devanagari(text: str) -> bool:
    return bool(re.search(r"[\u0900-\u097F]", text))


def preprocess_image_for_ocr(pil_img: Image.Image) -> Image.Image:
    gray = pil_img.convert("L")
    gray = ImageOps.autocontrast(gray)
    gray = ImageEnhance.Sharpness(gray).enhance(1.8)
    gray = gray.filter(ImageFilter.MedianFilter(size=3))
    return gray


def ocr_page(page: fitz.Page, reader) -> str:
    pix = page.get_pixmap(dpi=300, alpha=False)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    img = preprocess_image_for_ocr(img)
    arr = np.array(img)
    lines = reader.readtext(arr, detail=0, paragraph=True)
    return clean_text("\n".join(lines))


def extract_pdf_text(pdf_path: Path, get_reader) -> str:
    doc = fitz.open(pdf_path)
    page_texts: List[str] = []
    total_pages = len(doc)
    reader = None

    for idx, page in enumerate(doc):
        print(f"  - Reading page {idx + 1}/{total_pages}", flush=True)
        extracted = clean_text(page.get_text("text"))
        # OCR only when text extraction clearly failed; this avoids heavy OCR on already-readable pages.
        use_ocr = (not extracted) or (len(extracted) < 60)

        if use_ocr:
            if reader is None:
                reader = get_reader()
            if reader is not None:
                ocr_txt = ocr_page(page, reader)
                best = ocr_txt if len(ocr_txt) > len(extracted) else extracted
                page_text = best
            else:
                page_text = extracted
        else:
            page_text = extracted

        if not page_text:
            page_text = f"[Unclear text on page {idx + 1}]"
        page_texts.append(page_text)

    doc.close()
    return clean_text("\n\n".join(page_texts))


def extract_docx_text(docx_path: Path) -> str:
    doc = Document(docx_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
    return clean_text("\n".join(lines))


def chunks_for_translation(text: str, max_chars: int = 4500) -> Iterable[str]:
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


def translate_text(text: str, translator: GoogleTranslator, pause: float = 0.05) -> str:
    if not text.strip():
        return ""
    out: List[str] = []
    chunk_list = list(chunks_for_translation(text))
    total = len(chunk_list)
    for i, chunk in enumerate(chunk_list, start=1):
        print(f"  - Translating chunk {i}/{total}", flush=True)
        translated = None
        for attempt in range(3):
            try:
                translated = translator.translate(chunk)
                break
            except Exception:
                time.sleep(1.5 * (attempt + 1))
        if translated is None:
            translated = "[Translation failed for one chunk. Original text retained below]\n" + chunk
        out.append(translated)
        if pause > 0:
            time.sleep(pause)
    return clean_text("\n\n".join(out))


def add_section(story: list, title: str, content: str, heading_style, body_style) -> None:
    story.append(PageBreak())
    story.append(Paragraph(escape(title), heading_style))
    story.append(Spacer(1, 0.4 * cm))

    paragraphs = [p.strip() for p in re.split(r"\n\n+", content) if p.strip()]
    if not paragraphs:
        story.append(Paragraph("No content found.", body_style))
        return

    for para in paragraphs:
        para = escape(para).replace("\n", "<br/>")
        story.append(Paragraph(para, body_style))
        story.append(Spacer(1, 0.18 * cm))


def build_pdf(output_pdf: Path, translated_sections: List[tuple]) -> None:
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    heading_style = styles["Heading1"]
    body_style = ParagraphStyle(
        "Body",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=14,
    )

    doc = SimpleDocTemplate(
        str(output_pdf),
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=1.6 * cm,
        bottomMargin=1.6 * cm,
    )

    story = []
    story.append(Paragraph("Tally Study Book (English Translation)", title_style))
    story.append(Spacer(1, 0.6 * cm))
    story.append(Paragraph("Compiled from source PDFs and DOCX files in the requested order.", body_style))

    for title, content in translated_sections:
        add_section(story, title, content, heading_style, body_style)

    doc.build(story)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build merged English Tally PDF book from PDFs and DOCX files.")
    parser.add_argument(
        "--base-dir",
        default=str(Path(__file__).resolve().parent),
        help="Folder that contains the source files.",
    )
    parser.add_argument(
        "--output",
        default="Tally_English_Merged_Book.pdf",
        help="Output PDF filename.",
    )
    parser.add_argument(
        "--debug-text",
        default="Tally_English_Merged_Book.txt",
        help="Debug output text file filename.",
    )
    parser.add_argument(
        "--checkpoint",
        default="tally_translation_checkpoint.json",
        help="Checkpoint JSON filename for resume support.",
    )
    parser.add_argument(
        "--mode",
        choices=["next", "all", "build"],
        default="next",
        help="next=process only first pending section, all=process all pending, build=only build PDF from checkpoint.",
    )
    parser.add_argument(
        "--enable-ocr",
        action="store_true",
        default=False,
        help="Enable OCR for scanned pages (disabled by default for stability).",
    )
    args = parser.parse_args()

    base_dir = Path(args.base_dir)
    output_pdf = base_dir / args.output
    debug_txt = base_dir / args.debug_text
    checkpoint_path = base_dir / args.checkpoint

    missing = [name for _, name, _ in SECTIONS if not (base_dir / name).exists()]
    if missing:
        print("Missing required files:")
        for item in missing:
            print(f" - {item}")
        return 1

    translator = GoogleTranslator(source="auto", target="en")
    reader_box = {"obj": None, "attempted": False, "unavailable": False}

    def get_reader():
        if not args.enable_ocr:
            return None
        if reader_box["unavailable"]:
            return None
        if reader_box["obj"] is None and not reader_box["attempted"]:
            reader_box["attempted"] = True
            # Lazy import because EasyOCR model download can take time on first run.
            try:
                import easyocr

                reader_box["obj"] = easyocr.Reader(["hi", "en"], gpu=False)
            except Exception as exc:
                reader_box["unavailable"] = True
                print(f"Warning: OCR unavailable; continuing without OCR. Reason: {exc}")
        return reader_box["obj"]

    cached = {}
    if checkpoint_path.exists():
        try:
            cached = json.loads(checkpoint_path.read_text(encoding="utf-8"))
            print(f"Loaded checkpoint: {checkpoint_path}")
        except Exception:
            cached = {}

    pending_sections = []
    for title, filename, ftype in SECTIONS:
        if title not in cached or not str(cached[title]).strip():
            pending_sections.append((title, filename, ftype))

    if args.mode == "build":
        if pending_sections:
            print(f"Cannot build yet. Pending sections: {len(pending_sections)}")
            for title, _, _ in pending_sections:
                print(f" - {title}")
            return 2
    elif args.mode == "next":
        if not pending_sections:
            print("All sections already processed in checkpoint.")
        else:
            pending_sections = [pending_sections[0]]
    # mode == "all" keeps all pending sections.

    for title, filename, ftype in pending_sections:
        source_path = base_dir / filename
        print(f"Processing: {filename}")
        try:
            if ftype == "pdf":
                raw = extract_pdf_text(source_path, get_reader)
            else:
                raw = extract_docx_text(source_path)

            if not raw:
                raw = "[No text content detected]"

            translated = translate_text(raw, translator)
        except Exception:
            translated = "[Section processing failed; partial fallback follows]\n\n" + traceback.format_exc()

        cached[title] = translated
        checkpoint_path.write_text(
            json.dumps(cached, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        done_count = sum(1 for section_title, _, _ in SECTIONS if str(cached.get(section_title, "")).strip())
        print(f"Checkpoint saved: {done_count}/{len(SECTIONS)} sections completed.")

    still_pending = [title for title, _, _ in SECTIONS if not str(cached.get(title, "")).strip()]
    if still_pending and args.mode != "all":
        print("Run again to process next section.")
        for title in still_pending:
            print(f" - Pending: {title}")
        return 0

    if still_pending and args.mode == "all":
        print("Processing stopped with pending sections:")
        for title in still_pending:
            print(f" - {title}")
        return 3

    translated_sections: List[tuple] = []
    debug_parts: List[str] = []
    for title, _, _ in SECTIONS:
        translated = str(cached.get(title, "")).strip()
        if not translated:
            translated = "[No content captured for this section]"
        translated_sections.append((title, translated))

        debug_parts.append(f"\n\n===== {title} =====\n")
        debug_parts.append(translated)

    build_pdf(output_pdf, translated_sections)
    debug_txt.write_text("\n".join(debug_parts), encoding="utf-8")

    print(f"Created: {output_pdf}")
    print(f"Created: {debug_txt}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
