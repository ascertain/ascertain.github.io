import json
import re
import time
from pathlib import Path
from typing import Iterable, List

from deep_translator import GoogleTranslator
from docx import Document
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer
from xml.sax.saxutils import escape

BASE_DIR = Path(__file__).resolve().parent
CHECKPOINT = BASE_DIR / "tally_translation_checkpoint.json"
OUTPUT_PDF = BASE_DIR / "Tally_English_Merged_Book.pdf"
OUTPUT_TXT = BASE_DIR / "Tally_English_Merged_Book.txt"

SECTIONS = [
    "Syllabus",
    "Notes",
    "Tally Workbook Case Study - 1",
    "Tally Workbook Case Study - 2",
    "Tally Workbook Case Study - 3",
    "Tally Workbook Case Study - 4",
    "Tally Workbook Case Study - 5",
    "Tally Workbook Case Study - 6",
    "Tally.ERP 9 Keyboard Shortcuts",
    "Configuration",
    "TDS",
]


def clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


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


def translate_text(text: str) -> str:
    translator = GoogleTranslator(source="auto", target="en")
    out: List[str] = []
    chunk_list = list(chunks_for_translation(text))
    total = len(chunk_list)
    for i, chunk in enumerate(chunk_list, start=1):
        print(f"Translating TDS chunk {i}/{total}", flush=True)
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


def build_pdf(sections: List[tuple]) -> None:
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
        str(OUTPUT_PDF),
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=1.6 * cm,
        bottomMargin=1.6 * cm,
    )

    story = [
        Paragraph("Tally Study Book (English Translation)", title_style),
        Spacer(1, 0.6 * cm),
        Paragraph("Compiled from source PDFs and DOCX files in the requested order.", body_style),
    ]

    for title, content in sections:
        add_section(story, title, content, heading_style, body_style)

    doc.build(story)


def main() -> int:
    if not CHECKPOINT.exists():
        print(f"Missing checkpoint: {CHECKPOINT}")
        return 1

    data = json.loads(CHECKPOINT.read_text(encoding="utf-8"))

    if not str(data.get("TDS", "")).strip():
        tds_path = BASE_DIR / "tds.docx"
        if not tds_path.exists():
            print(f"Missing TDS source: {tds_path}")
            return 1
        raw_tds = extract_docx_text(tds_path)
        data["TDS"] = translate_text(raw_tds) if raw_tds else "[No text content detected]"
        CHECKPOINT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print("Saved TDS into checkpoint.")

    missing = [name for name in SECTIONS if not str(data.get(name, "")).strip()]
    if missing:
        print("Cannot build. Missing sections:")
        for name in missing:
            print(f" - {name}")
        return 2

    ordered_sections = [(name, str(data.get(name, "")).strip()) for name in SECTIONS]
    build_pdf(ordered_sections)

    debug_parts = []
    for title, content in ordered_sections:
        debug_parts.append(f"\n\n===== {title} =====\n")
        debug_parts.append(content)
    OUTPUT_TXT.write_text("\n".join(debug_parts), encoding="utf-8")

    print(f"Created: {OUTPUT_PDF}")
    print(f"Created: {OUTPUT_TXT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
