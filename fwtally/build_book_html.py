import html
import json
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CHECKPOINT = BASE_DIR / "tally_translation_checkpoint.json"
OUTPUT_HTML = BASE_DIR / "Tally_English_Book.html"

SECTION_ORDER = [
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


def slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r"[^a-z0-9\s-]", "", text)
    text = re.sub(r"\s+", "-", text)
    return text or "section"


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def apply_manual_legacy_fixes(text: str) -> str:
  fixes = [
    (
      r"og\s*\[kkrs\s*tks\s*fdfl\s*O;fDr\s*;k\s*laLFkk\s*ds\s*uke\s*ls\s*cuk,\s*tkrs\s*gS\]\s*os\s*Personal",
      "Those accounts that are created in the name of a person or an institution are called Personal",
    ),
    (
      r"Account\s*¼O;fDrxr\s*\[kkrs½\s*dgykrs\s*gSA\s*tSls[&:]?",
      "Accounts (individual accounts), such as:",
    ),
    (
      r"og \[kkrs tks fdfl O;fDr ;k laLFkk ds uke ls cuk, tkrs gS\] os Personal\s*Account ¼O;fDrxr \[kkrs½ dgykrs gSA tSls&",
      "Those accounts that are created in the name of a person or an institution are called Personal Accounts (individual accounts), such as:",
    ),
    (
      r"og\s*\[kkrs\s*tks\s*fdfl\s*oLrq\s*;k\s*lEifRr\s*vkfn\s*ls\s*lEcfU\/kr\s*gksrs\s*gSa\s*Real",
      "Those accounts related to assets, property, or things are called Real",
    ),
    (
      r"Account\s*¼okLrfod\s*\[kkrs½\s*dgykrs\s*gSaA\s*tSls[&:]?",
      "Accounts, such as:",
    ),
    (
      r"og \[kkrs tks fdfl oLrq ;k lEifRr vkfn ls lEcfU\/kr gksrs gSa Real\s*Account ¼okLrfod \[kkrs½ dgykrs gSaA tSls&",
      "Those accounts related to assets, property, or things are called Real Accounts, such as:",
    ),
    (
      r"og\s*\[kkrs\s*tks\s*ykHk&gkfu\]\s*vk;&O;;\]\s*vkSj\s*Ø;&foØ;\s*ls\s*lEcfU\/kr\s*gksrs\s*gSa\s*Nominal",
      "Those accounts related to profit-loss, income-expense, and purchase-sales are called Nominal",
    ),
    (
      r"Account\s*¼ukeek=\s*ds\s*\[kkrs½\s*dgykrs\s*gSaA\s*tSls[&:]?",
      "Accounts, such as:",
    ),
    (
      r"og \[kkrs tks ykHk&gkfu\] vk;&O;;\] vkSj Ø;&foØ; ls lEcfU\/kr gksrs gSa\s*Nominal Account ¼ukeek= ds \[kkrs½ dgykrs gSaA tSls&",
      "Those accounts related to profit-loss, income-expense, and purchase-sales are called Nominal Accounts, such as:",
    ),
    (r"¼ ikus okyk ½", "(the receiver)"),
    (r"¼ nsus okyk ½", "(the giver)"),
    (r"¼ vkus okyk ½", "(what comes in)"),
    (r"¼ tkus okyk ½", "(what goes out)"),
    (r"¼ gkfu\] \[kPkkZ ½", "(loss and expenses)"),
    (r"¼ ykHk\] vk; ½", "(profit and income)"),
  ]

  out = text
  for pattern, replacement in fixes:
    out = re.sub(pattern, replacement, out, flags=re.IGNORECASE)

  # Normalize full account-group block (1-28) into clean English definitions.
  heading_defs = {
    "capital account": "All capital-related accounts are recorded under Capital Account. Example: if Ram starts business with 10,000, Ram is treated as Capital A/c.",
    "loans liabilities": "Loans taken by the business are obligations of the business and are recorded under Loans Liabilities. Example: loan borrowed from Rahul against security.",
    "current liabilities": "Short-term obligations are recorded under Current Liabilities, such as outstanding loans and creditors. Example: computer purchased on credit from Rahul.",
    "fixed assets": "Long-term assets are recorded under Fixed Assets, such as building, machinery, plant, and equipment. Example: cell phone purchased for business use.",
    "investment account": "Investment-related accounts are recorded under Investment Account, such as dividend investment, securities, or share investments.",
    "current assets": "Short-term assets are recorded under Current Assets, such as cash account, bank account, debtors account, and advance account.",
    "miscellaneous exp assets": "Deferred or preliminary expenses that are carried initially are recorded under Miscellaneous Exp. Assets.",
    "suspense account": "If trial balance does not match, the temporary balancing difference is posted to Suspense Account.",
    "branch division": "Accounts of branch offices created by the main company are recorded under Branch Division.",
    "sales account": "All sales of goods are recorded under Sales Account, whether cash sales or credit sales.",
    "purchase account": "All purchases of goods are recorded under Purchase Account, whether cash purchases or credit purchases.",
    "direct income": "Income earned directly from business operations is recorded under Direct Income.",
    "direct expenses": "Expenses directly related to business operations are recorded under Direct Expenses.",
    "indirect income": "Indirect incomes are recorded under Indirect Income, such as rent received, discount received, and commission received.",
    "indirect expenses": "Indirect expenses are recorded under Indirect Expenses, such as commission expense, general expense, and rent expense.",
    "reserve & surplus": "Reserves created from profit are recorded under Reserve & Surplus, such as capital reserve and redemption reserve.",
    "bank over draft": "When withdrawal exceeds bank balance, it is treated as Bank Overdraft (effectively a short-term borrowing from bank).",
    "secured loans": "Loans backed by security are recorded under Secured Loans.",
    "unsecured loans": "Loans without formal security or legal collateral are recorded under Unsecured Loans.",
    "duties & tax": "All tax-related accounts are recorded under Duties & Tax.",
    "provision": "All provision accounts are recorded under Provision, such as provision for salary and provision for rent.",
    "sundry creditors": "Parties to whom money is payable in future are recorded under Sundry Creditors.",
    "sundry debtors": "Parties from whom money is receivable in future are recorded under Sundry Debtors.",
    "deposit assets": "Money deposited as security for an asset or agreement is recorded under Deposit Assets.",
    "loans & advance assets": "Advances given and loan-like receivables are recorded under Loans & Advance Assets.",
    "cash in hand": "Cash balance maintained by the business is recorded under Cash in Hand.",
    "stock on hand": "Closing/available inventory is recorded under Stock on Hand.",
    "bank account": "All bank ledgers are recorded under Bank Account, such as ICICI Bank, State Bank, and HDFC Bank.",
    "cons voucher": "Contra Voucher is used for transactions between Cash and Bank, such as cash deposited into bank or cash withdrawn from bank.",
    "contra voucher": "Contra Voucher is used for transactions between Cash and Bank, such as cash deposited into bank or cash withdrawn from bank.",
    "payment voucher": "Payment Voucher is used when payment is made to a person by cash or cheque.",
    "receipt voucher": "Receipt Voucher is used when money is received, such as commission, discount, or cash income.",
    "journal voucher": "Journal Voucher is used for non-cash/adjustment and credit transactions, including provisions and transfer entries.",
    "sales voucher": "Sales Voucher is used to record all sales transactions, whether cash sales or credit sales.",
    "purchase voucher": "Purchase Voucher is used to record all purchase transactions, whether cash purchases or credit purchases.",
    "credit score voucher": "Credit Note Voucher is used to record sales return (goods returned by customer).",
    "credit note voucher": "Credit Note Voucher is used to record sales return (goods returned by customer).",
    "debit note voucher": "Debit Note Voucher is used to record purchase return (goods returned to supplier).",
    "memorandum voucher": "Memorandum Voucher is used for provisional or reminder entries that may be converted later.",
  }

  lines = out.split("\n")
  rebuilt = []

  heading_re = re.compile(r"^\s*(?:[-*•–]+\s*)?(\d{1,2})\s*-\s*([A-Za-z&.\s]+?)\s*:\s*-?\s*$")

  i = 0
  while i < len(lines):
    line = lines[i].strip()

    m = heading_re.match(line)
    if not m:
      rebuilt.append(lines[i])
      i += 1
      continue

    title = re.sub(r"\s+", " ", m.group(2)).strip().lower()
    title = title.replace("suspence", "suspense")
    title_norm = re.sub(r"[^a-z0-9\s&]", " ", title)
    title_norm = re.sub(r"\s+", " ", title_norm).strip()

    key = None
    for candidate in heading_defs.keys():
      if candidate in title or candidate in title_norm:
        key = candidate
        break

    if key is None:
      rebuilt.append(lines[i])
      i += 1
      continue

    # Keep heading, inject clean English definition, and skip legacy continuation lines
    rebuilt.append(lines[i])
    rebuilt.append(heading_defs[key])

    i += 1
    while i < len(lines):
      nxt = lines[i].strip()
      if heading_re.match(nxt):
        break
      # Drop page markers and legacy continuation lines for this heading block.
      i += 1

  out = "\n".join(rebuilt)
  out = out.replace("APRIL २०१6", "APRIL 2016")
  return out


def is_subheading(line: str) -> bool:
    if len(line) > 120:
        return False
    if line.endswith(":"):
        return True
    if re.match(r"^\d+[\.-]\s+", line):
        return True
    if re.match(r"^[A-Z][A-Za-z0-9\s&()/-]{2,60}$", line):
        words = line.split()
        return len(words) <= 9
    return False


def render_line(line: str) -> str:
    line = html.escape(line.strip())
    line = re.sub(r"\s+", " ", line)

    if re.match(r"^(\d+[-\.)]|[-*•])\s+", line):
        return f"<li>{line}</li>"

    if is_subheading(line):
        return f"<h3>{line}</h3>"

    return f"<p>{line}</p>"


def render_section(title: str, content: str, index: int) -> str:
    sec_id = f"chapter-{index}-{slugify(title)}"
    content = apply_manual_legacy_fixes(normalize_text(content))
    lines = [ln.strip() for ln in content.split("\n") if ln.strip()]

    body_parts = []
    list_open = False

    for ln in lines:
        rendered = render_line(ln)
        if rendered.startswith("<li>"):
            if not list_open:
                body_parts.append("<ul>")
                list_open = True
            body_parts.append(rendered)
        else:
            if list_open:
                body_parts.append("</ul>")
                list_open = False
            body_parts.append(rendered)

    if list_open:
        body_parts.append("</ul>")

    return f"""
    <section id=\"{sec_id}\" class=\"chapter\">
      <header class=\"chapter-head\">
        <span class=\"kicker\">Chapter {index}</span>
        <h2>{html.escape(title)}</h2>
      </header>
      <article class=\"chapter-body\">
        {''.join(body_parts)}
      </article>
    </section>
    """


def build_html(data: dict) -> str:
    toc_items = []
    sections_html = []

    for idx, title in enumerate(SECTION_ORDER, start=1):
        content = str(data.get(title, "")).strip()
        if not content:
            content = "Content pending for this chapter."
        sec_id = f"chapter-{idx}-{slugify(title)}"
        toc_items.append(f"<li><a href=\"#{sec_id}\"><span>{idx:02d}</span> {html.escape(title)}</a></li>")
        sections_html.append(render_section(title, content, idx))

    return f"""<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>Tally Study Book - English Edition</title>
  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">
  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>
  <link href=\"https://fonts.googleapis.com/css2?family=Playfair+Display:wght@500;700;800&family=Source+Serif+4:opsz,wght@8..60,400;8..60,600&family=Space+Grotesk:wght@500;700&display=swap\" rel=\"stylesheet\">
  <style>
    :root {{
      --bg: #f6f0e3;
      --paper: #fffaf0;
      --ink: #1e1a17;
      --muted: #6d5a47;
      --accent: #0f6c5b;
      --accent-soft: #dcefe9;
      --line: #d9c9af;
      --shadow: 0 12px 34px rgba(52, 36, 19, 0.11);
    }}

    * {{ box-sizing: border-box; }}

    body {{
      margin: 0;
      color: var(--ink);
      background:
        radial-gradient(circle at 15% 0%, #fdf8ef 0%, transparent 45%),
        radial-gradient(circle at 85% 10%, #efe5d1 0%, transparent 40%),
        var(--bg);
      font-family: "Source Serif 4", Georgia, serif;
      line-height: 1.72;
    }}

    .frame {{
      max-width: 1180px;
      margin: 24px auto;
      padding: 0 16px;
      display: grid;
      grid-template-columns: 320px 1fr;
      gap: 18px;
    }}

    .toc {{
      position: sticky;
      top: 18px;
      align-self: start;
      border: 1px solid var(--line);
      border-radius: 16px;
      background: rgba(255, 250, 240, 0.9);
      backdrop-filter: blur(3px);
      box-shadow: var(--shadow);
      padding: 20px;
      max-height: calc(100vh - 36px);
      overflow: auto;
    }}

    .toc h1 {{
      font: 800 1.55rem/1.1 "Playfair Display", serif;
      margin: 0;
      letter-spacing: .3px;
    }}

    .toc .sub {{
      color: var(--muted);
      margin: 8px 0 18px;
      font-family: "Space Grotesk", sans-serif;
      font-size: .88rem;
      text-transform: uppercase;
      letter-spacing: .9px;
    }}

    .toc ul {{
      list-style: none;
      margin: 0;
      padding: 0;
    }}

    .toc li + li {{ margin-top: 8px; }}

    .toc a {{
      color: var(--ink);
      text-decoration: none;
      display: block;
      border-radius: 10px;
      padding: 10px 12px;
      border: 1px solid transparent;
      transition: .2s ease;
      font-size: .95rem;
    }}

    .toc a span {{
      color: var(--accent);
      font-family: "Space Grotesk", sans-serif;
      font-weight: 700;
      margin-right: 8px;
    }}

    .toc a:hover {{
      border-color: var(--line);
      background: #fff;
      transform: translateX(2px);
    }}

    .book {{
      border: 1px solid var(--line);
      border-radius: 20px;
      background: var(--paper);
      box-shadow: var(--shadow);
      overflow: hidden;
    }}

    .cover {{
      padding: 52px 54px 44px;
      background:
        linear-gradient(130deg, rgba(15, 108, 91, .1), rgba(255, 255, 255, .4)),
        repeating-linear-gradient(90deg, rgba(15,108,91,.08), rgba(15,108,91,.08) 1px, transparent 1px, transparent 20px);
      border-bottom: 1px solid var(--line);
    }}

    .cover .edition {{
      display: inline-block;
      background: var(--accent-soft);
      color: var(--accent);
      border: 1px solid #b6dbd2;
      border-radius: 999px;
      padding: 6px 12px;
      font: 700 .78rem/1 "Space Grotesk", sans-serif;
      letter-spacing: .8px;
      text-transform: uppercase;
    }}

    .cover h1 {{
      margin: 18px 0 8px;
      font: 800 clamp(2rem, 4vw, 3.1rem)/1.04 "Playfair Display", serif;
    }}

    .cover p {{
      max-width: 62ch;
      margin: 0;
      color: var(--muted);
      font-size: 1.02rem;
    }}

    .chapter {{
      padding: 34px 54px;
      border-top: 1px solid var(--line);
      scroll-margin-top: 22px;
    }}

    .chapter-head {{ margin-bottom: 18px; }}

    .kicker {{
      display: inline-block;
      font: 700 .75rem/1 "Space Grotesk", sans-serif;
      letter-spacing: 1px;
      color: var(--accent);
      background: var(--accent-soft);
      border-radius: 999px;
      padding: 6px 10px;
      text-transform: uppercase;
    }}

    .chapter h2 {{
      margin: 12px 0 0;
      font: 700 clamp(1.5rem, 3vw, 2rem)/1.18 "Playfair Display", serif;
    }}

    .chapter h3 {{
      margin: 1.2rem 0 .5rem;
      font: 600 1.14rem/1.4 "Space Grotesk", sans-serif;
      color: #2f2821;
    }}

    .chapter p {{ margin: .35rem 0 .7rem; font-size: 1.03rem; }}

    .chapter p:first-of-type::first-letter {{
      float: left;
      font: 700 2.6rem/1 "Playfair Display", serif;
      padding-right: .32rem;
      color: var(--accent);
    }}

    .chapter ul {{
      margin: .7rem 0 1rem 1.2rem;
      padding: 0;
    }}

    .chapter li {{ margin: .3rem 0; }}

    @media (max-width: 980px) {{
      .frame {{
        grid-template-columns: 1fr;
        margin-top: 12px;
      }}

      .toc {{
        position: relative;
        top: 0;
        max-height: none;
      }}
    }}

    @media (max-width: 640px) {{
      .cover, .chapter {{ padding: 26px 20px; }}
      .chapter p {{ font-size: 1rem; }}
    }}
  </style>
</head>
<body>
  <div class=\"frame\">
    <aside class=\"toc\">
      <h1>Tally Study Book</h1>
      <p class=\"sub\">English Edition</p>
      <ul>
        {''.join(toc_items)}
      </ul>
    </aside>

    <main class=\"book\">
      <section class=\"cover\">
        <span class=\"edition\">Readable Book Format</span>
        <h1>Tally Study Companion</h1>
        <p>This version is formatted for readability with clear chapter structure, typographic hierarchy, and a maintainable single-source HTML layout.</p>
      </section>
      {''.join(sections_html)}
    </main>
  </div>
</body>
</html>
"""


def main() -> int:
    if not CHECKPOINT.exists():
        print(f"Checkpoint not found: {CHECKPOINT}")
        return 1

    data = json.loads(CHECKPOINT.read_text(encoding="utf-8"))
    html_text = build_html(data)
    OUTPUT_HTML.write_text(html_text, encoding="utf-8")
    print(f"Created: {OUTPUT_HTML}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
