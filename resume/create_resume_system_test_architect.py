from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


BASE_DIR = Path(r"C:/Users/MOKAS10/vcs/csrs-vcs-core")
DOCX_PATH = BASE_DIR / "Mohammad_Kashif_System_Test_Architect_Resume.docx"
DOC_PATH = BASE_DIR / "Mohammad_Kashif_System_Test_Architect_Resume.doc"

ACCENT = "1F4E79"
ACCENT_LIGHT = "D9EAF7"
TEXT_DARK = RGBColor(31, 41, 55)
TEXT_MUTED = RGBColor(75, 85, 99)
SECTION_COLOR = RGBColor(31, 78, 121)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_paragraph_bottom_border(paragraph, color=ACCENT, size="8"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def set_table_borders(table, color="B8CCE4", size="6"):
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        borders.append(border)
    tbl_pr.append(borders)


def add_text(paragraph, text, *, bold=False, size=10, color=TEXT_DARK, highlight=None):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.color.rgb = color
    if highlight:
        run.font.highlight_color = highlight
    return run


def add_section_heading(document, title):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(6)
    paragraph.paragraph_format.space_after = Pt(2)
    add_text(paragraph, title.upper(), bold=True, size=10.5, color=SECTION_COLOR)
    set_paragraph_bottom_border(paragraph)


HIGHLIGHT_TOKENS = [
    "15+ years",
    "system-level test",
    "test architecture",
    "test strategy",
    "test specification",
    "hardware/software",
    "hardware and software",
    "coverage analysis",
    "coverage definition",
    "requirements analysis",
    "test lead",
    "verification",
    "validation",
    "system integration",
    "SoC",
    "Android",
    "iOS",
    "Linux",
    "cross-domain",
    "ISTQB",
    "functional safety",
    "FuSA",
    "CI/CD",
    "automation",
    "Playwright",
    "Selenium",
    "Appium",
    "POS devices",
    "device testing",
    "hardware testing",
    "mentoring",
    "cross-functional",
    "quality engineering",
    "GCP",
    "DevOps",
]


def add_highlighted_bullet(document, text):
    paragraph = document.add_paragraph(style="List Bullet")
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    for token in HIGHLIGHT_TOKENS:
        if token.lower() in text.lower():
            idx = text.lower().index(token.lower())
            before = text[:idx]
            matched = text[idx : idx + len(token)]
            after = text[idx + len(token) :]
            add_text(paragraph, before, size=10.5)
            add_text(paragraph, matched, bold=True, size=10.5, highlight=WD_COLOR_INDEX.YELLOW)
            add_text(paragraph, after, size=10.5)
            return
    add_text(paragraph, text, size=10.5)


def build_docx():
    document = Document()

    section = document.sections[0]
    section.top_margin = Cm(0.8)
    section.bottom_margin = Cm(0.8)
    section.left_margin = Cm(1.2)
    section.right_margin = Cm(1.2)

    # ── Header ──
    header = document.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header.paragraph_format.space_after = Pt(0)
    add_text(header, "MOHAMMAD KASHIF", bold=True, size=18, color=SECTION_COLOR)

    role = document.add_paragraph()
    role.alignment = WD_ALIGN_PARAGRAPH.CENTER
    role.paragraph_format.space_before = Pt(0)
    role.paragraph_format.space_after = Pt(1)
    add_text(
        role,
        "System Test Architect | Hardware/Software Verification | Test Strategy & Leadership",
        bold=True,
        size=10.5,
    )

    contact = document.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact.paragraph_format.space_before = Pt(0)
    contact.paragraph_format.space_after = Pt(2)
    add_text(contact, "Malmö, Sweden | +46 702624230 | mo.kashif@gmail.com", size=9.5, color=TEXT_MUTED)
    add_text(
        contact,
        "  |  LinkedIn: linkedin.com/in/md-kashif  |  GitHub: github.com/ascertain",
        size=9.5,
        color=TEXT_MUTED,
    )
    set_paragraph_bottom_border(contact, color="7F8FA6", size="10")

    # ── Profile ──
    add_section_heading(document, "Profile")
    profile = document.add_paragraph()
    profile.paragraph_format.space_after = Pt(2)
    profile.paragraph_format.line_spacing = 1.0
    add_text(
        profile,
        "System Test Architect with 15+ years of experience in system-level test strategy, "
        "test architecture, and verification & validation across complex hardware/software systems. "
        "Proven track record leading test teams, defining requirements-driven test specifications, "
        "driving coverage analysis, and ensuring system integration quality across embedded, mobile, "
        "and cloud-native platforms. Hands-on experience testing Android and iOS applications across "
        "diverse device versions and hardware configurations, including POS devices. Strong test lead "
        "background — mentoring engineers, aligning cross-functional teams, and bridging domain "
        "boundaries. ISTQB certified with CI/CD pipelines, automation frameworks (Playwright, "
        "Selenium, Appium), and DevOps practices.",
        size=10,
    )

    # ── Core Competencies ──
    add_section_heading(document, "Core Competencies")
    skill_lines = [
        (
            "System Test Architecture: ",
            "System-level test strategy, test specification authoring, requirements analysis, "
            "coverage definition & analysis, traceability, V&V planning",
        ),
        (
            "Hardware/Software Testing: ",
            "Hardware testing across device variants, Android/iOS app testing on multiple OS levels, "
            "POS device testing, SoC-based verification, embedded platform validation, cross-domain integration",
        ),
        (
            "Automation & Frameworks: ",
            "Playwright, Selenium, Appium, Karate, TDD/BDD, API testing, E2E automation, UAT, BAT",
        ),
        (
            "Test Leadership: ",
            "Test lead, team mentoring, cross-functional collaboration, stakeholder management, Agile/Scrum, release readiness",
        ),
        (
            "DevOps & CI/CD: ",
            "GCP, Cloud Run, GitHub Actions, Docker, Kubernetes, Terraform, CI/CD pipelines",
        ),
        (
            "Domain: ",
            "GPU/SoC awareness, Linux/Android platforms, mobile, e-commerce, core banking (Finacle), government payment systems",
        ),
    ]
    for label, value in skill_lines:
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        add_text(paragraph, label, bold=True, size=10)
        add_text(paragraph, value, size=10)

    # ── Professional Experience ──
    add_section_heading(document, "Professional Experience")
    jobs = [
        (
            "IKEA IT AB, Malmö, Sweden",
            "Software Engineer / System Test & Platform Leadership Contributor | Mar 2022 – Present",
            [
                "Define and drive system-level test strategy for VCS — a complex hardware/software platform "
                "integrating CoBrowser, Real-Time Transcription, and AI-powered Call Quality Assessment.",
                "Lead requirements analysis, author test specifications with full coverage definition and "
                "traceability for system integration validation.",
                "Architect E2E test automation using Playwright — verification across browsers, devices, and "
                "OS configurations with CI/CD via GitHub Actions.",
                "Perform hardware/software integration testing across Android and iOS devices on multiple OS "
                "versions — validating audio, video, network quality across hardware configurations.",
                "Drive coverage analysis across functional, non-functional, and cross-domain boundaries; "
                "lead and mentor cross-functional engineering team.",
                "Built data layer on GCP (BigQuery, Pub/Sub, Cloud Functions) for system telemetry and quality analytics. "
                "Recognized as Exceptional Performer.",
            ],
        ),
        (
            "Truecaller, Sweden",
            "Release & Automation Engineer | Sep 2021 – Feb 2022",
            [
                "Led verification for Android and iOS apps serving 300M+ users — testing across 50+ device "
                "models and OS versions for hardware/software compatibility.",
                "Defined test strategy for release readiness across device configurations, chipsets, and OS versions. "
                "Built automated suites using Appium and Selenium for device-level testing.",
                "Collaborated on requirements analysis, defect triage, and cross-functional release coordination.",
            ],
        ),
        (
            "LEGO and IKEA Group (via HCLTech), Denmark & Sweden",
            "Technical Specialist / Test Lead | 2016 – 2021",
            [
                "Led system-level test architecture for LEGO digital commerce and IKEA App — defining test "
                "strategies, specifications, and coverage models for complex hardware/software systems.",
                "Managed hardware testing for Android and iOS apps across 30+ device variants and OS versions. "
                "Test lead managing 8+ engineers with mentoring and Agile delivery.",
                "Built automation using Selenium, Appium, and Karate (BDD/TDD) — integrated into CI/CD. "
                "Drove shift-left testing and requirements traceability through system integration.",
            ],
        ),
        (
            "SW Global — AMA Accra Metropolitan Assembly, Ghana",
            "Technical Specialist / Test Engineer | 2014 – 2016",
            [
                "Tested AMA-IGRA tax collection system — hardware/software solution integrating web portal "
                "with POS devices for government payment processing.",
                "Performed hardware testing on POS terminals — validating payment flows, firmware interactions, "
                "and device configurations. Authored test specifications for end-to-end transaction verification.",
            ],
        ),
        (
            "HCLTech / Enterprise Programs",
            "Automation Lead / SDET | Dec 2013 – 2014",
            [
                "Led automation and test architecture — scalable frameworks, CI/CD pipelines, and quality "
                "dashboards. Built system-level test automation reducing manual effort by 70%.",
                "Mentored 15+ engineers on test strategy, automation architecture, and cross-domain testing.",
            ],
        ),
        (
            "Earlier Career — Banking & Government",
            "Software Test Engineer / Consultant | 2008 – 2013",
            [
                "System integration testing for core banking (Finacle) — hardware/software interactions across "
                "ATM networks, POS terminals, and backend systems. Tested Android and iOS banking apps.",
                "Led UAT and BAT cycles for government and financial platforms — requirements-driven test "
                "specifications with full coverage tracking.",
            ],
        ),
    ]
    for company, title, bullets in jobs:
        cp = document.add_paragraph()
        cp.paragraph_format.space_before = Pt(3)
        cp.paragraph_format.space_after = Pt(0)
        add_text(cp, company, bold=True, size=10, color=SECTION_COLOR)

        tp = document.add_paragraph()
        tp.paragraph_format.space_after = Pt(0)
        add_text(tp, title, bold=True, size=10)

        for bullet in bullets:
            add_highlighted_bullet(document, bullet)

    # ── Key Achievements ──
    add_section_heading(document, "Key Achievements")
    achievements = [
        "Defined system-level test strategy across complex hardware/software platforms — V&V from "
        "requirements through system integration.",
        "Led hardware testing across 50+ Android and iOS device variants — 99.5% device compatibility.",
        "Migrated Selenium to Playwright — 3x reliability, 50% faster CI pipelines.",
        "Tested POS devices and government payment systems (AMA-IGRA, Ghana) — end-to-end transaction verification.",
        "Led and mentored teams of 8-15+ engineers — driving quality culture and test architecture standards.",
    ]
    for achievement in achievements:
        add_highlighted_bullet(document, achievement)

    # ── Education & Certifications ──
    add_section_heading(document, "Education & Certifications")
    table = document.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table, color="9DBAD5")

    for cell, title in zip(table.rows[0].cells, ["Education", "Certifications"]):
        set_cell_shading(cell, "EEF4FA")
        p = cell.paragraphs[0]
        add_text(p, title, bold=True, size=10, color=SECTION_COLOR)

    edu = table.cell(1, 0).paragraphs[0]
    add_text(edu, "M.Tech / B.Tech in Information Technology\nPG Diploma in Operations Management", size=9.5)
    certs = table.cell(1, 1).paragraphs[0]
    add_text(
        certs,
        "ISTQB Certified Tester Foundation\nGoogle Cloud – Associate Cloud Engineer\n"
        "AWS Cloud Practitioner\nITIL Foundation\nCertified Ethical Hacker (CEH)\n"
        "Six Sigma Green Belt\nUiPath RPA Certified",
        size=9.5,
    )

    # ── Languages ──
    add_section_heading(document, "Languages")
    lp = document.add_paragraph()
    lp.paragraph_format.space_after = Pt(0)
    add_text(lp, "English — Fluent  |  Swedish — Basic  |  Hindi / Urdu — Native", size=10)

    document.save(DOCX_PATH)


def build_doc():
    html = """\
<html>
<head>
  <meta charset="utf-8">
  <title>Mohammad Kashif – System Test Architect Resume</title>
  <style>
    body { font-family: Calibri, Arial, sans-serif; margin: 24px; color: #1f2937; font-size: 10pt; }
    h1 { text-align: center; color: #1f4e79; margin-bottom: 2px; font-size: 18pt; }
    h2 { text-align: center; font-size: 10.5pt; margin-top: 0; margin-bottom: 4px; }
    .contact { text-align: center; color: #4b5563; font-size: 9.5pt; border-bottom: 2px solid #1f4e79; padding-bottom: 6px; margin-bottom: 10px; }
    .section { color: #1f4e79; font-weight: 700; font-size: 10.5pt; border-bottom: 1.5px solid #1f4e79; padding-bottom: 2px; margin-top: 8px; margin-bottom: 4px; text-transform: uppercase; }
    .job-title { font-weight: 700; color: #1f4e79; margin-top: 4px; margin-bottom: 0; }
    .job-sub { font-weight: 700; margin-bottom: 2px; }
    ul { margin-top: 2px; margin-bottom: 4px; padding-left: 20px; }
    li { margin-bottom: 2px; }
    .hl { background: #fff59d; font-weight: 700; }
    table { width: 100%; border-collapse: collapse; margin-top: 4px; }
    td { border: 1px solid #b8cce4; padding: 5px; vertical-align: top; font-size: 9.5pt; }
    .tag { background: #d9eaf7; font-weight: 700; }
    p { margin: 2px 0; }
  </style>
</head>
<body>
  <h1>MOHAMMAD KASHIF</h1>
  <h2>System Test Architect | Hardware/Software Verification | Test Strategy &amp; Leadership</h2>
  <div class="contact">Malmö, Sweden | +46 702624230 | mo.kashif@gmail.com  |  LinkedIn: linkedin.com/in/md-kashif  |  GitHub: github.com/ascertain</div>

  <div class="section">Profile</div>
  <p>System Test Architect with <span class="hl">15+ years</span> of experience in <span class="hl">system-level test</span> strategy, <span class="hl">test architecture</span>, and <span class="hl">verification</span> &amp; <span class="hl">validation</span> across complex <span class="hl">hardware/software</span> systems. Proven track record leading test teams, defining requirements-driven <span class="hl">test specification</span>s, driving <span class="hl">coverage analysis</span>, and ensuring <span class="hl">system integration</span> quality across embedded, mobile, and cloud-native platforms. Hands-on experience testing <span class="hl">Android</span> and <span class="hl">iOS</span> applications across diverse device versions and hardware configurations, including <span class="hl">POS devices</span>. Strong <span class="hl">test lead</span> background — <span class="hl">mentoring</span> engineers, aligning <span class="hl">cross-functional</span> teams, and bridging domain boundaries. <span class="hl">ISTQB</span> certified with <span class="hl">CI/CD</span> pipelines, <span class="hl">automation</span> frameworks (<span class="hl">Playwright</span>, <span class="hl">Selenium</span>, <span class="hl">Appium</span>), and <span class="hl">DevOps</span> practices.</p>

  <div class="section">Core Competencies</div>
  <p><b>System Test Architecture:</b> System-level test strategy, test specification authoring, requirements analysis, coverage definition &amp; analysis, traceability, V&amp;V planning<br>
  <b>Hardware/Software Testing:</b> Hardware testing across device variants, Android/iOS app testing on multiple OS levels, POS device testing, SoC-based verification, embedded platform validation, cross-domain integration<br>
  <b>Automation &amp; Frameworks:</b> Playwright, Selenium, Appium, Karate, TDD/BDD, API testing, E2E automation, UAT, BAT<br>
  <b>Test Leadership:</b> Test lead, team mentoring, cross-functional collaboration, stakeholder management, Agile/Scrum, release readiness<br>
  <b>DevOps &amp; CI/CD:</b> GCP, Cloud Run, GitHub Actions, Docker, Kubernetes, Terraform, CI/CD pipelines<br>
  <b>Domain:</b> GPU/SoC awareness, Linux/Android platforms, mobile, e-commerce, core banking (Finacle), government payment systems</p>

  <div class="section">Professional Experience</div>

  <div class="job-title">IKEA IT AB, Malmö, Sweden</div>
  <div class="job-sub">Software Engineer / System Test &amp; Platform Leadership Contributor | Mar 2022 – Present</div>
  <ul>
    <li>Define and drive <span class="hl">system-level test</span> strategy for VCS — a complex <span class="hl">hardware/software</span> platform integrating CoBrowser, Real-Time Transcription, and AI-powered Call Quality Assessment.</li>
    <li>Lead <span class="hl">requirements analysis</span>, author <span class="hl">test specification</span>s with full <span class="hl">coverage definition</span> and traceability for system integration validation.</li>
    <li>Architect E2E <span class="hl">automation</span> using <span class="hl">Playwright</span> — <span class="hl">verification</span> across browsers, devices, and OS configurations with <span class="hl">CI/CD</span> via GitHub Actions.</li>
    <li>Perform <span class="hl">hardware testing</span> across <span class="hl">Android</span> and <span class="hl">iOS</span> devices on multiple OS versions — validating audio, video, network quality across hardware configurations.</li>
    <li>Drive <span class="hl">coverage analysis</span> across functional, non-functional, and <span class="hl">cross-domain</span> boundaries; lead and mentor <span class="hl">cross-functional</span> engineering team.</li>
    <li>Built data layer on <span class="hl">GCP</span> (BigQuery, Pub/Sub, Cloud Functions) for system telemetry and quality analytics. Recognized as Exceptional Performer.</li>
  </ul>

  <div class="job-title">Truecaller, Sweden</div>
  <div class="job-sub">Release &amp; Automation Engineer | Sep 2021 – Feb 2022</div>
  <ul>
    <li>Led <span class="hl">verification</span> for <span class="hl">Android</span> and <span class="hl">iOS</span> apps serving 300M+ users — testing across 50+ device models and OS versions for <span class="hl">hardware/software</span> compatibility.</li>
    <li>Defined <span class="hl">test strategy</span> for release readiness across device configurations, chipsets, and OS versions. Built automated suites using <span class="hl">Appium</span> and <span class="hl">Selenium</span> for <span class="hl">device testing</span>.</li>
    <li>Collaborated on <span class="hl">requirements analysis</span>, defect triage, and <span class="hl">cross-functional</span> release coordination.</li>
  </ul>

  <div class="job-title">LEGO and IKEA Group (via HCLTech), Denmark &amp; Sweden</div>
  <div class="job-sub">Technical Specialist / Test Lead | 2016 – 2021</div>
  <ul>
    <li>Led <span class="hl">system-level test</span> architecture for LEGO digital commerce and IKEA App — defining strategies, specifications, and coverage models for complex <span class="hl">hardware/software</span> systems.</li>
    <li>Managed <span class="hl">hardware testing</span> for <span class="hl">Android</span> and <span class="hl">iOS</span> apps across 30+ device variants and OS versions. <span class="hl">Test lead</span> managing 8+ engineers with <span class="hl">mentoring</span> and Agile delivery.</li>
    <li>Built <span class="hl">automation</span> using <span class="hl">Selenium</span>, <span class="hl">Appium</span>, and Karate (BDD/TDD) — integrated into <span class="hl">CI/CD</span>. Drove shift-left testing and requirements traceability through <span class="hl">system integration</span>.</li>
  </ul>

  <div class="job-title">SW Global — AMA Accra Metropolitan Assembly, Ghana</div>
  <div class="job-sub">Technical Specialist / Test Engineer | 2014 – 2016</div>
  <ul>
    <li>Tested AMA-IGRA tax collection system — <span class="hl">hardware/software</span> solution integrating web portal with <span class="hl">POS devices</span> for government payment processing.</li>
    <li>Performed <span class="hl">hardware testing</span> on POS terminals — validating payment flows, firmware interactions, and device configurations. Authored <span class="hl">test specification</span>s for end-to-end transaction <span class="hl">verification</span>.</li>
  </ul>

  <div class="job-title">HCLTech / Enterprise Programs</div>
  <div class="job-sub">Automation Lead / SDET | Dec 2013 – 2014</div>
  <ul>
    <li>Led <span class="hl">automation</span> and <span class="hl">test architecture</span> — scalable frameworks, <span class="hl">CI/CD</span> pipelines, and quality dashboards. Built <span class="hl">system-level test</span> <span class="hl">automation</span> reducing manual effort by 70%.</li>
    <li>Mentored 15+ engineers on <span class="hl">test strategy</span>, <span class="hl">automation</span> architecture, and <span class="hl">cross-domain</span> testing.</li>
  </ul>

  <div class="job-title">Earlier Career — Banking &amp; Government</div>
  <div class="job-sub">Software Test Engineer / Consultant | 2008 – 2013</div>
  <ul>
    <li><span class="hl">System integration</span> testing for core banking (Finacle) — <span class="hl">hardware/software</span> interactions across ATM networks, POS terminals, and backend systems. Tested <span class="hl">Android</span> and <span class="hl">iOS</span> banking apps.</li>
    <li>Led UAT and BAT cycles for government and financial platforms — requirements-driven <span class="hl">test specification</span>s with full coverage tracking.</li>
  </ul>

  <div class="section">Key Achievements</div>
  <ul>
    <li>Defined <span class="hl">system-level test</span> strategy across complex <span class="hl">hardware/software</span> platforms — V&amp;V from requirements through <span class="hl">system integration</span>.</li>
    <li>Led <span class="hl">hardware testing</span> across 50+ <span class="hl">Android</span> and <span class="hl">iOS</span> device variants — 99.5% device compatibility.</li>
    <li>Migrated <span class="hl">Selenium</span> to <span class="hl">Playwright</span> — 3x reliability, 50% faster CI pipelines.</li>
    <li>Tested <span class="hl">POS devices</span> and government payment systems (AMA-IGRA, Ghana) — end-to-end transaction <span class="hl">verification</span>.</li>
    <li>Led and mentored teams of 8-15+ engineers — driving quality culture and <span class="hl">test architecture</span> standards.</li>
  </ul>

  <div class="section">Education &amp; Certifications</div>
  <table>
    <tr><td class="tag">Education</td><td class="tag">Certifications</td></tr>
    <tr>
      <td>M.Tech / B.Tech in Information Technology<br>PG Diploma in Operations Management</td>
      <td>ISTQB Certified Tester Foundation<br>GCP Associate Cloud Engineer<br>AWS Cloud Practitioner<br>ITIL Foundation<br>Certified Ethical Hacker<br>Six Sigma Green Belt<br>UiPath RPA</td>
    </tr>
  </table>

  <div class="section">Languages</div>
  <p>English — Fluent  |  Swedish — Basic  |  Hindi / Urdu — Native</p>
</body>
</html>
"""
    DOC_PATH.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    build_docx()
    build_doc()
    print(DOCX_PATH)
    print(DOC_PATH)
