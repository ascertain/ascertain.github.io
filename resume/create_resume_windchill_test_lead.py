from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


BASE_DIR = Path(r"C:/Users/MOKAS10/vcs/csrs-vcs-core")
DOCX_PATH = BASE_DIR / "Mohammad_Kashif_Windchill_Test_Lead_Resume.docx"
DOC_PATH = BASE_DIR / "Mohammad_Kashif_Windchill_Test_Lead_Resume.doc"

ACCENT = "1F4E79"
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
    "Windchill",
    "PLM",
    "test strategy",
    "test plan",
    "test lead",
    "UAT",
    "integration",
    "regression",
    "performance",
    "system testing",
    "defect management",
    "quality metrics",
    "stakeholder",
    "lifecycle",
    "workflow",
    "access control",
    "BOM",
    "Change Management",
    "Document Management",
    "CAD",
    "Creo",
    "ERP",
    "SAP",
    "migration",
    "upgrade",
    "cloud",
    "Agile",
    "Waterfall",
    "ISTQB",
    "automation",
    "Selenium",
    "Playwright",
    "CI/CD",
    "Azure DevOps",
    "Jira",
    "HP ALM",
    "SDLC",
    "entry/exit criteria",
    "test coverage",
    "risk",
    "triage",
    "root-cause",
    "regulated",
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
            add_text(paragraph, before, size=10)
            add_text(paragraph, matched, bold=True, size=10, highlight=WD_COLOR_INDEX.YELLOW)
            add_text(paragraph, after, size=10)
            return
    add_text(paragraph, text, size=10)


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
        "Test Lead  |  PLM/Windchill, Enterprise Integrations & Quality Assurance  |  10+ Years",
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
        "  |  LinkedIn: linkedin.com/in/md-kashif  |  Blog: ascertain.github.io",
        size=9.5,
        color=TEXT_MUTED,
    )
    set_paragraph_bottom_border(contact, color="0058A3", size="10")

    # ── Profile ──
    add_section_heading(document, "Profile")
    profile = document.add_paragraph()
    profile.paragraph_format.space_after = Pt(2)
    profile.paragraph_format.line_spacing = 1.0
    add_text(
        profile,
        "Test Lead with 10+ years of experience defining testing strategies, creating "
        "test plans, managing risks, and reporting on quality metrics across large-scale "
        "enterprise implementations, upgrades, and integrations. Combines technical testing "
        "excellence with team leadership to ensure high-quality product delivery. Experienced "
        "in end-to-end testing across PLM/enterprise systems — system, integration, regression, "
        "performance, and UAT. Strong background in enterprise platform testing including "
        "configuration, workflows, lifecycle states, access control, and integrations with "
        "ERP, CAD, and downstream systems. Manages defect lifecycle including triage, "
        "prioritisation, root-cause analysis, and retesting. Communicates test progress, "
        "risks, and quality metrics to stakeholders at all levels. Drives quality best "
        "practices and continuous improvement. ISTQB certified. Agile and Waterfall.",
        size=10,
    )

    # ── Core Competencies ──
    add_section_heading(document, "Core Competencies")
    skill_lines = [
        (
            "Test Leadership & Strategy: ",
            "Test strategy definition, test plans, schedules, entry/exit criteria. "
            "Task assignment and team coordination. Risk management. Quality metrics "
            "and reporting. Stakeholder communication. Continuous improvement",
        ),
        (
            "Testing Types: ",
            "System testing, integration testing, regression testing, performance "
            "testing, UAT. End-to-end validation. Smoke, sanity, exploratory. "
            "Data migration testing. Upgrade and cloud deployment testing",
        ),
        (
            "Enterprise & PLM Systems: ",
            "Enterprise platform testing — configuration, workflows, lifecycle states, "
            "access control, BOM structures, Change Management, Document Management. "
            "Integrations with ERP (SAP), CAD (Creo, Inventor), and downstream systems",
        ),
        (
            "Defect Management: ",
            "Defect triage, prioritisation, root-cause analysis, retesting. "
            "Defect tracking tools (Jira, Azure DevOps, HP ALM). Quality gate "
            "decisions. Test coverage and traceability",
        ),
        (
            "Automation & Tools: ",
            "Playwright, Selenium, Cypress. API testing (REST, GraphQL). CI/CD "
            "(GitHub Actions, Jenkins, Azure DevOps). Python, TypeScript, C#, Java. "
            "TestRail, Jira, Confluence, HP ALM. Grafana (dashboards)",
        ),
        (
            "Methodologies & Standards: ",
            "Agile (Scrum/SAFe), Waterfall. ISTQB certified. SDLC/STLC. "
            "Six Sigma Green Belt. ITIL Foundation. Regulated environments. "
            "Audit-ready documentation",
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

    # -- IKEA --
    cp1 = document.add_paragraph()
    cp1.paragraph_format.space_before = Pt(3)
    cp1.paragraph_format.space_after = Pt(0)
    add_text(cp1, "IKEA IT AB, Malmö — Test Lead / Team Lead", bold=True, size=10, color=SECTION_COLOR)
    tp1 = document.add_paragraph()
    tp1.paragraph_format.space_after = Pt(0)
    add_text(tp1, "Mar 2022 – Present  |  Enterprise Platform — Configuration, Workflows, Integrations, 30+ Markets", bold=True, size=10)
    ikea_bullets = [
        "Defined and owned the test strategy, test plans, schedules, and entry/exit "
        "criteria for a complex enterprise platform with configuration, workflows, "
        "lifecycle states, and access control. Led end-to-end testing activities "
        "across implementations, upgrades, and enhancements serving 30+ global markets.",
        "Managed system, integration, regression, performance, and UAT across platform "
        "modules and integrations with external SaaS and downstream enterprise systems. "
        "Ensured test coverage through requirements validation and traceability. "
        "Designed, reviewed, and maintained test cases, scenarios, and test data.",
        "Led defect management — triage, prioritisation, root-cause analysis, and "
        "retesting. Communicated test progress, risks, and quality metrics to "
        "stakeholders including program leadership, business, and vendors. Supported "
        "UAT by guiding business users, managing test cycles, and ensuring signoff readiness.",
        "Drove quality best practices and continuous improvement — led Selenium-to-Playwright "
        "migration (3x faster execution, 50% CI reduction). Automation: Playwright, API "
        "testing, CI/CD (GitHub Actions). Python, TypeScript. Grafana dashboards for "
        "reporting. AI-assisted — 30% velocity improvement.",
        "Team leadership — managed engineers and consultants. Coaching, mentoring, "
        "competence development. Coordinated with SaaS vendor on release alignment "
        "and defect resolution. Bridged business and technology stakeholders. "
        "Recognized as Exceptional Performer.",
    ]
    for b in ikea_bullets:
        add_highlighted_bullet(document, b)

    # -- Truecaller --
    cp2 = document.add_paragraph()
    cp2.paragraph_format.space_before = Pt(3)
    cp2.paragraph_format.space_after = Pt(0)
    add_text(cp2, "Truecaller, Stockholm — Release & Test Lead", bold=True, size=10, color=SECTION_COLOR)
    tp2 = document.add_paragraph()
    tp2.paragraph_format.space_after = Pt(0)
    add_text(tp2, "Sep 2021 – Feb 2022  |  Enterprise Platform — 300M+ Users, Distributed Architecture", bold=True, size=10)
    tc_bullets = [
        "Owned test strategy and release readiness for a distributed enterprise platform "
        "serving 300M+ users. Defined entry/exit criteria, managed risk, and reported "
        "quality metrics. System, integration, and regression testing across microservices. "
        "Defect management and go/no-go decisions. Stakeholder communication.",
        "Release automation and CI/CD pipelines. Cross-functional coordination across "
        "engineering, product, and operations. Feature flag management. Agile (Scrum/Kanban). "
        "AWS cloud infrastructure.",
    ]
    for b in tc_bullets:
        add_highlighted_bullet(document, b)

    # -- LEGO/IKEA via HCLTech --
    cp3 = document.add_paragraph()
    cp3.paragraph_format.space_before = Pt(3)
    cp3.paragraph_format.space_after = Pt(0)
    add_text(cp3, "HCLTech — LEGO & IKEA Group, Denmark & Sweden — Test Lead / SDET Lead", bold=True, size=10, color=SECTION_COLOR)
    tp3 = document.add_paragraph()
    tp3.paragraph_format.space_after = Pt(0)
    add_text(tp3, "2013 – 2021  |  Enterprise, E-Commerce & PLM — Integrations, Multi-Partner, Upgrades", bold=True, size=10)
    lego_bullets = [
        "LEGO & IKEA (2017–2021): Led end-to-end testing across enterprise platforms "
        "with complex integrations (ERP, third-party systems, APIs). Defined test "
        "strategy, managed test plans and schedules. System, integration, regression, "
        "and UAT. Coordinated testing across modules and integration points. Managed "
        "upgrades and migrations. Led teams of 8–10 across onshore-offshore. "
        "Stakeholder communication and quality reporting. Agile and Waterfall.",
        "SDET Lead (2013–2017): Designed and scaled automation frameworks (Selenium, C#, "
        "NUnit). CI/CD integration. Full SDLC — test planning, execution, defect "
        "management, quality metrics reporting. Led RFP/RFI technical input and POC "
        "delivery. Tested enterprise configurations, workflows, and access controls. "
        "Managed testing in regulated environments. Mentored team members.",
    ]
    for b in lego_bullets:
        add_highlighted_bullet(document, b)

    # -- Banking / Earlier --
    cp4 = document.add_paragraph()
    cp4.paragraph_format.space_before = Pt(3)
    cp4.paragraph_format.space_after = Pt(0)
    add_text(cp4, "Banking & Enterprise — Test Lead / Consultant", bold=True, size=10, color=SECTION_COLOR)
    tp4 = document.add_paragraph()
    tp4.paragraph_format.space_after = Pt(0)
    add_text(tp4, "2008 – 2013  |  Core Banking (Finacle CBS) — Migrations, Integrations, Regulated", bold=True, size=10)
    fin_bullets = [
        "Led testing for enterprise banking platform (Finacle CBS) — system, integration, "
        "regression, and UAT. Tested configurations, workflows, lifecycle states, and "
        "access controls. Data migration testing between legacy and new platforms. "
        "ETL automation (Pentaho). Integration with biometric hardware. Defect management, "
        "quality reporting. Regulatory compliance. Audit-ready documentation. "
        "Stakeholder coordination across business and technology teams.",
    ]
    for b in fin_bullets:
        add_highlighted_bullet(document, b)

    # ── Key Achievements ──
    add_section_heading(document, "Key Achievements")
    achievements = [
        "Test strategy & quality ownership — defined and implemented test strategies "
        "across enterprise platforms at IKEA (30+ markets), Truecaller (300M+ users), "
        "LEGO, and banking. End-to-end quality across complex integrations and upgrades.",
        "Enterprise platform testing — configurations, workflows, lifecycle states, "
        "access control, BOM-like structures, document management. Integrations with "
        "ERP, CAD/SaaS, and downstream systems. Managed upgrades and cloud migrations.",
        "Defect management excellence — triage, prioritisation, root-cause analysis. "
        "Communicated quality metrics and risks to stakeholders. Ensured signoff "
        "readiness for UAT. Go-live quality gate decisions.",
        "Automation & continuous improvement — Playwright, Selenium, CI/CD. Led "
        "Selenium-to-Playwright migration (3x faster, 50% CI reduction). Built "
        "reusable frameworks. Scaled IKEA platform 2K→50K. Exceptional Performer.",
        "Team leadership — led teams of 4–15. Coaching, mentoring, competence "
        "development. Coordinated with vendors, architects, and business stakeholders. "
        "Multi-partner delivery. Agile and Waterfall.",
    ]
    for a in achievements:
        add_highlighted_bullet(document, a)

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
    add_text(edu, "B.Tech in Information Technology\nPG Diploma in Operations Management", size=9.5)
    certs = table.cell(1, 1).paragraphs[0]
    add_text(
        certs,
        "ISTQB Certified Tester Foundation\n"
        "Six Sigma Green Belt\n"
        "ITIL Foundation\n"
        "Google Cloud – Associate Cloud Engineer\n"
        "AWS Cloud Practitioner\n"
        "Certified Ethical Hacker (CEH)\n"
        "UiPath RPA Certified",
        size=9.5,
    )

    document.save(DOCX_PATH)


def build_doc():
    html = """\
<html>
<head>
  <meta charset="utf-8">
  <title>Mohammad Kashif – Test Lead / Windchill Resume</title>
  <style>
    body { font-family: Calibri, Arial, sans-serif; margin: 24px; color: #1f2937; font-size: 10pt; }
    h1 { text-align: center; color: #1f4e79; margin-bottom: 2px; font-size: 18pt; }
    h2 { text-align: center; font-size: 10.5pt; margin-top: 0; margin-bottom: 4px; }
    .contact { text-align: center; color: #4b5563; font-size: 9.5pt; border-bottom: 2px solid #0058a3; padding-bottom: 6px; margin-bottom: 10px; }
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
  <h2>Test Lead | PLM/Windchill, Enterprise Integrations &amp; Quality Assurance | 10+ Years</h2>
  <div class="contact">Malmö, Sweden | +46 702624230 | mo.kashif@gmail.com | LinkedIn: linkedin.com/in/md-kashif | Blog: ascertain.github.io</div>

  <div class="section">Profile</div>
  <p>Test Lead with 10+ years of experience defining <span class="hl">testing strategies</span>, creating <span class="hl">test plans</span>, managing <span class="hl">risk</span>s, and reporting on <span class="hl">quality metrics</span> across large-scale enterprise implementations, <span class="hl">upgrade</span>s, and <span class="hl">integration</span>s. Combines technical testing excellence with team <span class="hl">leadership</span> to ensure high-quality product delivery. Experienced in end-to-end testing across <span class="hl">PLM</span>/enterprise systems — <span class="hl">system testing</span>, <span class="hl">integration</span>, <span class="hl">regression</span>, <span class="hl">performance</span>, and <span class="hl">UAT</span>. Strong background in enterprise platform testing including configuration, <span class="hl">workflow</span>s, <span class="hl">lifecycle</span> states, <span class="hl">access control</span>, and integrations with <span class="hl">ERP</span>, <span class="hl">CAD</span>, and downstream systems. <span class="hl">Defect management</span> — <span class="hl">triage</span>, prioritisation, <span class="hl">root-cause</span> analysis, retesting. <span class="hl">Stakeholder</span> communication. <span class="hl">ISTQB</span> certified. <span class="hl">Agile</span> and <span class="hl">Waterfall</span>.</p>

  <div class="section">Core Competencies</div>
  <p><b>Test Leadership &amp; Strategy:</b> Test strategy, plans, schedules, entry/exit criteria. Task assignment, team coordination. Risk management. Quality metrics &amp; reporting. Stakeholder communication. Continuous improvement<br>
  <b>Testing Types:</b> System, integration, regression, performance, UAT. End-to-end. Smoke, sanity, exploratory. Data migration. Upgrade &amp; cloud deployment testing<br>
  <b>Enterprise &amp; PLM Systems:</b> Configuration, workflows, lifecycle states, access control, BOM structures, Change Management, Document Management. ERP (SAP), CAD (Creo, Inventor) integrations<br>
  <b>Defect Management:</b> Triage, prioritisation, root-cause analysis, retesting. Jira, Azure DevOps, HP ALM. Quality gates. Test coverage &amp; traceability<br>
  <b>Automation &amp; Tools:</b> Playwright, Selenium, Cypress. API testing (REST). CI/CD (GitHub Actions, Jenkins, Azure DevOps). Python, TypeScript, C#, Java. TestRail, Jira, Confluence, HP ALM<br>
  <b>Methodologies:</b> Agile (Scrum/SAFe), Waterfall. ISTQB. SDLC/STLC. Six Sigma GB. ITIL. Regulated environments. Audit-ready documentation</p>

  <div class="section">Professional Experience</div>

  <div class="job-title">IKEA IT AB, Malmö — Test Lead / Team Lead</div>
  <div class="job-sub">Mar 2022 – Present | Enterprise Platform — Configuration, Workflows, Integrations, 30+ Markets</div>
  <ul>
    <li>Defined and owned <span class="hl">test strategy</span>, <span class="hl">test plan</span>s, schedules, <span class="hl">entry/exit criteria</span> for complex enterprise platform with configuration, <span class="hl">workflow</span>s, <span class="hl">lifecycle</span> states, <span class="hl">access control</span>. Led end-to-end testing across implementations, <span class="hl">upgrade</span>s, enhancements — 30+ markets.</li>
    <li>Managed <span class="hl">system testing</span>, <span class="hl">integration</span>, <span class="hl">regression</span>, <span class="hl">performance</span>, and <span class="hl">UAT</span> across modules and <span class="hl">integration</span>s with external <span class="hl">SaaS</span>/<span class="hl">ERP</span> systems. Validated requirements for <span class="hl">test coverage</span>. Designed/maintained test cases, scenarios, test data.</li>
    <li>Led <span class="hl">defect management</span> — <span class="hl">triage</span>, prioritisation, <span class="hl">root-cause</span> analysis, retesting. Communicated test progress, <span class="hl">risk</span>s, <span class="hl">quality metrics</span> to <span class="hl">stakeholder</span>s. Supported <span class="hl">UAT</span> — guided business users, managed test cycles, ensured signoff readiness.</li>
    <li><span class="hl">Automation</span> &amp; continuous improvement — <span class="hl">Playwright</span>, <span class="hl">CI/CD</span> (GitHub Actions), <span class="hl">Selenium</span>-to-<span class="hl">Playwright</span> <span class="hl">migration</span> (3x faster, 50% CI reduction). Python, TypeScript. Grafana dashboards. AI-assisted — 30% velocity. Exceptional Performer.</li>
    <li>Team leadership — engineers + consultants. Coaching, mentoring. <span class="hl">Stakeholder</span> coordination (business, vendor, architects). Bridged business and technology. <span class="hl">Agile</span> (Scrum/SAFe).</li>
  </ul>

  <div class="job-title">Truecaller, Stockholm — Release &amp; Test Lead</div>
  <div class="job-sub">Sep 2021 – Feb 2022 | Enterprise Platform — 300M+ Users, Distributed Architecture</div>
  <ul>
    <li><span class="hl">Test strategy</span>, release readiness, <span class="hl">quality metrics</span> for distributed platform (300M+ users). <span class="hl">Entry/exit criteria</span>, <span class="hl">risk</span> management. <span class="hl">System testing</span>, <span class="hl">integration</span>, <span class="hl">regression</span>. <span class="hl">Defect management</span>. Go/no-go. <span class="hl">Stakeholder</span> communication. AWS.</li>
    <li>Release <span class="hl">automation</span>, <span class="hl">CI/CD</span>. Cross-functional coordination. Feature flags. <span class="hl">Agile</span> (Scrum/Kanban).</li>
  </ul>

  <div class="job-title">HCLTech — LEGO &amp; IKEA, Denmark &amp; Sweden — Test Lead / SDET Lead</div>
  <div class="job-sub">2013 – 2021 | Enterprise, E-Commerce &amp; PLM — Integrations, Multi-Partner, Upgrades</div>
  <ul>
    <li>LEGO &amp; IKEA (2017–21): End-to-end testing across enterprise platforms — complex <span class="hl">integration</span>s (<span class="hl">ERP</span>, APIs, third-party). <span class="hl">Test strategy</span>, plans, schedules. <span class="hl">System testing</span>, <span class="hl">integration</span>, <span class="hl">regression</span>, <span class="hl">UAT</span>. <span class="hl">Upgrade</span>s and <span class="hl">migration</span>s. Teams of 8–10. <span class="hl">Stakeholder</span> reporting. <span class="hl">Agile</span> &amp; <span class="hl">Waterfall</span>.</li>
    <li>SDET Lead (2013–17): <span class="hl">Automation</span> frameworks (<span class="hl">Selenium</span>, C#, NUnit). <span class="hl">CI/CD</span>. Full <span class="hl">SDLC</span> — <span class="hl">test plan</span>ning, execution, <span class="hl">defect management</span>, <span class="hl">quality metrics</span>. Tested enterprise configurations, <span class="hl">workflow</span>s, <span class="hl">access control</span>. <span class="hl">Regulated</span> environments. RFP/RFI. Mentoring.</li>
  </ul>

  <div class="job-title">Banking &amp; Enterprise — Test Lead / Consultant</div>
  <div class="job-sub">2008 – 2013 | Core Banking (Finacle CBS) — Migrations, Integrations, Regulated</div>
  <ul>
    <li>Enterprise banking (<span class="hl">Finacle</span> CBS) — <span class="hl">system testing</span>, <span class="hl">integration</span>, <span class="hl">regression</span>, <span class="hl">UAT</span>. Configurations, <span class="hl">workflow</span>s, <span class="hl">lifecycle</span> states, <span class="hl">access control</span>. Data <span class="hl">migration</span>. ETL <span class="hl">automation</span>. Hardware <span class="hl">integration</span>. <span class="hl">Defect management</span>. <span class="hl">Regulated</span>, audit-ready. <span class="hl">Stakeholder</span> coordination.</li>
  </ul>

  <div class="section">Key Achievements</div>
  <ul>
    <li><span class="hl">Test strategy</span> &amp; quality across enterprise platforms — IKEA (30+ markets), Truecaller (300M+), LEGO, banking. End-to-end across <span class="hl">integration</span>s and <span class="hl">upgrade</span>s.</li>
    <li>Enterprise platform testing — configurations, <span class="hl">workflow</span>s, <span class="hl">lifecycle</span>, <span class="hl">access control</span>, <span class="hl">BOM</span>-like structures, <span class="hl">Document Management</span>. <span class="hl">ERP</span>/<span class="hl">CAD</span>/<span class="hl">SaaS</span> <span class="hl">integration</span>s. <span class="hl">Upgrade</span>s, <span class="hl">cloud</span> <span class="hl">migration</span>s.</li>
    <li><span class="hl">Defect management</span> — <span class="hl">triage</span>, <span class="hl">root-cause</span>, <span class="hl">quality metrics</span>. <span class="hl">Stakeholder</span> communication. <span class="hl">UAT</span> signoff readiness. Go-live gates.</li>
    <li><span class="hl">Automation</span> — <span class="hl">Playwright</span>, <span class="hl">Selenium</span>, <span class="hl">CI/CD</span>. 3x faster, 50% CI reduction. Scaled IKEA platform 2K→50K. Exceptional Performer.</li>
    <li>Team leadership (4–15). Coaching, mentoring. Vendor/<span class="hl">stakeholder</span> coordination. Multi-partner. <span class="hl">Agile</span> &amp; <span class="hl">Waterfall</span>.</li>
  </ul>

  <div class="section">Education &amp; Certifications</div>
  <table>
    <tr><td class="tag">Education</td><td class="tag">Certifications</td></tr>
    <tr>
      <td>B.Tech in Information Technology<br>PG Diploma in Operations Management</td>
      <td>ISTQB Certified Tester Foundation<br>Six Sigma Green Belt<br>ITIL Foundation<br>GCP Associate Cloud Engineer<br>AWS Cloud Practitioner<br>Certified Ethical Hacker (CEH)<br>UiPath RPA Certified</td>
    </tr>
  </table>
</body>
</html>
"""
    DOC_PATH.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    build_docx()
    build_doc()
    print(DOCX_PATH)
    print(DOC_PATH)
