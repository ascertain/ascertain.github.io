from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


BASE_DIR = Path(r"C:/Users/MOKAS10/vcs/csrs-vcs-core")
DOCX_PATH = BASE_DIR / "Mohammad_Kashif_Senior_Test_Manager_Resume.docx"
DOC_PATH = BASE_DIR / "Mohammad_Kashif_Senior_Test_Manager_Resume.doc"

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
    "test strategy",
    "test manager",
    "test planning",
    "defect management",
    "integration",
    "API",
    "Kafka",
    "batch",
    "event stream",
    "SaaS",
    "billing",
    "invoicing",
    "financial",
    "banking",
    "end-to-end",
    "distributed",
    "stakeholder",
    "vendor",
    "quality",
    "reporting",
    "architecture",
    "subledger",
    "accounts receivable",
    "data exchange",
    "migration",
    "transformation",
    "Swedish",
    "Stockholm",
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
        "Senior Test Manager  |  Integration, SaaS Transformation & Financial Systems  |  15+ Years",
        bold=True,
        size=10.5,
    )

    contact = document.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    contact.paragraph_format.space_before = Pt(0)
    contact.paragraph_format.space_after = Pt(2)
    add_text(contact, "Malmö / Stockholm, Sweden | +46 702624230 | mo.kashif@gmail.com", size=9.5, color=TEXT_MUTED)
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
        "Senior Test Manager with 15+ years leading test strategy, planning, and execution "
        "across large-scale IT transformation programs. Deep experience managing testing of "
        "integration-heavy solutions — API, event streams (Kafka), batch files, and data "
        "exchange between SaaS platforms and in-house systems. Background in the financial "
        "sector — core banking (Finacle CBS), billing, invoicing, accounts reconciliation, "
        "and regulatory compliance. Owns end-to-end quality across distributed architectures "
        "combining SaaS and custom-built components. Drives test planning, execution, "
        "reporting, and defect management. Manages multiple stakeholders across business, "
        "technology, architects, and external SaaS vendors. Ensures alignment between internal "
        "teams and third-party providers. Structured, pragmatic, and delivery-focused. "
        "Fluent in Swedish and English.",
        size=10,
    )

    # ── How I Match the Role ──
    add_section_heading(document, "How I Match the Role")
    match_lines = [
        (
            "Senior Test Manager — Large IT Programs: ",
            "8+ years in test management/lead roles across large programs at IKEA (30+ "
            "markets, multi-partner), Truecaller (300M+ users), LEGO, and banking. Defined "
            "and implemented test strategies. Drove test planning, execution, reporting, "
            "and defect management. Managed teams of 4–15 across multiple workstreams.",
        ),
        (
            "Integration-Heavy Solutions (API, Kafka, Batch): ",
            "Extensive hands-on experience testing integration layers — REST APIs, event "
            "streams, batch file processing, and data exchange between systems. At IKEA: "
            "built and tested data pipelines (BigQuery, Cloud Functions, Pub/Sub), API "
            "integrations between in-house platform and external SaaS vendor, event-driven "
            "architectures. At banking: batch processing, ETL, data reconciliation.",
        ),
        (
            "End-to-End Testing in Distributed Architectures: ",
            "Owns end-to-end quality across distributed systems combining SaaS platforms, "
            "integration layers, and downstream systems. Test coverage from API contract "
            "testing through integration, system, E2E, and UAT. Experience with microservices, "
            "event-driven, and hybrid cloud/on-prem architectures.",
        ),
        (
            "SaaS Transformations: ",
            "At IKEA — managed testing for SaaS platform integration (VCS), coordinated "
            "with external SaaS vendor, aligned delivery between vendor roadmap and internal "
            "requirements. End-to-end validation of SaaS + in-house integration. Migration "
            "from on-prem to cloud-based systems.",
        ),
        (
            "Financial Sector & Billing: ",
            "Core banking experience (Finacle CBS) — billing, invoicing, accounts "
            "receivable/subledger, payment processing, financial reconciliation. "
            "Regulatory compliance and audit-ready documentation. Data migration between "
            "legacy and new platforms. Understanding of financial processes end-to-end.",
        ),
        (
            "Stakeholder Management & Vendor Alignment: ",
            "Manages multiple stakeholders across business, technology, and external "
            "vendors. At IKEA: bridged business leadership and SaaS vendor — managed "
            "service relationship, aligned delivery with business KPIs. Coordinated with "
            "architects, developers, and business teams. Fluent in Swedish and English.",
        ),
    ]
    for label, value in match_lines:
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.line_spacing = 1.0
        add_text(paragraph, label, bold=True, size=10)
        add_text(paragraph, value, size=10)

    # ── Core Competencies ──
    add_section_heading(document, "Core Competencies")
    skill_lines = [
        (
            "Test Management: ",
            "Test strategy definition & implementation. Test planning, execution, "
            "reporting. Defect management. Quality gates. Risk-based testing. "
            "UAT coordination. Go-live readiness. Audit-ready documentation",
        ),
        (
            "Integration Testing: ",
            "API testing (REST, GraphQL). Event streams (Kafka, Pub/Sub). "
            "Batch file validation. Data exchange & reconciliation. ETL testing. "
            "Contract testing. End-to-end across distributed architectures",
        ),
        (
            "Domain Expertise: ",
            "Financial systems — billing, invoicing, accounts receivable, subledger. "
            "Core banking (Finacle CBS). Payment processing. Data migration. "
            "SaaS transformation. Regulatory compliance",
        ),
        (
            "Stakeholder & Vendor: ",
            "Multi-stakeholder management (business, tech, architects, vendors). "
            "SaaS vendor alignment. Cross-team coordination. Reporting to "
            "program leadership. Swedish & English fluency",
        ),
        (
            "Automation & Tools: ",
            "Playwright, Selenium, Cypress. CI/CD (GitHub Actions, Jenkins). "
            "Python, TypeScript, C#, Java. Jira, Confluence, TestRail. "
            "Grafana (dashboards, reporting). Postman, REST-assured",
        ),
        (
            "Architecture & DevOps: ",
            "Microservices, event-driven, SaaS + on-prem hybrid. GCP, AWS. "
            "Docker, Kubernetes, Terraform. Data pipelines (BigQuery, "
            "Cloud Functions). Agile (Scrum/SAFe), Waterfall",
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
    add_text(cp1, "IKEA IT AB, Malmö — Test Manager / Team Lead", bold=True, size=10, color=SECTION_COLOR)
    tp1 = document.add_paragraph()
    tp1.paragraph_format.space_after = Pt(0)
    add_text(tp1, "Mar 2022 – Present  |  VCS Platform — SaaS Integration, Multi-Partner, 30+ Markets", bold=True, size=10)
    ikea_bullets = [
        "Test strategy & management — defined and implemented test strategy across the "
        "VCS program. Drove test planning, execution, reporting, and defect management. "
        "End-to-end quality ownership across integration layer, external SaaS platform, "
        "and downstream systems. Quality gates and go-live readiness.",
        "Integration testing — API integrations (REST) between in-house platform and "
        "external SaaS vendor. Event-driven data exchange (Pub/Sub). Data pipelines "
        "(BigQuery, Cloud Functions). Batch processing and data reconciliation. "
        "Validated end-to-end data flow across distributed architecture.",
        "SaaS transformation — managed testing for SaaS platform integration. Coordinated "
        "with external SaaS vendor on test alignment, release coordination, and defect "
        "resolution. Ensured quality across both SaaS and in-house components. Migrated "
        "from legacy approaches to cloud-native architecture.",
        "Stakeholder & vendor management — bridged IKEA business leadership and SaaS "
        "vendor. Managed service relationship, aligned delivery with business KPIs. "
        "Reported to program leadership. Multi-stakeholder coordination across business, "
        "technology, and architects.",
        "Team leadership & automation — led team of engineers and consultants. Coaching, "
        "mentoring, competence development. Automation: Playwright, CI/CD (GitHub Actions), "
        "Terraform, Docker on GCP. AI-assisted — 30% velocity improvement. Scaled platform "
        "from 2K to 50K usage across 30+ global markets. Exceptional Performer.",
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
    add_text(tp2, "Sep 2021 – Feb 2022  |  Communication Platform — 300M+ Users, Distributed Architecture", bold=True, size=10)
    tc_bullets = [
        "Test management at scale — owned test strategy, release readiness, and quality "
        "for a distributed platform serving 300M+ users. Go/no-go decisions. Defect "
        "management and reporting. API and integration testing across microservices. "
        "Event-driven architecture. AWS cloud infrastructure.",
        "Cross-functional coordination — managed stakeholders across engineering, product, "
        "and operations. Release automation and CI pipelines. Feature flag management "
        "and data-driven rollouts. Agile (Scrum/Kanban). Stockholm-based.",
    ]
    for b in tc_bullets:
        add_highlighted_bullet(document, b)

    # -- LEGO/IKEA via HCLTech --
    cp3 = document.add_paragraph()
    cp3.paragraph_format.space_before = Pt(3)
    cp3.paragraph_format.space_after = Pt(0)
    add_text(cp3, "HCLTech — LEGO & IKEA Group, Denmark & Sweden — Test Lead / Technical Specialist", bold=True, size=10, color=SECTION_COLOR)
    tp3 = document.add_paragraph()
    tp3.paragraph_format.space_after = Pt(0)
    add_text(tp3, "2013 – 2021  |  E-Commerce & Enterprise — Integration, Multi-Partner Programs", bold=True, size=10)
    lego_bullets = [
        "LEGO & IKEA (2017–2021): Test planning and execution across large multi-partner "
        "programs. API and integration testing. End-to-end validation across distributed "
        "systems. Managed test scope, risk, timelines, and reporting. Coordinated across "
        "multiple stakeholders (business, tech, vendors). Led teams of 8–10 engineers. "
        "Quality gates and release readiness.",
        "SDET Lead (2013–2017): Full SDLC test management. Designed automation frameworks "
        "(Selenium, C#). CI/CD integration. Test reporting for leadership. Managed "
        "defects and quality metrics. Cross-team coordination. Delivered across "
        "multiple concurrent workstreams.",
    ]
    for b in lego_bullets:
        add_highlighted_bullet(document, b)

    # -- Banking --
    cp4 = document.add_paragraph()
    cp4.paragraph_format.space_before = Pt(3)
    cp4.paragraph_format.space_after = Pt(0)
    add_text(cp4, "Banking & Financial Systems — Test Lead / Consultant", bold=True, size=10, color=SECTION_COLOR)
    tp4 = document.add_paragraph()
    tp4.paragraph_format.space_after = Pt(0)
    add_text(tp4, "2008 – 2013  |  Finacle CBS — Billing, Accounts, Data Migration, Regulatory", bold=True, size=10)
    fin_bullets = [
        "Financial sector testing — core banking systems (Finacle CBS): billing, invoicing, "
        "accounts receivable, subledger reconciliation. Tested payment processing workflows, "
        "transaction posting, and financial data integrity. Regulatory compliance and "
        "audit-ready documentation. Security testing (biometric authentication).",
        "Integration & data migration — batch file processing, ETL automation (Pentaho), "
        "data exchange between legacy and new systems. Data reconciliation and validation "
        "(SQL, UNIX). Post go-live stabilization. Client interface and defect management. "
        "Led testing workstreams end-to-end across multiple projects.",
    ]
    for b in fin_bullets:
        add_highlighted_bullet(document, b)

    # ── Key Achievements ──
    add_section_heading(document, "Key Achievements")
    achievements = [
        "Test strategy & quality ownership — defined and implemented test strategies "
        "across large programs at IKEA, Truecaller, LEGO. End-to-end quality across "
        "SaaS, integration layers, and downstream systems. Go-live readiness.",
        "Integration testing at scale — API, event streams, batch, data pipelines. "
        "Validated end-to-end data flow in distributed architectures. SaaS + in-house "
        "hybrid systems. 30+ markets, 300M+ users.",
        "Financial sector & billing — core banking (Finacle CBS): billing, invoicing, "
        "accounts receivable, subledger. Data migration, reconciliation, regulatory "
        "compliance. Audit-ready. 5+ years in financial systems.",
        "Stakeholder & vendor alignment — managed relationships across business, "
        "technology, architects, and external SaaS vendors. Bridged business and vendor. "
        "Multi-stakeholder reporting and coordination.",
        "Automation & delivery — Playwright, Selenium, CI/CD. Led Selenium-to-Playwright "
        "migration (3x faster, 50% CI reduction). Scaled IKEA VCS 2K→50K. Exceptional Performer.",
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

    # ── Languages ──
    add_section_heading(document, "Languages")
    lp = document.add_paragraph()
    lp.paragraph_format.space_after = Pt(0)
    add_text(lp, "Swedish — Fluent  |  English — Fluent  |  Hindi/Urdu — Native", size=10)

    document.save(DOCX_PATH)


def build_doc():
    html = """\
<html>
<head>
  <meta charset="utf-8">
  <title>Mohammad Kashif – Senior Test Manager Resume</title>
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
  <h2>Senior Test Manager | Integration, SaaS Transformation &amp; Financial Systems | 15+ Years</h2>
  <div class="contact">Malmö / Stockholm, Sweden | +46 702624230 | mo.kashif@gmail.com | LinkedIn: linkedin.com/in/md-kashif | Blog: ascertain.github.io</div>

  <div class="section">Profile</div>
  <p>Senior <span class="hl">Test Manager</span> with 15+ years leading <span class="hl">test strategy</span>, planning, and execution across large-scale IT <span class="hl">transformation</span> programs. Deep experience managing testing of <span class="hl">integration</span>-heavy solutions — <span class="hl">API</span>, <span class="hl">event stream</span>s (<span class="hl">Kafka</span>), <span class="hl">batch</span> files, and <span class="hl">data exchange</span> between <span class="hl">SaaS</span> platforms and in-house systems. Background in the <span class="hl">financial</span> sector — core <span class="hl">banking</span> (Finacle CBS), <span class="hl">billing</span>, <span class="hl">invoicing</span>, accounts reconciliation, and regulatory compliance. Owns <span class="hl">end-to-end</span> <span class="hl">quality</span> across <span class="hl">distributed</span> <span class="hl">architecture</span>s combining <span class="hl">SaaS</span> and custom-built components. Drives <span class="hl">test planning</span>, execution, <span class="hl">reporting</span>, and <span class="hl">defect management</span>. Manages multiple <span class="hl">stakeholder</span>s across business, technology, architects, and external <span class="hl">SaaS</span> <span class="hl">vendor</span>s. Fluent in <span class="hl">Swedish</span> and English.</p>

  <div class="section">How I Match the Role</div>
  <p><b>Senior Test Manager — Large IT Programs:</b> 8+ years in test management/lead roles at IKEA (30+ markets), Truecaller (300M+ users), LEGO, banking. Defined/implemented <span class="hl">test strategy</span>. Drove <span class="hl">test planning</span>, execution, <span class="hl">reporting</span>, <span class="hl">defect management</span>. Teams of 4–15.<br><br>
  <b>Integration-Heavy Solutions (API, Kafka, Batch):</b> REST <span class="hl">API</span>s, <span class="hl">event stream</span>s (Pub/Sub), <span class="hl">batch</span> processing, <span class="hl">data exchange</span>. At IKEA: data pipelines (BigQuery, Cloud Functions, Pub/Sub), <span class="hl">API</span> <span class="hl">integration</span>s between in-house and <span class="hl">SaaS</span> <span class="hl">vendor</span>. Banking: <span class="hl">batch</span>, ETL, data reconciliation.<br><br>
  <b>End-to-End in Distributed Architectures:</b> <span class="hl">End-to-end</span> <span class="hl">quality</span> across <span class="hl">distributed</span> systems — <span class="hl">SaaS</span>, <span class="hl">integration</span> layers, downstream. Contract testing → <span class="hl">integration</span> → system → E2E → UAT. Microservices, event-driven, hybrid cloud/on-prem.<br><br>
  <b>SaaS Transformations:</b> IKEA — <span class="hl">SaaS</span> platform <span class="hl">integration</span> (VCS). Coordinated with external <span class="hl">SaaS</span> <span class="hl">vendor</span>. Aligned <span class="hl">vendor</span> roadmap with internal requirements. <span class="hl">Migration</span> from on-prem to cloud-based.<br><br>
  <b>Financial Sector &amp; Billing:</b> Core <span class="hl">banking</span> (Finacle CBS) — <span class="hl">billing</span>, <span class="hl">invoicing</span>, <span class="hl">accounts receivable</span>/<span class="hl">subledger</span>, payment processing, <span class="hl">financial</span> reconciliation. Regulatory compliance. Data <span class="hl">migration</span>.<br><br>
  <b>Stakeholder &amp; Vendor Alignment:</b> Multiple <span class="hl">stakeholder</span>s (business, tech, <span class="hl">vendor</span>). Bridged business and <span class="hl">SaaS</span> <span class="hl">vendor</span>. Aligned delivery with KPIs. Coordinated architects, developers, business. <span class="hl">Swedish</span> &amp; English fluent.</p>

  <div class="section">Core Competencies</div>
  <p><b>Test Management:</b> Test strategy, planning, execution, reporting. Defect management. Quality gates. Risk-based. UAT. Go-live readiness. Audit-ready<br>
  <b>Integration Testing:</b> API (REST, GraphQL). Event streams (Kafka, Pub/Sub). Batch. Data exchange/reconciliation. ETL. Contract testing. E2E distributed<br>
  <b>Domain:</b> Financial — billing, invoicing, accounts receivable, subledger. Banking (Finacle). Payments. Data migration. SaaS transformation. Compliance<br>
  <b>Stakeholder &amp; Vendor:</b> Multi-stakeholder (business, tech, architects, vendors). SaaS vendor alignment. Reporting. Swedish &amp; English<br>
  <b>Automation &amp; Tools:</b> Playwright, Selenium, Cypress. CI/CD (GitHub Actions, Jenkins). Python, TypeScript, C#. Jira, Confluence, TestRail. Grafana<br>
  <b>Architecture:</b> Microservices, event-driven, SaaS+on-prem hybrid. GCP, AWS. Docker, Kubernetes, Terraform. Data pipelines. Agile (Scrum/SAFe)</p>

  <div class="section">Professional Experience</div>

  <div class="job-title">IKEA IT AB, Malmö — Test Manager / Team Lead</div>
  <div class="job-sub">Mar 2022 – Present | VCS Platform — SaaS Integration, Multi-Partner, 30+ Markets</div>
  <ul>
    <li><span class="hl">Test strategy</span> &amp; management — defined/implemented <span class="hl">test strategy</span>. <span class="hl">Test planning</span>, execution, <span class="hl">reporting</span>, <span class="hl">defect management</span>. <span class="hl">End-to-end</span> <span class="hl">quality</span> across <span class="hl">integration</span> layer, <span class="hl">SaaS</span>, and downstream systems. <span class="hl">Quality</span> gates, go-live readiness.</li>
    <li><span class="hl">Integration</span> testing — <span class="hl">API</span> (REST) between in-house and <span class="hl">SaaS</span> <span class="hl">vendor</span>. Event-driven <span class="hl">data exchange</span> (Pub/Sub). Data pipelines (BigQuery, Cloud Functions). <span class="hl">Batch</span> processing. <span class="hl">End-to-end</span> <span class="hl">distributed</span> <span class="hl">architecture</span>.</li>
    <li><span class="hl">SaaS</span> <span class="hl">transformation</span> — <span class="hl">SaaS</span> platform <span class="hl">integration</span>. Coordinated with external <span class="hl">vendor</span>. Release coordination, defect resolution. <span class="hl">Quality</span> across <span class="hl">SaaS</span> + in-house. Cloud-native <span class="hl">migration</span>.</li>
    <li><span class="hl">Stakeholder</span> &amp; <span class="hl">vendor</span> — bridged business and <span class="hl">SaaS</span> <span class="hl">vendor</span>. Managed service relationship. Aligned delivery with KPIs. Multi-<span class="hl">stakeholder</span> <span class="hl">reporting</span>.</li>
    <li>Team &amp; automation — led engineers + consultants. Playwright, CI/CD, Terraform, Docker, GCP. Scaled 2K → 50K. Exceptional Performer.</li>
  </ul>

  <div class="job-title">Truecaller, Stockholm — Release &amp; Test Lead</div>
  <div class="job-sub">Sep 2021 – Feb 2022 | Communication Platform — 300M+ Users, Distributed Architecture</div>
  <ul>
    <li><span class="hl">Test</span> management — <span class="hl">test strategy</span>, release readiness, <span class="hl">quality</span> for <span class="hl">distributed</span> platform (300M+ users). <span class="hl">Defect management</span>, <span class="hl">reporting</span>. <span class="hl">API</span> and <span class="hl">integration</span> testing across microservices. Event-driven. AWS. <span class="hl">Stockholm</span>.</li>
    <li><span class="hl">Stakeholder</span> coordination — engineering, product, operations. Release automation, CI pipelines. Feature flags, data-driven rollouts. Agile.</li>
  </ul>

  <div class="job-title">HCLTech — LEGO &amp; IKEA, Denmark &amp; Sweden — Test Lead / Technical Specialist</div>
  <div class="job-sub">2013 – 2021 | E-Commerce &amp; Enterprise — Integration, Multi-Partner Programs</div>
  <ul>
    <li>LEGO &amp; IKEA (2017–21): <span class="hl">Test planning</span>/execution across large multi-partner programs. <span class="hl">API</span> and <span class="hl">integration</span> testing. <span class="hl">End-to-end</span> across <span class="hl">distributed</span> systems. Scope, risk, <span class="hl">reporting</span>. Multiple <span class="hl">stakeholder</span>s. Teams of 8–10. <span class="hl">Quality</span> gates.</li>
    <li>SDET Lead (2013–17): Full SDLC test management. Automation frameworks (Selenium, C#). CI/CD. <span class="hl">Defect management</span>, <span class="hl">quality</span> metrics. Multi-workstream delivery.</li>
  </ul>

  <div class="job-title">Banking &amp; Financial Systems — Test Lead / Consultant</div>
  <div class="job-sub">2008 – 2013 | Finacle CBS — Billing, Accounts, Data Migration, Regulatory</div>
  <ul>
    <li><span class="hl">Financial</span> sector — core <span class="hl">banking</span> (Finacle CBS): <span class="hl">billing</span>, <span class="hl">invoicing</span>, <span class="hl">accounts receivable</span>, <span class="hl">subledger</span> reconciliation. Payment processing. <span class="hl">Financial</span> data integrity. Regulatory compliance. Security testing.</li>
    <li><span class="hl">Integration</span> &amp; <span class="hl">migration</span> — <span class="hl">batch</span> processing, ETL automation (Pentaho), <span class="hl">data exchange</span> between legacy and new systems. Data reconciliation (SQL, UNIX). Post go-live. <span class="hl">Defect management</span>. Led workstreams <span class="hl">end-to-end</span>.</li>
  </ul>

  <div class="section">Key Achievements</div>
  <ul>
    <li><span class="hl">Test strategy</span> &amp; <span class="hl">quality</span> — defined/implemented across IKEA, Truecaller, LEGO. <span class="hl">End-to-end</span> across <span class="hl">SaaS</span>, <span class="hl">integration</span>, downstream. Go-live readiness.</li>
    <li><span class="hl">Integration</span> at scale — <span class="hl">API</span>, <span class="hl">event stream</span>s, <span class="hl">batch</span>, data pipelines. <span class="hl">Distributed</span> <span class="hl">architecture</span>s. <span class="hl">SaaS</span> + in-house. 30+ markets, 300M+ users.</li>
    <li><span class="hl">Financial</span> &amp; <span class="hl">billing</span> — Finacle CBS: <span class="hl">billing</span>, <span class="hl">invoicing</span>, <span class="hl">accounts receivable</span>, <span class="hl">subledger</span>. Data <span class="hl">migration</span>, reconciliation. 5+ years.</li>
    <li><span class="hl">Stakeholder</span> &amp; <span class="hl">vendor</span> — business, tech, architects, <span class="hl">SaaS</span> <span class="hl">vendor</span>s. Bridged business and <span class="hl">vendor</span>. Multi-<span class="hl">stakeholder</span> <span class="hl">reporting</span>.</li>
    <li>Automation — Playwright, Selenium, CI/CD. 3x faster execution, 50% CI reduction. Scaled IKEA VCS 2K→50K. Exceptional Performer.</li>
  </ul>

  <div class="section">Education &amp; Certifications</div>
  <table>
    <tr><td class="tag">Education</td><td class="tag">Certifications</td></tr>
    <tr>
      <td>B.Tech in Information Technology<br>PG Diploma in Operations Management</td>
      <td>ISTQB Certified Tester Foundation<br>Six Sigma Green Belt<br>ITIL Foundation<br>GCP Associate Cloud Engineer<br>AWS Cloud Practitioner<br>Certified Ethical Hacker (CEH)<br>UiPath RPA Certified</td>
    </tr>
  </table>

  <div class="section">Languages</div>
  <p>Swedish — Fluent | English — Fluent | Hindi/Urdu — Native</p>
</body>
</html>
"""
    DOC_PATH.write_text(html, encoding="utf-8")


if __name__ == "__main__":
    build_docx()
    build_doc()
    print(DOCX_PATH)
    print(DOC_PATH)
