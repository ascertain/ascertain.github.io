import datetime
import html
import re
from pathlib import Path

import requests
from bs4 import BeautifulSoup

OUT_HTML = Path(__file__).resolve().parent / "Tally_Book_BY_PARNAMI.html"
LEGACY_OUT_HTML = Path(__file__).resolve().parent / "Tally_Student_Book_Internet.html"

SOURCES = {
    "TallyPrime Overview": "https://tallysolutions.com/",
    "Tally Help Home": "https://help.tallysolutions.com/",
    "Quick Start Guide": "https://help.tallysolutions.com/quick-start-guide/",
    "Accounting in TallyPrime": "https://help.tallysolutions.com/tally-prime/accounting/accounting-in-tally-prime/",
    "Inventory in TallyPrime": "https://help.tallysolutions.com/tally-prime/inventory/inventory-tally/",
    "Sales Process": "https://help.tallysolutions.com/tally-prime/sales-process-tally/sales-process-tally/",
    "Purchase Process": "https://help.tallysolutions.com/tally-prime/purchase-process-tally/purchase-process-tally/",
    "Tally Solutions (Background)": "https://en.wikipedia.org/wiki/Tally_Solutions",
    "SSC Study Tally Course Outline": "https://sscstudy.com/tally-prime-book-pdf-free-download/",
    "Mkuzak Tally Overview": "https://mkuzak.am/tally-prime-book-pdf-free-download/?lang=en",
}

ORDER = [
    "Syllabus",
    "Notes",
    "Tally Workbook Case Study - 1",
    "Tally Workbook Case Study - 2",
    "Tally Workbook Case Study - 3",
    "Tally Workbook Case Study - 4",
    "Tally Workbook Case Study - 5",
    "Tally Workbook Case Study - 6",
    "Tally.ERP  9 Keyboard Shortcuts",
    "Configuration",
]


def clean_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def slugify(text: str) -> str:
    s = text.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "-", s)
    s = re.sub(r"-+", "-", s)
    return s.strip("-")


def fetch_summary(url: str, limit: int = 1200) -> str:
    try:
        r = requests.get(url, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        for tag in soup(["script", "style", "noscript"]):
            tag.extract()

        title = clean_text(soup.title.get_text(" ")) if soup.title else "Untitled"
        paras = []
        for p in soup.find_all(["p", "li"], limit=80):
            t = clean_text(p.get_text(" "))
            if len(t) >= 50:
                paras.append(t)
            if sum(len(x) for x in paras) > limit:
                break

        body = " ".join(paras)
        body = body[:limit].rsplit(" ", 1)[0]
        return f"{title}. {body}" if body else title
    except Exception as exc:
        return f"Unavailable ({exc.__class__.__name__})"


def esc(s: str) -> str:
    return html.escape(s, quote=True)


def chapter_visuals(title: str) -> str:
    title = title.strip()

    if title == "Syllabus":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Tally Screen: First Day Orientation</h3>
        <div class="mock-screen">
            <div class="screen-title">Gateway of Tally - Student Orientation</div>
            <pre>
Top area:
    Active Company | Current Date | Current Period

Main zones:
    Masters         -> Create Ledger / Item / Group
    Transactions    -> Vouchers (daily entries)
    Reports         -> Day Book, Trial Balance, P&amp;L

Bottom buttons:
    F2 Date | F3 Company | Alt+G Go To | Ctrl+A Save

Rule:
    Create masters first, then post vouchers, then verify reports.
            </pre>
        </div>
    </div>
    <div class="visual-card">
        <h3>How Students Should Use Tally Daily</h3>
        <ol class="step-list">
            <li>Open company and verify date (F2).</li>
            <li>Create or review required ledgers/items.</li>
            <li>Enter vouchers in sequence by date.</li>
            <li>Check Day Book for correctness.</li>
            <li>Check Trial Balance for matched totals.</li>
            <li>Backup data at the end of session.</li>
        </ol>
    </div>
</div>
<div class="visual-card" style="margin-top:14px;">
    <h3>Syllabus Hyperlinks (Jump to Respective Pages)</h3>
    <ul class="link-list">
        <li><a href="#ch2-what-is-tally">What is Tally</a></li>
        <li><a href="#ch2-what-is-accounts">What is Accounts</a></li>
        <li><a href="#ch2-classification-of-accounts">Classification of Accounts</a></li>
        <li><a href="#ch2-terminology-of-accounting">Terminology of Accounting</a></li>
        <li><a href="#ch2-golden-rules-of-accounting">Golden Rules of Accounting</a></li>
        <li><a href="#ch2-general-entry">General Entry</a></li>
        <li><a href="#ch2-pre-define-tally-group">Pre-Define Tally Group</a></li>
        <li><a href="#ch2-types-of-voucher">Types of Voucher</a></li>
        <li><a href="#ch10-overview-of-tally-erp-9">Overview of Tally.ERP 9</a></li>
        <li><a href="#ch10-download-and-install">Download and Install</a></li>
        <li><a href="#ch10-create-company">Create Company</a></li>
        <li><a href="#ch10-configurations">Configurations</a></li>
        <li><a href="#ch10-basic-introduction-of-features">Basic Introduction of Features</a></li>
        <li><a href="#ch10-voucher-transaction-entry">Voucher Transaction Entry</a></li>
        <li><a href="#ch10-view-reports">View Reports</a></li>
        <li><a href="#ch3-accounting-case-study-1">Accounting (Case Study-1)</a></li>
        <li><a href="#ch3-financial-year-from-1-apr-2016-to-31-mar-2017">Financial Year from 1-Apr-2016 to 31-Mar-2017</a></li>
        <li><a href="#ch3-debit-credit-by-to">Debit, Credit - By, To</a></li>
        <li><a href="#ch3-create-ledger-under-group">Create Ledger (Under Group)</a></li>
        <li><a href="#ch3-create-group">Create Group</a></li>
        <li><a href="#ch3-service-organization-general-entry-12-months">Service Organization General Entry 12 Months</a></li>
        <li><a href="#ch3-day-book-trail-balance">Day Book, Trail Balance</a></li>
        <li><a href="#ch3-profit-loss-accounting-balance-sheet">Profit and Loss Accounting, Balance Sheet</a></li>
        <li><a href="#ch3-closing-opening-balance">Closing and Opening Balance</a></li>
        <li><a href="#ch3-exercise-with-solution">Exercise With Solution</a></li>
        <li><a href="#ch4-accounts-with-inventory-case-study-2">Accounts with Inventory (Case Study-2)</a></li>
        <li><a href="#ch4-financial-year-from-1-apr-2016-to-31-mar-2017">Financial year from 1-Apr-2016 to 31-Mar-2017</a></li>
        <li><a href="#ch4-create-stock-item-stock-group-unit-of-measure">Create - Stock Item, Stock Group, Unit of measure</a></li>
        <li><a href="#ch4-trading-organization-general-entry-12-months">Trading organization General Entry 12 Months</a></li>
        <li><a href="#ch4-purchase-sales-invoice">Purchase, Sales, Invoice</a></li>
        <li><a href="#ch4-provision-entry-depreciation-entry">Provision entry, Depreciation entry</a></li>
        <li><a href="#ch4-adjustment-entry-closing-entry">Adjustment entry, Closing entry</a></li>
        <li><a href="#ch4-inventory-reports-stock-summary">Inventory Reports, Stock Summary</a></li>
        <li><a href="#ch4-export-printing">Export & Printing</a></li>
        <li><a href="#ch4-closing-opening-balance">Closing & Opening Balance</a></li>
        <li><a href="#ch4-exercise-with-solution">Exercise With Solution</a></li>
        <li><a href="#ch5-inventory-management-with-gst-case-study-3">Inventory Management with GST (Case Study-3)</a></li>
        <li><a href="#ch5-financial-year-from-1-apr-2017-to-31-mar-2018">Financial year from 1-Apr-2017 to 31-Mar-2018</a></li>
        <li><a href="#ch5-taxation-system-gst">Taxation System GST</a></li>
        <li><a href="#ch5-sgst-cgst-igst">SGST, CGST, IGST</a></li>
        <li><a href="#ch5-create-godowns">Create - Godowns</a></li>
        <li><a href="#ch5-create-stock-item-stock-group-unit-of-measure">Create - Stock Item, Stock Group, Unit of measure</a></li>
        <li><a href="#ch5-create-cost-category-cost-centre">Create - Cost Category, Cost Centre</a></li>
        <li><a href="#ch5-trading-organization-general-entry-12-months">Trading organization General Entry 12 Months</a></li>
        <li><a href="#ch5-discounts-interest">Discounts, Interest</a></li>
        <li><a href="#ch5-purchase-returns-sales-returns">Purchase returns, Sales Returns</a></li>
        <li><a href="#ch5-credit-note-debit-note-voucher">Credit Note, Debit Note - Voucher</a></li>
        <li><a href="#ch5-inventory-reports-stock-summary-tax-report">Inventory Reports, Stock Summary, Tax Report</a></li>
        <li><a href="#ch5-closing-opening-balance">Closing & Opening Balance</a></li>
        <li><a href="#ch5-exercise-with-solution">Exercise With Solution</a></li>
        <li><a href="#ch8-advance-level">Advance Level</a></li>
        <li><a href="#ch8-service-organization-with-gst">Service Organization With GST</a></li>
        <li><a href="#ch8-manufacturing-with-gst">Manufacturing with GST</a></li>
        <li><a href="#ch8-payroll-management">Payroll Management</a></li>
        <li><a href="https://tallysolutions.com/download/" target="_blank" rel="noopener">Download and Install Tally (Official Site)</a></li>
        <li><a href="https://help.tallysolutions.com/" target="_blank" rel="noopener">Tally Help Documentation</a></li>
    </ul>
</div>
<div class="visual-grid" style="margin-top:12px;">
    <div class="visual-card">
        <h3>Classification of Accounts Chart</h3>
        <svg viewBox="0 0 760 280" role="img" aria-label="Classification of accounts chart">
            <rect x="280" y="20" width="200" height="48" rx="10" class="node"/>
            <text x="380" y="49" text-anchor="middle">Accounts</text>

            <rect x="50" y="120" width="190" height="56" rx="10" class="node"/>
            <text x="145" y="146" text-anchor="middle">Personal Account</text>
            <text x="145" y="166" text-anchor="middle">(person/entity)</text>

            <rect x="285" y="120" width="190" height="56" rx="10" class="node"/>
            <text x="380" y="146" text-anchor="middle">Real Account</text>
            <text x="380" y="166" text-anchor="middle">(assets)</text>

            <rect x="520" y="120" width="190" height="56" rx="10" class="node"/>
            <text x="615" y="146" text-anchor="middle">Nominal Account</text>
            <text x="615" y="166" text-anchor="middle">(income/expense)</text>

            <line x1="340" y1="68" x2="180" y2="120" class="arrow"/>
            <line x1="380" y1="68" x2="380" y2="120" class="arrow"/>
            <line x1="420" y1="68" x2="580" y2="120" class="arrow"/>

            <text x="60" y="232">Use this chart before deciding Dr/Cr in vouchers.</text>
        </svg>
    </div>
    <div class="visual-card">
        <h3>Golden Rules Quick Matrix</h3>
        <table class="check-table">
            <tr><th>Account Type</th><th>Debit Rule</th><th>Credit Rule</th></tr>
            <tr><td>Personal</td><td>Receiver</td><td>Giver</td></tr>
            <tr><td>Real</td><td>Comes In</td><td>Goes Out</td></tr>
            <tr><td>Nominal</td><td>Expense/Loss</td><td>Income/Gain</td></tr>
        </table>
        <p>Student tip: identify account type first, then apply rule, then choose voucher type.</p>
    </div>
</div>
<div class="visual-grid" style="margin-top:12px;">
    <div class="visual-card">
        <h3>Accounting Cycle Diagram</h3>
        <svg viewBox="0 0 760 250" role="img" aria-label="Accounting cycle">
            <rect x="20" y="90" width="120" height="50" rx="8" class="node"/><text x="80" y="120" text-anchor="middle">Source Doc</text>
            <rect x="160" y="90" width="120" height="50" rx="8" class="node"/><text x="220" y="120" text-anchor="middle">Voucher</text>
            <rect x="300" y="90" width="120" height="50" rx="8" class="node"/><text x="360" y="120" text-anchor="middle">Ledger</text>
            <rect x="440" y="90" width="120" height="50" rx="8" class="node"/><text x="500" y="120" text-anchor="middle">Trial Balance</text>
            <rect x="580" y="90" width="160" height="50" rx="8" class="node"/><text x="660" y="120" text-anchor="middle">P&amp;L + B/S</text>
            <line x1="140" y1="115" x2="160" y2="115" class="arrow"/>
            <line x1="280" y1="115" x2="300" y2="115" class="arrow"/>
            <line x1="420" y1="115" x2="440" y2="115" class="arrow"/>
            <line x1="560" y1="115" x2="580" y2="115" class="arrow"/>
            <text x="24" y="36">Entry rule: Evidence -> Voucher -> Posting -> Verification -> Reporting</text>
        </svg>
    </div>
    <div class="visual-card">
        <h3>Voucher Selection Matrix</h3>
        <table class="check-table">
            <tr><th>Situation</th><th>Voucher</th><th>Key Shortcut</th></tr>
            <tr><td>Cash to Bank transfer</td><td>Contra</td><td>F4</td></tr>
            <tr><td>Money paid</td><td>Payment</td><td>F5</td></tr>
            <tr><td>Money received</td><td>Receipt</td><td>F6</td></tr>
            <tr><td>Adjustment/Provision</td><td>Journal</td><td>F7</td></tr>
            <tr><td>Sales bill</td><td>Sales</td><td>F8</td></tr>
            <tr><td>Purchase bill</td><td>Purchase</td><td>F9</td></tr>
        </table>
    </div>
</div>
"""

    if title == "Notes":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Tally Screen: Daily Entry Flow</h3>
        <div class="mock-screen">
            <div class="screen-title">Gateway of Tally - Company: Practice Co.</div>
            <pre>
Display More Reports
Books of Accounts
Statement of Accounts
Inventory Books

Vouchers [Press: V]
    F4 Contra    F5 Payment    F6 Receipt
    F7 Journal   F8 Sales      F9 Purchase

Quick path:
Alt+G -> Type: Vouchers -> Enter
            </pre>
        </div>
    </div>
    <div class="visual-card">
        <h3>Voucher Decision Chart</h3>
        <svg viewBox="0 0 760 260" role="img" aria-label="Voucher decision chart">
            <rect x="20" y="20" width="170" height="42" rx="8" class="node"/>
            <text x="105" y="47" text-anchor="middle">Cash/Bank Transfer</text>

            <rect x="220" y="20" width="130" height="42" rx="8" class="node"/>
            <text x="285" y="47" text-anchor="middle">Contra</text>

            <rect x="20" y="92" width="170" height="42" rx="8" class="node"/>
            <text x="105" y="119" text-anchor="middle">Money Paid Out</text>

            <rect x="220" y="92" width="130" height="42" rx="8" class="node"/>
            <text x="285" y="119" text-anchor="middle">Payment</text>

            <rect x="20" y="164" width="170" height="42" rx="8" class="node"/>
            <text x="105" y="191" text-anchor="middle">Money Received</text>

            <rect x="220" y="164" width="130" height="42" rx="8" class="node"/>
            <text x="285" y="191" text-anchor="middle">Receipt</text>

            <line x1="190" y1="41" x2="220" y2="41" class="arrow"/>
            <line x1="190" y1="113" x2="220" y2="113" class="arrow"/>
            <line x1="190" y1="185" x2="220" y2="185" class="arrow"/>

            <rect x="400" y="20" width="160" height="42" rx="8" class="node"/>
            <text x="480" y="47" text-anchor="middle">Goods Sold</text>
            <rect x="590" y="20" width="130" height="42" rx="8" class="node"/>
            <text x="655" y="47" text-anchor="middle">Sales</text>
            <line x1="560" y1="41" x2="590" y2="41" class="arrow"/>

            <rect x="400" y="92" width="160" height="42" rx="8" class="node"/>
            <text x="480" y="119" text-anchor="middle">Goods Purchased</text>
            <rect x="590" y="92" width="130" height="42" rx="8" class="node"/>
            <text x="655" y="119" text-anchor="middle">Purchase</text>
            <line x1="560" y1="113" x2="590" y2="113" class="arrow"/>

            <rect x="400" y="164" width="160" height="42" rx="8" class="node"/>
            <text x="480" y="191" text-anchor="middle">Adjustment Entry</text>
            <rect x="590" y="164" width="130" height="42" rx="8" class="node"/>
            <text x="655" y="191" text-anchor="middle">Journal</text>
            <line x1="560" y1="185" x2="590" y2="185" class="arrow"/>
        </svg>
    </div>
</div>
<div class="visual-grid" style="margin-top:12px;">
    <div class="visual-card">
        <h3>Case 1 Month Plan Chart</h3>
        <div class="mini-bars">
            <div><label>Apr-Jun Setup</label><span style="--w:95%"></span></div>
            <div><label>Jul-Sep Billing</label><span style="--w:88%"></span></div>
            <div><label>Oct-Dec Controls</label><span style="--w:82%"></span></div>
            <div><label>Jan-Mar Finalization</label><span style="--w:76%"></span></div>
        </div>
        <p>Practice each quarter with at least 20 voucher entries and one report review session.</p>
    </div>
    <div class="visual-card">
        <h3>Case 1 Report Check Grid</h3>
        <table class="check-table">
            <tr><th>Report</th><th>What to Verify</th><th>Action if Error</th></tr>
            <tr><td>Day Book</td><td>Date, voucher type, narration</td><td>Alter voucher immediately</td></tr>
            <tr><td>Trial Balance</td><td>Debit equals Credit</td><td>Find wrong Dr/Cr line</td></tr>
            <tr><td>P&amp;L</td><td>Income and expense grouping</td><td>Check ledger group mapping</td></tr>
            <tr><td>Balance Sheet</td><td>Asset/liability correctness</td><td>Review opening and closing balances</td></tr>
        </table>
    </div>
</div>
"""

    if title == "Tally Workbook Case Study - 1":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Case 1 Entry Sequence (Student Flow)</h3>
        <div class="mock-screen">
            <div class="screen-title">Gateway - Alt+G quick navigation</div>
            <pre>
1) Create Company
2) Alter -> Enable GST + Bill-wise details
3) Create Ledgers
4) Open Vouchers (V)
5) Pass Contra -> Sales -> Receipt -> Payment -> Journal
6) Open Reports: Day Book, Trial Balance, P&amp;L, Balance Sheet

Shortcut rhythm:
F2 Date -> F4/F6/F8/F5/F7 -> Ctrl+A save
            </pre>
        </div>
    </div>
    <div class="visual-card">
        <h3>Validation Check Diagram</h3>
        <svg viewBox="0 0 700 260" role="img" aria-label="Case study validation flow">
            <rect x="20" y="30" width="130" height="44" rx="8" class="node"/>
            <text x="85" y="57" text-anchor="middle">Vouchers</text>

            <rect x="190" y="30" width="130" height="44" rx="8" class="node"/>
            <text x="255" y="57" text-anchor="middle">Day Book</text>

            <rect x="360" y="30" width="130" height="44" rx="8" class="node"/>
            <text x="425" y="57" text-anchor="middle">Trial Balance</text>

            <rect x="530" y="30" width="150" height="44" rx="8" class="node"/>
            <text x="605" y="57" text-anchor="middle">P&amp;L / B&amp;S</text>

            <line x1="150" y1="52" x2="190" y2="52" class="arrow"/>
            <line x1="320" y1="52" x2="360" y2="52" class="arrow"/>
            <line x1="490" y1="52" x2="530" y2="52" class="arrow"/>

            <rect x="120" y="130" width="220" height="92" rx="10" class="node"/>
            <text x="230" y="155" text-anchor="middle">If mismatch:</text>
            <text x="230" y="180" text-anchor="middle">Check date, ledger, Dr/Cr side,</text>
            <text x="230" y="205" text-anchor="middle">GST ledger and voucher type</text>

            <rect x="390" y="130" width="250" height="92" rx="10" class="node"/>
            <text x="515" y="155" text-anchor="middle">If matched:</text>
            <text x="515" y="180" text-anchor="middle">Export reports and submit</text>
            <text x="515" y="205" text-anchor="middle">screenshots with narration</text>
        </svg>
    </div>
</div>
<div class="visual-grid" style="margin-top:12px;">
    <div class="visual-card">
        <h3>Navigation Map: From Startup to Reports</h3>
        <svg viewBox="0 0 760 250" role="img" aria-label="Tally navigation map">
            <rect x="20" y="90" width="140" height="50" rx="8" class="node"/><text x="90" y="120" text-anchor="middle">Open Company</text>
            <rect x="180" y="90" width="140" height="50" rx="8" class="node"/><text x="250" y="120" text-anchor="middle">F11/F12 Setup</text>
            <rect x="340" y="90" width="140" height="50" rx="8" class="node"/><text x="410" y="120" text-anchor="middle">Create Masters</text>
            <rect x="500" y="90" width="120" height="50" rx="8" class="node"/><text x="560" y="120" text-anchor="middle">Vouchers</text>
            <rect x="640" y="90" width="100" height="50" rx="8" class="node"/><text x="690" y="120" text-anchor="middle">Reports</text>
            <line x1="160" y1="115" x2="180" y2="115" class="arrow"/>
            <line x1="320" y1="115" x2="340" y2="115" class="arrow"/>
            <line x1="480" y1="115" x2="500" y2="115" class="arrow"/>
            <line x1="620" y1="115" x2="640" y2="115" class="arrow"/>
        </svg>
    </div>
    <div class="visual-card">
        <h3>Beginner Action Checklist</h3>
        <table class="check-table">
            <tr><th>Task</th><th>Target Time</th><th>Student Check</th></tr>
            <tr><td>Create Company</td><td>5 min</td><td>FY and date correct</td></tr>
            <tr><td>Create Core Ledgers</td><td>10 min</td><td>Group mapping correct</td></tr>
            <tr><td>Pass 5 Vouchers</td><td>15 min</td><td>Dr/Cr accurate</td></tr>
            <tr><td>Check Day Book</td><td>5 min</td><td>All entries visible</td></tr>
            <tr><td>Check Trial Balance</td><td>5 min</td><td>Totals match</td></tr>
        </table>
    </div>
</div>
"""

    if title == "Tally Workbook Case Study - 2":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Tally Screen: Inventory Voucher Entry</h3>
        <div class="mock-screen">
            <div class="screen-title">Gateway -> Vouchers -> Inventory Allocation</div>
            <pre>
Purchase (F9)
Party: Bright Supply Co.
Item: LED Bulb 12W
Qty: 100 Nos   Rate: 500   Amount: 50,000
Add: Freight Inward 2,000

Sales (F8)
Party: City Retail
Item: LED Bulb 12W
Qty: 35 Nos    Rate: 750   Amount: 26,250

Return flow:
Ctrl+F9 Debit Note (Purchase Return)
Ctrl+F8 Credit Note (Sales Return)
            </pre>
        </div>
    </div>
    <div class="visual-card">
        <h3>Stock Movement Chart (Case 2)</h3>
        <svg viewBox="0 0 720 300" role="img" aria-label="Case 2 stock movement chart">
            <rect x="50" y="180" width="70" height="80" class="node"/>
            <text x="85" y="172" text-anchor="middle">Opening 20</text>

            <rect x="170" y="60" width="70" height="200" class="node"/>
            <text x="205" y="52" text-anchor="middle">Purchase +100</text>

            <rect x="290" y="120" width="70" height="140" class="node"/>
            <text x="325" y="112" text-anchor="middle">Sales -35</text>

            <rect x="410" y="150" width="70" height="110" class="node"/>
            <text x="445" y="142" text-anchor="middle">P.Return -5</text>

            <rect x="530" y="140" width="70" height="120" class="node"/>
            <text x="565" y="132" text-anchor="middle">S.Return +2</text>

            <rect x="620" y="96" width="70" height="164" class="node"/>
            <text x="655" y="88" text-anchor="middle">Closing 82</text>

            <line x1="120" y1="220" x2="170" y2="220" class="arrow"/>
            <line x1="240" y1="220" x2="290" y2="220" class="arrow"/>
            <line x1="360" y1="220" x2="410" y2="220" class="arrow"/>
            <line x1="480" y1="220" x2="530" y2="220" class="arrow"/>
            <line x1="600" y1="220" x2="620" y2="220" class="arrow"/>

            <line x1="40" y1="260" x2="700" y2="260" class="axis"/>
        </svg>
    </div>
</div>
"""

    if title == "Tally Workbook Case Study - 3":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Tally Screen: GST Voucher Pattern</h3>
        <div class="mock-screen">
            <div class="screen-title">Sales Voucher - Tax Classification</div>
            <pre>
Local Sales (Intra-state)
  Sales Value: 140,000
  Output CGST 9%: 12,600
  Output SGST 9%: 12,600

Outside State Sales (Inter-state)
  Sales Value: 80,000
  Output IGST 18%: 14,400

Quick check:
If buyer state == company state -> CGST+SGST
Else -> IGST
            </pre>
        </div>
    </div>
    <div class="visual-card">
        <h3>GST Flow Diagram</h3>
        <svg viewBox="0 0 760 300" role="img" aria-label="GST input output flow">
            <rect x="40" y="40" width="240" height="70" rx="10" class="node"/>
            <text x="160" y="70" text-anchor="middle">Input Tax (Purchases)</text>
            <text x="160" y="95" text-anchor="middle">CGST 9,000 + SGST 9,000</text>

            <rect x="40" y="170" width="240" height="80" rx="10" class="node"/>
            <text x="160" y="200" text-anchor="middle">Output Tax (Sales)</text>
            <text x="160" y="225" text-anchor="middle">CGST 12,600 + SGST 12,600</text>

            <rect x="330" y="105" width="170" height="80" rx="10" class="node"/>
            <text x="415" y="137" text-anchor="middle">Net Local GST</text>
            <text x="415" y="162" text-anchor="middle">3,600 + 3,600</text>

            <rect x="550" y="105" width="170" height="80" rx="10" class="node"/>
            <text x="635" y="137" text-anchor="middle">IGST Payable</text>
            <text x="635" y="162" text-anchor="middle">14,400</text>

            <line x1="280" y1="75" x2="330" y2="130" class="arrow"/>
            <line x1="280" y1="210" x2="330" y2="160" class="arrow"/>
            <line x1="500" y1="145" x2="550" y2="145" class="arrow"/>
        </svg>
    </div>
</div>
"""

    if title == "Tally Workbook Case Study - 4":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Tally Screen: Bill-wise Receipt Adjustment</h3>
        <div class="mock-screen">
            <div class="screen-title">Receipt Voucher (F6) - Against References</div>
            <pre>
Party: MNO Services

Pending Bills:
  INV-A    60,000    Due: 17-Apr
  INV-B    40,000    Due: 05-May

Receipt Entry 14-Apr:
  Bank A/c Dr 35,000
  MNO Services Cr 35,000
  Against: INV-A (Part)

Settlement 28-Apr:
  Bank A/c Dr 38,000
  Discount Allowed Dr 2,000
  MNO Services Cr 40,000
  Against: INV-B (Full)
            </pre>
        </div>
    </div>
    <div class="visual-card">
        <h3>Receivables Ageing Snapshot</h3>
        <div class="mini-bars">
            <div><label>Current</label><span style="--w:20%"></span></div>
            <div><label>1-15 Days</label><span style="--w:45%"></span></div>
            <div><label>16-30 Days</label><span style="--w:65%"></span></div>
            <div><label>31-45 Days</label><span style="--w:30%"></span></div>
            <div><label>46+ Days</label><span style="--w:12%"></span></div>
        </div>
        <p>Use this view to prioritize follow-up calls: highest bucket first, then oldest invoice first.</p>
    </div>
</div>
"""

    if title == "Tally Workbook Case Study - 5":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Case 5 Workflow (Practice Path)</h3>
        <ol class="step-list">
            <li>Create masters: ledger, stock group, stock items, units.</li>
            <li>Enter purchase vouchers with tax and transport values.</li>
            <li>Enter sales vouchers and apply bill-wise references.</li>
            <li>Post returns using Credit Note and Debit Note.</li>
            <li>Run stock summary, ageing, and GST summary checks.</li>
        </ol>
    </div>
    <div class="visual-card">
        <h3>Practice Progress Chart</h3>
        <div class="mini-bars">
            <div><label>Master Setup</label><span style="--w:95%"></span></div>
            <div><label>Purchase Cycle</label><span style="--w:88%"></span></div>
            <div><label>Sales Cycle</label><span style="--w:90%"></span></div>
            <div><label>Returns</label><span style="--w:75%"></span></div>
            <div><label>Reconciliation</label><span style="--w:70%"></span></div>
        </div>
    </div>
</div>
"""

    if title == "Tally Workbook Case Study - 6":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Year-End Checklist Board</h3>
        <table class="check-table">
            <tr><th>Task</th><th>Status</th><th>Why It Matters</th></tr>
            <tr><td>Depreciation Entry</td><td>Required</td><td>Correct fixed asset value and P&amp;L impact</td></tr>
            <tr><td>Outstanding Expenses</td><td>Required</td><td>Match expenses to period</td></tr>
            <tr><td>GST Review</td><td>Required</td><td>Tax compliance and reconciliation</td></tr>
            <tr><td>Ledger Scrutiny</td><td>Required</td><td>Remove posting errors before closure</td></tr>
            <tr><td>Backup &amp; Lock Period</td><td>Required</td><td>Audit safety and data integrity</td></tr>
        </table>
    </div>
    <div class="visual-card">
        <h3>MIS View for Students</h3>
        <svg viewBox="0 0 640 260" role="img" aria-label="MIS trend chart">
            <polyline points="40,200 120,170 200,160 280,140 360,135 440,120 520,110 600,96" class="line-a"/>
            <polyline points="40,220 120,210 200,205 280,188 360,178 440,172 520,164 600,158" class="line-b"/>
            <line x1="40" y1="230" x2="610" y2="230" class="axis"/>
            <line x1="40" y1="40" x2="40" y2="230" class="axis"/>
            <text x="50" y="36">Monthly Trend: Gross Margin (A) vs Expense Ratio (B)</text>
        </svg>
    </div>
</div>
"""

    if title == "Configuration":
        return """
<div class="visual-grid">
    <div class="visual-card">
        <h3>Tally Screen: Configuration Navigation</h3>
        <div class="mock-screen">
            <div class="screen-title">Gateway -> F11 Features -> Set/Alter</div>
            <pre>
F11 Features
    Accounting Features
        Maintain bill-wise details: Yes
        Cost centres: As needed

    Statutory &amp; Taxation
        Enable GST: Yes
        Set GST details: Yes

F12 Configure
    Show Narration: Yes
    Use common narration: Optional
            </pre>
        </div>
    </div>
    <div class="visual-card">
        <h3>Workbook Follow Diagram</h3>
        <ol class="step-list">
            <li>Read chapter concept.</li>
            <li>Open Tally and configure prerequisites.</li>
            <li>Enter sample transactions.</li>
            <li>Validate with report checkpoints.</li>
            <li>Repeat with your own dataset.</li>
        </ol>
    </div>
</div>
"""

    return ""


def section_html(title: str, body: str, idx: int) -> str:
    sid = f"chapter-{idx}"
    lines = [x.strip() for x in body.strip().split("\n") if x.strip()]

    out = [
        f'<section id="{sid}" class="chapter">',
        f'<div class="chapter-meta">Chapter {idx}</div>',
        f"<h2>{esc(title)}</h2>",
    ]

    list_open = False
    for ln in lines:
        if re.match(r"^(\d+\.|-|\*)\s+", ln):
            if not list_open:
                out.append("<ul>")
                list_open = True
            out.append(f"<li>{esc(re.sub(r'^(\\d+\\.|-|\\*)\\s+', '', ln))}</li>")
        else:
            if list_open:
                out.append("</ul>")
                list_open = False
            if ln.endswith(":") and len(ln) < 120:
                htxt = ln[:-1].strip()
                hid = f"ch{idx}-{slugify(htxt)}"
                out.append(f'<h3 id="{esc(hid)}">{esc(ln)}</h3>')
            else:
                out.append(f"<p>{esc(ln)}</p>")

    if list_open:
        out.append("</ul>")

    visuals = chapter_visuals(title)
    if visuals:
        out.append(visuals)

    out.append("</section>")
    return "\n".join(out)


def build_chapters(source_blurbs: dict[str, str]) -> dict[str, str]:
    syllabus = """
Workbook Syllabus (Tally.ERP 9 Foundation to Advance)

Part A - Basics
1. What is Tally.
2. What is Accounts.
3. Classification of Accounts.
4. Terminology of Accounting.
5. Golden Rules of Accounting.
6. General Entry.
7. Pre-Defined Tally Groups.
8. Types of Voucher.

Part B - Overview of Tally.ERP 9
1. Download and Install.
2. Create Company.
3. Configurations.
4. Basic Introduction of Features.
5. Voucher Transaction Entry.
6. View Reports.

Part C - Accounting (Case Study 1)
1. Financial Year: 1-Apr-2016 to 31-Mar-2017.
2. Debit, Credit, By and To concept.
3. Voucher usage.
4. Create Ledger (Under Group).
5. Create Group.
6. Service organization general entries for 12 months.
7. Day Book and Trial Balance.
8. Profit and Loss Account and Balance Sheet.
9. Closing and Opening Balance.
10. Exercise with Solution.

Part D - Accounts with Inventory (Case Study 2)
1. Financial Year: 1-Apr-2016 to 31-Mar-2017.
2. Create Stock Item, Stock Group, Unit of Measure.
3. Trading organization general entries for 12 months.
4. Purchase, Sales, and Invoice.
5. Provision Entry and Depreciation Entry.
6. Adjustment Entry and Closing Entry.
7. Inventory Reports and Stock Summary.
8. Export and Printing.
9. Closing and Opening Balance.
10. Exercise with Solution.

Part E - Inventory Management with GST (Case Study 3)
1. Financial Year: 1-Apr-2017 to 31-Mar-2018.
2. Taxation system GST.
3. SGST, CGST, IGST.
4. Create Godowns.
5. Create Stock Item, Stock Group, Unit of Measure.
6. Create Cost Category and Cost Centre.
7. Trading organization general entries for 12 months.
8. Discounts and Interest.
9. Purchase Returns and Sales Returns.
10. Credit Note and Debit Note Voucher.
11. Inventory Reports, Stock Summary, and Tax Report.
12. Closing and Opening Balance.
13. Exercise with Solution.

Part F - Advance Level
1. Service Organization with GST.
2. Manufacturing with GST.
3. Payroll Management.

Tally.ERP 9 Full Tutorial Reference:
Website: www.upcissyoutube.com
"""

    notes = """
What is Tally:
Tally is business management and accounting software used by SMEs for bookkeeping, invoicing, inventory, taxation, reporting, and compliance.
Short Summary:
- Tally is a complete accounting workflow tool from entry to reports.

What is Accounts:
Accounts is the systematic process of recording, classifying, summarizing, and reporting financial transactions of a business.
Short Summary:
- Accounts answers what happened, why it happened, and what is the current financial position.

Classification of Accounts:
1. Personal Accounts:
- Natural person, artificial person, and representative accounts.
2. Real Accounts:
- Tangible and intangible assets.
3. Nominal Accounts:
- Expenses, losses, incomes, and gains.
Short Summary:
- Every transaction must touch at least two accounts from these classifications.

Terminology of Accounting:
1. Capital: owner investment in business.
2. Drawings: amount withdrawn by owner for personal use.
3. Assets: resources owned by business.
4. Liabilities: obligations payable by business.
5. Revenue: income from operations.
6. Expense: cost incurred to earn revenue.
7. Debtor: customer who must pay business.
8. Creditor: supplier to whom business must pay.
9. Voucher: documentary evidence of transaction entry.
10. Ledger: classified book of final entry.
Short Summary:
- Without correct terminology, voucher posting errors increase.

Core concepts:
- Ledger: individual account head such as Cash, Sales, Purchase, Salary.
- Group: category that controls report behavior.
- Voucher: transaction entry format (Contra, Payment, Receipt, Journal, Sales, Purchase).
- Stock Item and Unit: base entities for inventory accounting.

Golden Rules of Accounting:
1. Personal Account: Debit receiver, Credit giver.
2. Real Account: Debit what comes in, Credit what goes out.
3. Nominal Account: Debit expenses/losses, Credit incomes/gains.
Short Summary:
- Golden rules are the fastest way to decide Dr/Cr in manual checks.

General Entry:
1. Cash introduced by owner:
- Cash A/c Dr
- To Capital A/c
2. Rent paid by bank:
- Rent A/c Dr
- To Bank A/c
3. Credit sales to customer:
- Customer A/c Dr
- To Sales/Service Income A/c
Example set for practice:
1. Salary paid 20,000 by bank.
- Salary A/c Dr 20,000
- To Bank A/c 20,000
2. Received from customer 15,000 in cash.
- Cash A/c Dr 15,000
- To Customer A/c 15,000

Pre-Define Tally Group:
1. Capital Account
2. Current Assets
3. Current Liabilities
4. Sundry Debtors
5. Sundry Creditors
6. Direct Incomes / Indirect Incomes
7. Direct Expenses / Indirect Expenses
8. Bank Accounts / Cash-in-Hand
Short Summary:
- Correct group selection controls where values appear in reports.

Types of Voucher:
1. Contra (F4): cash-bank transfer.
2. Payment (F5): money paid out.
3. Receipt (F6): money received.
4. Journal (F7): adjustment/non-cash entries.
5. Sales (F8): sales invoice entries.
6. Purchase (F9): purchase invoice entries.
7. Credit Note: sales return.
8. Debit Note: purchase return.
Short Summary:
- Choose voucher type first, then post amount lines.

Internet reference highlights:
- """ + source_blurbs.get("Accounting in TallyPrime", "")[:700] + """

New Student Onboarding: How to use Tally from zero
1. First launch and company selection:
- Open TallyPrime from desktop/start menu.
- If no company exists, choose Create Company.
- If company exists, choose Select Company.
- Always verify date (F2) before posting entries.

2. Understand main screen (Gateway of Tally):
- Masters: where you create ledgers, groups, stock items, and units.
- Transactions: where you post vouchers (purchase, sales, payment, receipt).
- Reports: where you check Day Book, Trial Balance, P&L, and Balance Sheet.
- Bottom button bar: shows function key actions available on current screen.

3. Keyboard-first usage (fast and accurate):
- F2 change date.
- F3 company menu.
- Alt+G go to any report/screen.
- F4 Contra, F5 Payment, F6 Receipt, F7 Journal, F8 Sales, F9 Purchase.
- Ctrl+A save.

4. First 30-minute student practice plan:
Step A - Create company.
Step B - Create 5 ledgers: Cash, Bank, Capital, Sales, Rent.
Step C - Pass 4 vouchers:
- Receipt for capital introduced.
- Contra for cash deposit to bank.
- Sales entry for invoice.
- Payment entry for expense.
Step D - Open reports and verify balances.

5. Screen discipline students must follow:
- Read voucher type title before entering amounts.
- Confirm Dr/Cr side every line.
- Add narration for every voucher.
- Save only after checking party name, date, amount, and tax lines.

6. Common beginner confusion and fix:
- Cannot see expected field: press F12 Configure and enable detailed mode.
- Wrong voucher screen opened: use Esc and press correct function key.
- Entry saved with wrong date: Alter voucher immediately and correct date.
- Report mismatch: open Day Book and inspect voucher one by one.

7. Minimum daily routine for new students:
- 10 minutes: create/alter masters.
- 20 minutes: pass vouchers with keyboard only.
- 10 minutes: verify Day Book and Trial Balance.
- 5 minutes: write 3 mistakes and corrections in notebook.

Accounting Cycle Detailed Summary:
1. Collect source document (invoice, receipt, payment proof).
2. Choose correct voucher type in Tally.
3. Pass entry with proper debit/credit and narration.
4. Post impacts to ledger and verify day book.
5. Check trial balance equality.
6. Interpret P&L and Balance Sheet for decisions.

Student Error-Control Checklist:
1. Is date correct (F2)?
2. Is voucher type correct?
3. Is amount posted on correct Dr/Cr side?
4. Is ledger group correct?
5. Is narration present?
6. Does report match expected business result?

Supplementary Learning Themes from Public Course Pages:
1. Many public Tally courses split learning into fundamentals, inventory, GST, payroll, and security modules.
2. Students usually learn faster when theory notes are paired with bill-book style practice entries.
3. A strong learning order is: concept, masters, sample voucher, report check, then repeat with variation.
4. Public course outlines also emphasize short assignments, printable examples, and repeated keyboard practice.
"""

    case1 = """
Accounting (Case Study-1):
Financial Year from 1-Apr-2016 to 31-Mar-2017:

Coverage required in this case:
1. Debit, Credit, By, and To logic.
2. Voucher posting process.
3. Create Group and Create Ledger under proper group.
4. Service organization general entries for all 12 months.
5. Day Book, Trial Balance, Profit and Loss Account, and Balance Sheet.
6. Closing and Opening Balance.
7. Exercise with Solution.

Create Group:
Step 1 - Create group and ledgers:
1. Create Group (if required): Direct Expenses, Indirect Expenses, Direct Income.
Create Ledger (Under Group):
2. Create ledgers under group:
- Capital A/c (Capital Account)
- Cash A/c (Cash-in-Hand)
- Bank A/c (Bank Accounts)
- Service Income (Direct Income)
- Rent Expense (Indirect Expense)
- Salary Expense (Indirect Expense)
- Electricity Expense (Indirect Expense)
- Customer A/c (Sundry Debtors)

Debit, Credit - By, To:
Step 2 - Understand By and To in voucher:
1. In payment-type style entry, expense is Dr and cash/bank is Cr.
2. In receipt-type style entry, cash/bank is Dr and source ledger is Cr.
3. For service invoice on credit, customer is Dr and service income is Cr.

Service organization General Entry 12 Months:
Step 3 - Monthly entry model (repeat for 12 months):
Month starter entries (example April):
1. Capital introduced: 400,000.
2. Cash deposited to bank: 250,000.
3. Service invoice to customer: 60,000.
4. Office rent paid: 18,000.
5. Salary paid: 25,000.
6. Electricity paid: 4,500.
7. Partial receipt from customer: 30,000.

Practice rule:
1. Post same pattern for each month with changed invoice amount.
2. Keep narration clear for every voucher.
3. Keep date sequence strict to avoid report mismatch.

Voucher:
Step 4 - Voucher types used:
1. Receipt (F6) for capital and customer collection.
2. Contra (F4) for cash-to-bank transfer.
3. Sales (F8) for service bill.
4. Payment (F5) for expenses.
5. Journal (F7) for adjustment entries if needed.

Day Book ,Trail Balance:
Step 5 - Report checks every month:
1. Day Book: all transactions visible with correct date.
2. Trial Balance: Debit total equals Credit total.
Profit & Loss accounting, Balance Sheet:
3. Profit and Loss: income and expenses posted in right heads.
4. Balance Sheet: cash, bank, receivables, and capital position valid.

Closing & Opening Balance:
Step 6 - Closing and opening balance flow:
1. At year-end (31-Mar-2017), finalize books after checking all vouchers.
2. Carry forward closing balances as opening balances of next year.
3. Validate opening trial balance in next year before new entries.

Exercise With Solution:
Exercise with solution (sample):
Question:
1. Service bill raised 75,000 on credit.
2. Received 45,000 by bank.
3. Paid rent 20,000 and salary 26,000.

Expected posting logic:
1. Customer Dr 75,000 to Service Income Cr 75,000.
2. Bank Dr 45,000 to Customer Cr 45,000.
3. Rent Dr 20,000 to Bank Cr 20,000.
4. Salary Dr 26,000 to Bank Cr 26,000.

Submission checklist:
1. Group and ledger creation screenshots.
2. One month Day Book screenshot.
3. Trial Balance screenshot.
4. Profit and Loss and Balance Sheet screenshot.
5. One solved exercise screenshot with narration.

12-Month Example Plan (quick reference):
1. Apr-Jun:
- Focus on setup, basic service invoices, and expense posting.
2. Jul-Sep:
- Add debtor collections and bill-wise follow-up.
3. Oct-Dec:
- Add adjustments, provisions, and correction entries.
4. Jan-Mar:
- Finalization, closing checks, and opening balance carry-forward.

Deep Summary for Students:
1. Case Study 1 trains complete accounting flow in service business.
2. If Day Book is correct but reports are wrong, group mapping is usually wrong.
3. If Trial Balance mismatches, voucher line-side error is likely.
4. Always verify before next month to avoid year-end correction burden.
"""

    case2 = """
Accounts with Inventory (Case Study-2):
Financial year from 1-Apr-2016 to 31-Mar-2017:

Create - Stock Item, Stock Group, Unit of measure:
1. Create unit of measure such as Nos, Box, Kg if required.
2. Create stock groups to classify items by category.
3. Create stock items with opening quantity and opening rate.
4. Map items carefully so stock reports remain meaningful.
Short Summary:
- This setup controls how inventory appears in invoices and stock summary.

Trading organization General Entry 12 Months:
1. Use a monthly cycle of purchase, sales, collection, payment, return, and adjustment.
2. Keep one realistic business story for the full year.
3. For each month, check both accounting impact and stock impact.
Example monthly cycle:
1. Purchase goods from supplier.
2. Sell goods to customer.
3. Receive payment and pay freight.
4. Adjust damages, shortage, or return if any.

Purchase, Sales, Invoice:
1. Purchase invoice increases stock and creditor balance.
2. Sales invoice reduces stock and increases customer/debtor balance.
3. Invoice entry must include item, quantity, rate, and party.
4. Narration should explain the business event.

Provision entry, Depreciation entry:
1. Provision entry records expected expense/liability before final payment.
2. Depreciation entry records fixed-asset usage cost.
3. These entries are generally passed through Journal voucher.
Example:
1. Depreciation A/c Dr 2,500
2. To Furniture A/c 2,500

Adjustment entry, Closing entry:
1. Record outstanding expense adjustments.
2. Record prepaid expense adjustments if applicable.
3. Perform closing stock check and year-end transfer validation.
4. Ensure all return vouchers are posted before closing reports.

Inventory Reports, Stock Summary:
1. Stock Summary verifies quantity and value item-wise.
2. Inventory books help students trace movement history.
3. If stock quantity is wrong, check purchase/sales/return sequence first.

Export & Printing:
1. Export Stock Summary, Trial Balance, and invoice reports to PDF or Excel.
2. Print one purchase invoice and one sales invoice for workbook submission.
3. Use export preview to confirm layout before final save.

Closing & Opening Balance:
1. Finalize balances as on 31-Mar-2017.
2. Carry forward closing stock and closing ledger balances to next year.
3. Verify opening values before posting new-year vouchers.

Exercise With Solution:
Question:
1. Purchase 120 units at 500.
2. Sales 70 units at 780.
3. Purchase return 10 units.
4. Sales return 5 units.
5. Pass depreciation 2,500.

Solution:
1. Net stock movement = Opening + Purchase - Sales - Purchase Return + Sales Return.
2. Stock Summary must match the computed closing quantity.
3. Purchase, Sales, Return, and Depreciation must reflect in reports.
4. Trial Balance should remain matched after all entries.

Deep Student Guidance:
1. Every inventory voucher changes both quantity and value behavior.
2. Wrong unit or wrong stock item creates invisible report mistakes.
3. Students should check stock after every major purchase/sales entry block.

Supplementary Inventory Topics:
1. Batch-wise inventory:
- Useful for medicine, expiry-based goods, and packed items.
- Students should understand when batch, manufacturing date, or expiry date affects stock control.
2. Bill-book practice entries:
- Repeated purchase and sales bill formats improve accuracy in invoice posting.
3. Stock Journal and material transfer:
- Internal movement and stock adjustment can be practiced separately from normal purchase/sales flow.
4. Zero-value and actual-vs-billed quantity:
- These cases help students understand quantity adjustments that do not always behave like normal sales values.
5. Cost centre, order processing, and BOM:
- These are natural extensions after the basic inventory chapter is understood.

Practice Add-on Exercises:
1. Create one item with assumed batch tracking.
2. Record one stock transfer between two internal locations.
3. Create one invoice with quantity difference explanation.
4. Compare Stock Summary before and after adjustment.

Further Reading Links:
1. https://sscstudy.com/tally-prime-book-pdf-free-download/
2. https://help.tallysolutions.com/tally-prime/inventory/inventory-tally/

Submission checklist:
1. Item, group, and unit creation screenshots.
2. Purchase and sales invoice screenshots.
3. Stock Summary screenshot.
4. Depreciation/provision journal screenshot.
5. Exported report sample.
"""

    case3 = """
Inventory Management with GST (Case Study-3):
Financial year from 1-Apr-2017 to 31-Mar-2018:

Taxation System GST:
1. GST is applied through correct tax setup, state configuration, and voucher posting.
2. Students must understand taxable value, tax amount, and final invoice value.
3. GST affects both reports and legal compliance checks.

SGST, CGST, IGST:
1. SGST and CGST are used for intra-state transactions.
2. IGST is used for inter-state transactions.
3. Always confirm customer or supplier state before invoice posting.
Short Summary:
- Wrong tax type is one of the most common student mistakes in GST chapters.

Create - Godowns:
1. Create Main, Transit, and Returns godowns.
2. Use godowns to track where stock is physically located.
3. Godown-wise reporting helps in internal control.

Create - Stock Item, Stock Group, Unit of measure:
1. Create groups for product classes.
2. Create items with GST rate, UOM, and opening stock.
3. Use correct rate and quantity to avoid report mismatch.

Create - Cost Category, Cost Centre:
1. Use cost categories when comparing multiple business dimensions.
2. Use cost centres for branch, department, or project tracking.
3. This helps advanced students analyze profitability beyond basic books.

Trading organization General Entry 12 Months:
1. Run one full year of purchase, sales, return, discount, and interest entries.
2. Keep local and inter-state transactions mixed to practice tax decisions.
3. Reconcile month-end stock and GST before the next month starts.

Discounts, Interest:
1. Record trade discount as per invoice treatment in workbook method.
2. Record interest on overdue receivable/payable through journal or invoice-linked flow.
3. Check whether discount changes taxable value as per workbook assumption.

Purchase returns, Sales Returns:
1. Purchase return reverses stock inward and supplier liability.
2. Sales return reverses stock outward and customer billing.
3. Always use proper return voucher type and party reference.

Credit Note, Debit Note - Voucher:
1. Credit Note is generally used for sales return.
2. Debit Note is generally used for purchase return.
3. These vouchers must reflect both tax and stock corrections where applicable.

Inventory Reports, Stock Summary, Tax Report:
1. Stock Summary validates physical and book quantity.
2. Inventory reports show item movement details.
3. Tax report validates GST amounts and taxable turnover.
4. Students should compare tax ledgers with GST summary monthly.

Closing & Opening Balance:
1. Close books on 31-Mar-2018 after report validation.
2. Carry forward stock, tax, and ledger balances correctly.
3. Verify next-year opening values before new vouchers.

Exercise With Solution:
Question:
1. Local purchase 150,000 with 18% GST.
2. Inter-state sales 110,000 with 18% IGST.
3. Sales return 10,000 and purchase return 8,000.
4. Interest charged to debtor 1,500.

Solution:
1. Apply CGST+SGST to local purchase and IGST to inter-state sales.
2. Post return vouchers through Credit Note and Debit Note.
3. Recheck stock and tax impact after returns.
4. Confirm GST summary and stock summary reflect corrected values.

Deep Student Guidance:
1. This case teaches tax logic, stock control, and error correction together.
2. Students should not move to reports before checking voucher tax lines.
3. Godown and cost-centre setup become useful only when entries are posted consistently.

Supplementary GST Practice Topics:
1. GST ledger discipline:
- Separate purchase, sales, input tax, output tax, and party ledgers carefully.
2. Bill book with GST:
- Ready-made purchase and sales bill formats help students practice tax fields faster.
3. Debit Note and Credit Note learning:
- These should be practiced with both quantity reversal and tax effect awareness.
4. Interest calculation on parties:
- Overdue-interest scenarios are a useful extension after base GST learning.
5. Multiple currencies:
- Advanced students can later test foreign-currency purchase and sale behavior.
6. Import/export and security:
- XML/Excel export, backup, and access control are real business skills beyond entry posting.

Practice Add-on Exercises:
1. Create one local sale and one inter-state sale on the same date.
2. Pass one sales return and verify GST reversal.
3. Add one interest adjustment and check report effect.
4. Export a GST-related report to PDF.

Further Reading Links:
1. https://sscstudy.com/tally-prime-book-pdf-free-download/
2. https://help.tallysolutions.com/tally-prime/accounting/accounting-in-tally-prime/
3. https://help.tallysolutions.com/tally-prime/inventory/inventory-tally/

Submission checklist:
1. GST configuration screenshots.
2. Godown and cost centre setup screenshots.
3. Sales and purchase voucher screenshots with tax.
4. Credit Note and Debit Note screenshots.
5. Stock Summary and Tax Report screenshots.
"""

    case4 = """
Scenario: Service company receivables and follow-up management.
Objective:
- Track dues customer-wise with bill references.
- Record partial payments and settlement discounts.
- Use ageing report for collection priority.

Practice dataset:
- Invoice A: 60,000 on 02-Apr, due in 15 days.
- Invoice B: 40,000 on 05-Apr, due in 30 days.
- Receipt against Invoice A: 35,000 on 14-Apr.
- Receipt against Invoice B: 38,000 with 2,000 discount on 28-Apr.

Step 0 - Configuration:
1. Enable bill-wise details in F11.
2. Ensure customer ledger is set to maintain balances bill-by-bill.

Step 1 - Create ledgers:
1. Customer: MNO Services (Sundry Debtor, bill-wise Yes).
2. Service Income.
3. Bank A/c.
4. Discount Allowed.

Step 2 - Pass invoices:
Transaction 1 (02-Apr): Invoice A.
1. Sales voucher (F8).
2. Debit MNO Services 60,000.
3. Credit Service Income 60,000.
4. Bill reference: INV-A, due date +15 days.

Transaction 2 (05-Apr): Invoice B.
1. Sales voucher (F8).
2. Debit MNO Services 40,000.
3. Credit Service Income 40,000.
4. Bill reference: INV-B, due date +30 days.

Step 3 - Pass receipts and discount:
Transaction 3 (14-Apr): Partial receipt for INV-A.
1. Receipt voucher (F6).
2. Debit Bank A/c 35,000.
3. Credit MNO Services 35,000.
4. Link against INV-A.

Transaction 4 (28-Apr): Settlement of INV-B with discount.
1. Receipt voucher (F6) or Journal + Receipt based on institute method.
2. Debit Bank A/c 38,000.
3. Debit Discount Allowed 2,000.
4. Credit MNO Services 40,000.
5. Link against INV-B as full settlement.

Step 4 - Outstanding and ageing validation:
1. INV-A pending should be 25,000 (60,000 - 35,000).
2. INV-B should be fully closed.
3. Open Bills Receivable report:
- Check bill references and due dates.
4. Open Ageing report:
- Confirm pending amount appears in correct ageing bucket.

Step 5 - Management interpretation:
1. If multiple pending bills are overdue, prioritize oldest first.
2. Compare discount policy vs faster collections.
3. Track monthly DSO trend if asked in workbook.

Common mistakes and fixes:
1. Receipt not linked to bill reference.
- Alter voucher and mark against correct invoice.

2. Discount posted as expense but bill not closed.
- Ensure receipt entry includes both bank and discount against same bill.

3. Wrong due date in invoice.
- Alter bill credit period and recheck ageing bucket.

Submission checklist:
1. Invoice A and Invoice B screenshots.
2. Receipt entries screenshot.
3. Bill-wise outstanding report screenshot.
4. Ageing analysis screenshot.
5. Short collection strategy note (5 lines).

Practice extension:
1. Add one new customer invoice and one overdue follow-up receipt.
2. Re-run ageing and identify top overdue customer.
3. Draft a collection action list for next week.
"""

    case5 = """
Detailed Practice Workbook (Advanced Trading + Controls).
Objective:
- Execute realistic high-volume monthly cycle.
- Combine inventory, receivables, payables, tax, and bank controls.
- Produce management-ready reports with cross-validation.

Business story:
A wholesaler handles electronics inventory with monthly purchase cycles, credit sales, returns, and bank reconciliation.

Dataset blueprint:
1. Opening stock:
- 4 SKUs x 50 units each.

2. Purchases:
- Weekly purchase bills from 3 suppliers.
- Include transport and occasional purchase discount.

3. Sales:
- Credit sales to 10 customers.
- Different credit periods (15/30/45 days).

4. Returns and adjustments:
- 2 damaged stock returns from customers.
- 1 supplier replacement entry.

Step 0 - Master readiness:
1. Create stock groups:
- Mobile Accessories, Audio, Networking, Office Devices.
2. Create 12 stock items with UQC and GST mapping.
3. Create godowns: Main and Transit.
4. Create party ledgers with bill-wise details.

Step 1 - Purchase cycle execution:
1. Pass all weekly purchase vouchers.
2. Allocate items to godown and include freight where applicable.
3. Ensure GST auto-calculation on each bill.

Step 2 - Sales cycle execution:
1. Pass customer-wise sales vouchers.
2. Apply pricing and discount policy from workbook.
3. Maintain bill references for every credit invoice.

Step 3 - Returns and replacement:
1. Customer returns via Credit Notes with item quantities.
2. Supplier-side return/replacement via Debit Note and receipt entry.
3. Confirm stock movement is balanced across godowns.

Step 4 - Cash and bank discipline:
1. Enter receipts/payments against references.
2. Perform bank reconciliation at week-end and month-end.
3. Identify unreconciled lines and add notes.

Step 5 - Control checks:
1. Stock movement vs invoice movement must align item-wise.
2. Supplier outstanding must match pending purchase references.
3. Debtor ageing must match due-date logic.
4. GST summary totals must match tax ledgers.

Expected analytical outputs:
1. Trial Balance (matched).
2. Stock Summary by godown with high/low movers.
3. Debtors ageing with top overdue 5 customers.
4. GST monthly summary.
5. Bank reconciliation statement.

Common mistakes and fixes:
1. Stock negative due to wrong date sequence.
- Reorder voucher dates and repost in chronological order.

2. Wrong godown selected.
- Alter voucher inventory allocation and revalidate stock summary.

3. Returns without reference linking.
- Link to original bill for correct ageing and profitability impact.

Submission checklist:
1. Masters report (items, ledgers, godowns).
2. Purchase/sales sample vouchers.
3. Return voucher set.
4. Stock Summary and GST Summary screenshots.
5. Reconciliation note with 3 observations.

Practice extension:
1. Introduce one slow-moving item and apply discount campaign.
2. Measure impact on sales and margin.
3. Present before/after insights in 8 lines.
"""

    case6 = """
Advance Level:

Service Organization With GST:
1. Practice taxable service billing with debtor tracking and GST ledgers.
2. Combine service invoice, receipt, adjustment, and GST reconciliation.
3. Check service income, output tax, and receivable ageing together.

Manufacturing with GST:
1. Track raw material purchase, production usage, and finished-goods output.
2. Understand how stock movement and tax interact in a manufacturing case.
3. Validate consumption, production, and sales impact through reports.

Payroll Management:
1. Create salary ledgers and employee-related payable ledgers.
2. Record gross salary, deductions, and net salary payment.
3. Review payroll cost impact in Profit and Loss.
4. Typical training modules also include employee groups, pay heads, attendance, and payroll vouchers.

Supplementary Advance Practice Topics:
1. Security of data and import/export:
- Learn export to PDF/Excel/XML and safe backup habits.
- Understand security control and restricted user access.
2. Payroll structure examples:
- Basic Pay, HRA, Transport Allowance, Bonus, PF, ESIC, and overtime are common training items.
3. Manufacturing preparation:
- BOM, raw material issue, finished goods, and cost tracking are common next-step topics.
4. Assignment-based learning:
- Public training outlines often pair each module with a practice transaction set.

Detailed Practice Workbook (Year-end Closure + MIS Review).
Objective:
- Perform year-end accounting adjustments with confidence.
- Prepare audit-supporting books and control evidence.
- Generate MIS insights from finalized books.

Advance Level Mapping:
1. Service Organization with GST:
- Mixed taxable services, receivable tracking, and GST reconciliation.
2. Manufacturing with GST:
- Raw material purchase, production consumption, finished-goods sale, and tax impact.
3. Payroll Management:
- Salary components, payable ledgers, payment voucher, and payroll-cost reporting.

Business story:
A distributor wants clean year-end books, category profitability, and audit-ready trails.

Year dataset scope:
1. Full 12-month transactions:
- Purchase, sales, direct/indirect expenses, collections, payments.
2. Recurring adjustments:
- Monthly depreciation.
- Provision and accrual entries.
3. Closure activities:
- Final tax review.
- Carry-forward checks.

Step 0 - Pre-close preparation:
1. Freeze transaction cut-off date.
2. Take data backup snapshot.
3. Export trial balance draft for comparison.

Step 1 - Pass year-end adjustments:
1. Depreciation:
- Debit Depreciation Expense, credit accumulated depreciation/fixed asset adjustment ledger.
2. Outstanding expenses:
- Debit relevant expense, credit outstanding liability.
3. Prepaid expenses:
- Credit expense and debit prepaid asset where applicable.
4. Provision entries:
- Create provision ledgers per workbook requirement.

Step 2 - Verify compliance blocks:
1. Review GST ledgers and summary alignment.
2. Confirm no unreconciled statutory entries.
3. Check voucher narration quality for audit readability.

Step 3 - Audit pack creation:
1. Export Day Book, Ledger extracts, Trial Balance, P&L, Balance Sheet.
2. Export party outstanding and ageing reports.
3. Export stock summary and valuation detail.

Step 4 - Period control and governance:
1. Lock previous period after validation.
2. Define user rights for post-close changes.
3. Maintain exception register for any reopening.

Step 5 - MIS preparation:
1. Monthly gross margin trend.
2. Top customer contribution analysis.
3. Expense ratio and variance analysis.
4. Category-wise profitability if item/account grouping exists.

Practice challenge (error detection lab):
1. Identify 5 intentionally incorrect entries:
- Wrong expense head.
- Tax side swapped.
- Missing bill reference.
- Wrong period date.
- Reversed debit-credit line.
2. Correct all entries.
3. Explain effect on P&L, Balance Sheet, and tax output.

Common mistakes and fixes:
1. Adjustments posted in next period.
- Re-date to year-end and relock period.

2. Depreciation posted to wrong group.
- Move ledger to proper expense/asset mapping.

3. Provision not reversed next cycle.
- Add reversal schedule in opening month.

Submission checklist:
1. Pre-close vs post-close Trial Balance comparison.
2. Adjustment voucher bundle screenshot/PDF.
3. Final P&L and Balance Sheet screenshot.
4. GST validation screenshot.
5. MIS summary page with three key insights.

Practice extension:
1. Add one what-if scenario (sales down 10%, expenses up 5%).
2. Recompute expected margin pressure.
3. Suggest two management actions based on MIS.
"""

    shortcuts = """
Keyboard Shortcuts (Student Core Set):
- F2: Change Date
- F3: Company
- F4: Contra
- F5: Payment
- F6: Receipt
- F7: Journal
- F8: Sales
- F9: Purchase
- Ctrl+A: Accept/Save
- Alt+G: Go To
- Alt+K: Company menu actions

Tip:
Practice by entering one full day of vouchers using only keyboard.
"""

    config = """
Overview of Tally.ERP 9:
This section teaches how to open software, configure it, and post first transactions correctly.

Download and Install:
1. Open official site and download supported installer.
2. Run setup and complete installation wizard.
3. Start Tally from desktop icon.
4. Verify version in help/about screen.
Short Summary:
- Install from trusted source and validate version before training.

Create Company:
1. Choose Create Company from startup screen.
2. Enter company name, address, state, and currency.
3. Set financial year and books beginning date.
4. Save company and confirm on top bar.
Short Summary:
- Company creation defines your accounting period and base identity.

Configurations:
1. Use F11 for feature-level configuration.
2. Use F12 for screen behavior configuration.
3. Set narration, voucher numbering, and optional settings.
4. Configure backup location and security basics.
Short Summary:
- Correct configuration reduces data-entry errors significantly.

Basic Introduction of Features:
1. Accounting Features.
2. Inventory Features.
3. Statutory and Taxation Features (GST).
4. Cost centres, bill-wise details, and controls.
Short Summary:
- Enable only features needed for the current workbook case.

Voucher Transaction Entry:
1. Open Vouchers screen from Gateway.
2. Set date with F2.
3. Choose voucher type (F4/F5/F6/F7/F8/F9).
4. Enter Dr/Cr lines carefully.
5. Add narration and save with Ctrl+A.
Example:
1. Rent paid from bank 18,000.
- Rent A/c Dr 18,000
- To Bank A/c 18,000
Short Summary:
- Voucher type first, then amounts, then verification, then save.

View Reports:
1. Day Book for daily verification.
2. Trial Balance for Dr/Cr equality.
3. Profit and Loss for performance.
4. Balance Sheet for financial position.
Short Summary:
- Enter, verify, analyze is the mandatory student loop.

New Student Screen Walkthrough (must practice once):
1. Open company:
- Gateway -> Company -> Select Company.
- Confirm active company name on top bar.

2. Create masters screen path:
- Alt+G -> type Create Ledger.
- Choose group correctly before saving.

3. Open voucher screen path:
- Gateway -> Vouchers (or press V).
- Use function keys to switch voucher type.

4. Enter first transaction safely:
- Press F2 and set date.
- Select voucher type.
- Enter debit and credit lines.
- Add narration and press Ctrl+A.

5. Check report immediately:
- Alt+G -> Day Book.
- Open entered voucher and re-check Dr/Cr lines.
- Then open Trial Balance.

6. Exit and data safety:
- Take backup after practice session.
- Exit company from company menu.
- Do not shut PC before save confirmation.

New student do and do not:
- Do use consistent ledger naming.
- Do pass entries in date sequence.
- Do verify reports daily.
- Do not post without knowing voucher type.
- Do not mix personal and business entries.
- Do not skip narration.
"""

    return {
        "Syllabus": syllabus,
        "Notes": notes,
        "Tally Workbook Case Study - 1": case1,
        "Tally Workbook Case Study - 2": case2,
        "Tally Workbook Case Study - 3": case3,
        "Tally Workbook Case Study - 4": case4,
        "Tally Workbook Case Study - 5": case5,
        "Tally Workbook Case Study - 6": case6,
        "Tally.ERP  9 Keyboard Shortcuts": shortcuts,
        "Configuration": config,
    }


def build_html(chapters: dict[str, str], source_blurbs: dict[str, str]) -> str:
    toc = []
    body = []

    for i, title in enumerate(ORDER, start=1):
        toc.append(f'<li><a href="#chapter-{i}"><span>{i:02d}</span> {esc(title)}</a></li>')
        body.append(section_html(title, chapters.get(title, "Content unavailable."), i))

    src_items = []
    for k, v in SOURCES.items():
        src_items.append(f'<li><a href="{esc(v)}" target="_blank" rel="noopener">{esc(k)}</a></li>')

    generated_on = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Tally_Book_BY_PARNAMI</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Merriweather:wght@400;700;900&family=Manrope:wght@500;700&display=swap" rel="stylesheet">
<style>
:root {{
  --bg:#f4efe6;
  --paper:#fffdf8;
  --ink:#1f1a14;
  --muted:#6e5f4e;
  --accent:#0d7a66;
  --line:#dfd1be;
}}
*{{box-sizing:border-box}}
body{{margin:0;background:radial-gradient(circle at 0 0,#fff8ec,transparent 35%),var(--bg);color:var(--ink);font:400 17px/1.72 Merriweather,serif;}}
.shell{{max-width:1200px;margin:22px auto;padding:0 16px;display:grid;grid-template-columns:320px 1fr;gap:18px}}
.toc{{position:sticky;top:16px;background:#fff9ee;border:1px solid var(--line);border-radius:16px;padding:18px;max-height:calc(100vh - 32px);overflow:auto}}
.toc h1{{font:900 1.45rem/1.2 Merriweather,serif;margin:0 0 6px}}
.toc p{{margin:0 0 14px;color:var(--muted);font:600 .82rem/1.35 Manrope,sans-serif;letter-spacing:.06em;text-transform:uppercase}}
.toc ul{{list-style:none;padding:0;margin:0}}
.toc li+li{{margin-top:7px}}
.toc a{{display:block;text-decoration:none;color:var(--ink);padding:8px 10px;border-radius:10px;border:1px solid transparent}}
.toc a:hover{{background:#fff;border-color:var(--line)}}
.toc a span{{color:var(--accent);font:700 .88rem Manrope,sans-serif;margin-right:8px}}
.book{{background:var(--paper);border:1px solid var(--line);border-radius:18px;overflow:hidden}}
.cover{{padding:46px 50px;background:linear-gradient(135deg,#e9f5f1,#fff7ea);border-bottom:1px solid var(--line)}}
.badge{{display:inline-block;background:#d9efe9;color:#0c6a58;border:1px solid #b9ddd4;padding:6px 10px;border-radius:999px;font:700 .74rem Manrope,sans-serif;letter-spacing:.06em;text-transform:uppercase}}
.cover h1{{margin:14px 0 8px;font:900 clamp(2rem,4vw,3rem)/1.07 Merriweather,serif}}
.cover p{{margin:0;color:var(--muted)}}
.meta{{margin-top:14px;font:600 .85rem Manrope,sans-serif;color:#5d5348}}
.chart{{margin-top:18px;padding:14px;background:#fff;border:1px solid var(--line);border-radius:12px}}
.chart h3{{margin:0 0 10px;font:700 .96rem Manrope,sans-serif}}
.bar{{display:flex;align-items:center;gap:10px;margin:7px 0}}
.bar label{{width:160px;font:600 .82rem Manrope,sans-serif;color:#51463a}}
.bar .track{{flex:1;height:12px;background:#efe5d4;border-radius:999px;overflow:hidden}}
.bar .fill{{height:100%;background:linear-gradient(90deg,#0d7a66,#3bb097)}}
.chapter{{padding:30px 50px;border-top:1px solid var(--line)}}
.chapter-meta{{font:700 .72rem Manrope,sans-serif;letter-spacing:.08em;color:#0f6f5d;text-transform:uppercase}}
.chapter h2{{margin:.35rem 0 1rem;font:800 1.7rem/1.2 Merriweather,serif}}
.chapter h3{{margin:1.1rem 0 .45rem;font:700 1.04rem/1.35 Manrope,sans-serif}}
.chapter p{{margin:.3rem 0 .75rem}}
.chapter ul{{margin:.3rem 0 1rem 1.1rem}}
.chapter li{{margin:.25rem 0}}
.callout{{border-left:4px solid var(--accent);padding:10px 12px;background:#eef8f5;border-radius:8px;margin:12px 0}}
.visual-grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:12px 0 0}}
.visual-card{{background:#fff;border:1px solid var(--line);border-radius:12px;padding:12px 14px}}
.visual-card h3{{margin:0 0 .55rem;font:700 .97rem/1.3 Manrope,sans-serif}}
.mock-screen{{background:#111827;color:#e5f0ff;border-radius:10px;padding:10px;overflow:auto}}
.screen-title{{font:700 .8rem/1.2 Manrope,sans-serif;color:#8fd3ff;margin-bottom:8px}}
.mock-screen pre{{margin:0;white-space:pre-wrap;font:500 .82rem/1.45 "Consolas","Courier New",monospace}}
svg{{width:100%;height:auto;background:#f8fbfa;border:1px solid #d8e7e2;border-radius:10px;padding:8px}}
.node{{fill:#ffffff;stroke:#8ac6b8;stroke-width:1.5}}
.arrow{{stroke:#2c7a6a;stroke-width:2}}
.line-a{{fill:none;stroke:#0f766e;stroke-width:3}}
.line-b{{fill:none;stroke:#9a5b2d;stroke-width:3}}
.axis{{stroke:#8a7b68;stroke-width:1.2}}
.step-list{{margin:.3rem 0 0 1.1rem}}
.step-list li{{margin:.3rem 0}}
.mini-bars > div{{display:grid;grid-template-columns:150px 1fr;gap:10px;align-items:center;margin:8px 0}}
.mini-bars label{{font:600 .84rem Manrope,sans-serif;color:#544839}}
.mini-bars span{{display:block;height:12px;background:linear-gradient(90deg,#0d7a66,#63c3ab);border-radius:99px;width:var(--w)}}
.check-table{{width:100%;border-collapse:collapse;font-size:.9rem}}
.check-table th,.check-table td{{border:1px solid var(--line);padding:8px;vertical-align:top;text-align:left}}
.check-table th{{background:#f4ece0;font-family:Manrope,sans-serif;font-weight:700}}
.link-list{{margin:.35rem 0 0 1rem}}
.link-list li{{margin:.4rem 0}}
.link-list a{{color:#0b6f5b;text-decoration:underline;text-underline-offset:2px}}
.link-list a:hover{{color:#084e40}}
.sources{{padding:24px 50px;border-top:1px solid var(--line);background:#fcf7ee}}
.sources h2{{margin:0 0 8px;font:800 1.2rem Merriweather,serif}}
.sources ul{{margin:0 0 0 1rem}}
@media (max-width:980px){{.shell{{grid-template-columns:1fr}}.toc{{position:relative;top:0;max-height:none}}}}
@media (max-width:900px){{.visual-grid{{grid-template-columns:1fr}}}}
@media (max-width:640px){{.cover,.chapter,.sources{{padding:22px 18px}}}}
</style>
</head>
<body>
<div class="shell">
  <aside class="toc">
    <h1>Tally Practice Book</h1>
    <p>Internet Edition</p>
    <ul>{''.join(toc)}</ul>
  </aside>

  <main class="book">
    <section class="cover">
    <span class="badge">Tally_Book_BY_PARNAMI</span>
    <h1>Tally_Book_BY_PARNAMI</h1>
      <p>This book is generated from internet-sourced Tally learning context and rewritten for students to read, practice, and execute directly in Tally.</p>
    <div class="meta">Generated by Rasid and Mukul Sir</div>
    <div class="meta">Generated on: {esc(generated_on)}</div>

      <div class="chart">
                <h3>Learning Focus Distribution</h3>
                <div class="bar"><label>Accounting Fundamentals - 90%</label><div class="track"><div class="fill" style="width:90%"></div></div></div>
                <div class="bar"><label>Tally.ERP 9 Setup &amp; Features - 85%</label><div class="track"><div class="fill" style="width:85%"></div></div></div>
                <div class="bar"><label>Inventory + GST Practice - 82%</label><div class="track"><div class="fill" style="width:82%"></div></div></div>
                <div class="bar"><label>Case Study Execution - 88%</label><div class="track"><div class="fill" style="width:88%"></div></div></div>
      </div>
    </section>

    {''.join(body)}

    <section class="sources">
      <h2>Internet Sources Used</h2>
      <ul>{''.join(src_items)}</ul>
      <div class="callout"><strong>Note:</strong> This edition is student-focused and rewritten for practice clarity. Validate tax/legal rules against the latest official TallyHelp and government notifications before production use.</div>
    </section>
  </main>
</div>
</body>
</html>
"""


def main() -> int:
    source_blurbs = {k: fetch_summary(v) for k, v in SOURCES.items()}
    chapters = build_chapters(source_blurbs)
    html_text = build_html(chapters, source_blurbs)
    OUT_HTML.write_text(html_text, encoding="utf-8")
    LEGACY_OUT_HTML.write_text(html_text, encoding="utf-8")
    print(f"Created: {OUT_HTML}")
    print(f"Synced: {LEGACY_OUT_HTML}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
