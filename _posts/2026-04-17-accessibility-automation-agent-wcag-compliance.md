---
layout: post
title: "The A11y Agent: Automating WCAG Compliance Across Selenium, Playwright, Cypress & Appium"
date: 2026-04-17
author: Mohammad Kashif
tags: [kashif, mohammad-kashif, kashif-subrati, kashif-sdet, kashif-quality-engineer, kashif-automation-specialist, kashif-testing-blog, accessibility, wcag, automation, playwright, selenium, appium, cypress, a11y, compliance, agentic-ai, sdet, quality-engineering, ci-cd-testing]
description: "Build a reusable Accessibility Agent and Skill that plugs into Selenium, Playwright, Cypress, or Appium — and validates against WCAG 2.0, 2.1, and 2.2 conformance levels mapped to EU, US, and APAC regulations. By Mohammad Kashif Subrati — Automation Specialist."
---

Accessibility isn't optional anymore. It's **law**. Whether you're shipping a web app in the US, a mobile app in the EU, or a digital product across APAC — your software must meet accessibility standards. And the testing industry has been slow to automate this properly.

Most teams treat accessibility as a manual audit done once before a major release. That's not good enough. What if your CI/CD pipeline could **automatically validate accessibility compliance** against WCAG 2.0, 2.1, and 2.2 — at every conformance level (A, AA, AAA) — and report results mapped to the regulations that matter for your market?

That's exactly what we're building here: **a reusable Accessibility Agent and Skill** that works with any automation tool — Selenium, Appium, Playwright, or Cypress.

---

## Why an Agent-Based Approach?

Traditional accessibility testing tools like axe-core or pa11y are powerful, but they're standalone. They don't integrate cleanly into cross-framework automation pipelines. And they definitely don't map results to **regional compliance requirements** out of the box.

An **Agent** wraps this logic into a self-contained, reusable unit:

- **Framework-agnostic**: Works with Selenium, Playwright, Cypress, or Appium
- **Version-aware**: Validates against WCAG 2.0, 2.1, or 2.2 criteria
- **Level-aware**: Reports results by conformance level (A, AA, AAA)
- **Region-aware**: Maps findings to EU, US, and APAC legal requirements
- **CI/CD ready**: Can be plugged into any pipeline as a quality gate

<div class="mermaid">
flowchart TB
    subgraph AGENT["🤖 Accessibility Agent"]
        direction TB
        A1["Page/Screen Under Test"] --> A2["Inject Accessibility Engine"]
        A2 --> A3["Run WCAG Rules"]
        A3 --> A4["Classify by Level: A / AA / AAA"]
        A4 --> A5["Map to WCAG Version: 2.0 / 2.1 / 2.2"]
        A5 --> A6["Map to Region: EU / US / APAC"]
        A6 --> A7["Generate Compliance Report"]
    end

    subgraph FRAMEWORKS["🔧 Supported Frameworks"]
        F1["Selenium WebDriver"]
        F2["Playwright"]
        F3["Cypress"]
        F4["Appium"]
    end

    FRAMEWORKS --> AGENT
</div>

---

## Understanding WCAG Versions and Conformance Levels

Before we build, let's get the standards straight.

### WCAG Versions

| Version | Released | Key Additions |
|---------|----------|---------------|
| **WCAG 2.0** | 2008 | Foundation: perceivable, operable, understandable, robust (POUR) |
| **WCAG 2.1** | 2018 | Mobile accessibility, cognitive disabilities, low vision — 17 new success criteria |
| **WCAG 2.2** | 2023 | Focus appearance, dragging movements, accessible authentication — 9 new success criteria |

Each version is **backward-compatible** — 2.2 includes all of 2.1, which includes all of 2.0.

### Conformance Levels

| Level | Description | Typical Use |
|-------|-------------|-------------|
| **Level A** | Minimum barrier removal. If these fail, content is fundamentally inaccessible | Absolute baseline — no one should ship below this |
| **Level AA** | The **industry standard**. Covers contrast, navigation, input assistance | Required by most laws (ADA, Section 508, EN 301 549) |
| **Level AAA** | Highest strictness. Sign language, extended audio descriptions, cognitive aids | Aspirational for most; required in specialized contexts |

---

## Regional Compliance Requirements

### 🇺🇸 United States

| Regulation | Applies To | Required Standard |
|-----------|-----------|-------------------|
| **ADA Title III** | Public-facing websites and apps | WCAG 2.1 AA (DOJ guidance, 2024) |
| **Section 508** | Federal agencies and contractors | WCAG 2.0 AA (revised 2017, aligning with 2.1 AA in practice) |
| **State Laws (CA, NY)** | Varies by state | Generally WCAG 2.1 AA |

**Key takeaway**: If you operate in the US, **WCAG 2.1 Level AA** is your compliance target.

### 🇪🇺 European Union

| Regulation | Applies To | Required Standard |
|-----------|-----------|-------------------|
| **EN 301 549** | ICT products and services | WCAG 2.1 AA (harmonized standard) |
| **European Accessibility Act (EAA)** | Private sector digital products (from June 2025) | WCAG 2.1 AA minimum |
| **Web Accessibility Directive** | Public sector websites and apps | WCAG 2.1 AA |

**Key takeaway**: The EU is the strictest enforcer. **WCAG 2.1 Level AA** is mandatory, and the EAA extends this to private companies.

### 🌏 APAC (Asia-Pacific)

| Country/Region | Regulation | Required Standard |
|---------------|-----------|-------------------|
| **Australia** | Disability Discrimination Act (DDA) | WCAG 2.1 AA (recommended) |
| **Japan** | JIS X 8341-3 | Based on WCAG 2.0 / 2.1 AA |
| **South Korea** | KWCAG | Based on WCAG 2.1 |
| **India** | RPwD Act + GIGW Guidelines | WCAG 2.0 AA |
| **Singapore** | Digital Service Standards | WCAG 2.1 AA |
| **New Zealand** | Web Accessibility Standard | WCAG 2.1 AA |

**Key takeaway**: Most APAC nations are converging on **WCAG 2.1 AA**, with some still referencing 2.0.

---

## Architecture: The Accessibility Agent & Skill

### Core Components

```
accessibility-agent/
├── core/
│   ├── accessibility_engine.py       # Wraps axe-core / pa11y
│   ├── wcag_rules.py                 # WCAG 2.0, 2.1, 2.2 rule mappings
│   ├── conformance_classifier.py     # Classify violations by A, AA, AAA
│   └── region_mapper.py              # Map violations to EU, US, APAC requirements
├── adapters/
│   ├── selenium_adapter.py           # Selenium WebDriver integration
│   ├── playwright_adapter.py         # Playwright integration
│   ├── cypress_adapter.js            # Cypress plugin
│   └── appium_adapter.py             # Appium (mobile) integration
├── reporters/
│   ├── compliance_report.py          # Unified compliance report generator
│   ├── html_report.py                # HTML dashboard output
│   └── json_report.py                # Machine-readable JSON output
├── config/
│   └── agent_config.yaml             # Target WCAG version, level, regions
└── skills/
    ├── web_accessibility_skill.py    # Skill: scan web pages
    └── mobile_accessibility_skill.py # Skill: scan mobile screens
```

### Agent Configuration

```yaml
# agent_config.yaml
agent:
  name: "AccessibilityComplianceAgent"
  version: "1.0.0"

wcag:
  versions:
    - "2.0"
    - "2.1"
    - "2.2"
  target_level: "AA"          # A, AA, or AAA
  check_all_levels: true       # Show results for all levels

regions:
  - name: "US"
    regulations:
      - "ADA Title III"
      - "Section 508"
    required_version: "2.1"
    required_level: "AA"
  - name: "EU"
    regulations:
      - "EN 301 549"
      - "European Accessibility Act"
    required_version: "2.1"
    required_level: "AA"
  - name: "APAC"
    regulations:
      - "DDA (Australia)"
      - "JIS X 8341-3 (Japan)"
      - "KWCAG (South Korea)"
    required_version: "2.1"
    required_level: "AA"

reporting:
  format: ["html", "json"]
  include_recommendations: true
  include_screenshots: true
```

---

## Building the Core: The Framework-Agnostic Accessibility Engine

### Engine Implementation

The core engine uses **axe-core** — the industry standard accessibility rules engine — and wraps it in a framework-agnostic interface.

```python
# core/accessibility_engine.py
import json
from dataclasses import dataclass, field
from typing import List, Dict, Optional
from enum import Enum

class WCAGVersion(Enum):
    V2_0 = "2.0"
    V2_1 = "2.1"
    V2_2 = "2.2"

class ConformanceLevel(Enum):
    A = "A"
    AA = "AA"
    AAA = "AAA"

@dataclass
class Violation:
    rule_id: str
    description: str
    impact: str                  # critical, serious, moderate, minor
    wcag_criteria: List[str]     # e.g., ["1.1.1", "4.1.2"]
    wcag_versions: List[str]     # e.g., ["2.0", "2.1", "2.2"]
    conformance_level: str       # A, AA, or AAA
    affected_elements: List[str]
    recommendation: str

@dataclass
class AccessibilityResult:
    url: str
    timestamp: str
    total_violations: int
    violations_by_level: Dict[str, List[Violation]] = field(default_factory=dict)
    violations_by_version: Dict[str, List[Violation]] = field(default_factory=dict)
    compliance_status: Dict[str, bool] = field(default_factory=dict)


class AccessibilityEngine:
    """
    Framework-agnostic accessibility scanning engine.
    Accepts a page source or a driver/page object via an adapter.
    """

    # axe-core script to inject into the page
    AXE_SCRIPT_URL = "https://cdnjs.cloudflare.com/ajax/libs/axe-core/4.9.1/axe.min.js"

    # Mapping of WCAG criteria to versions they were introduced in
    CRITERIA_VERSION_MAP = {
        # WCAG 2.0 criteria (examples)
        "1.1.1": "2.0", "1.2.1": "2.0", "1.3.1": "2.0", "1.4.1": "2.0",
        "2.1.1": "2.0", "2.4.1": "2.0", "3.1.1": "2.0", "4.1.1": "2.0",
        # WCAG 2.1 new criteria
        "1.3.4": "2.1", "1.3.5": "2.1", "1.4.10": "2.1", "1.4.11": "2.1",
        "1.4.12": "2.1", "1.4.13": "2.1", "2.1.4": "2.1", "2.5.1": "2.1",
        "2.5.2": "2.1", "2.5.3": "2.1", "2.5.4": "2.1",
        # WCAG 2.2 new criteria
        "2.4.11": "2.2", "2.4.12": "2.2", "2.4.13": "2.2",
        "2.5.7": "2.2", "2.5.8": "2.2", "3.2.6": "2.2",
        "3.3.7": "2.2", "3.3.8": "2.2", "3.3.9": "2.2",
    }

    def __init__(self, config: dict):
        self.target_versions = config.get("versions", ["2.1"])
        self.target_level = config.get("target_level", "AA")
        self.check_all_levels = config.get("check_all_levels", True)

    def scan(self, adapter) -> AccessibilityResult:
        """Run accessibility scan using the provided framework adapter."""
        # Inject axe-core into the page
        adapter.inject_script(self.AXE_SCRIPT_URL)

        # Run axe-core analysis
        raw_results = adapter.execute_script(
            "return axe.run(document, {"
            "  resultTypes: ['violations'],"
            "  runOnly: { type: 'tag', values: ['wcag2a', 'wcag2aa', 'wcag2aaa', "
            "    'wcag21a', 'wcag21aa', 'wcag22aa'] }"
            "})"
        )

        return self._process_results(raw_results, adapter.get_current_url())

    def _process_results(self, raw_results: dict, url: str) -> AccessibilityResult:
        """Process axe-core results into structured violations."""
        violations = []
        for violation_data in raw_results.get("violations", []):
            violation = self._map_violation(violation_data)
            violations.append(violation)

        result = AccessibilityResult(
            url=url,
            timestamp=self._get_timestamp(),
            total_violations=len(violations),
        )

        # Classify by conformance level
        for level in ["A", "AA", "AAA"]:
            result.violations_by_level[level] = [
                v for v in violations if v.conformance_level == level
            ]

        # Classify by WCAG version
        for version in ["2.0", "2.1", "2.2"]:
            result.violations_by_version[version] = [
                v for v in violations
                if any(self._criteria_in_version(c, version) for c in v.wcag_criteria)
            ]

        return result
```

---

## One Agent, Four Frameworks: Writing the Adapters

The power of this architecture is in the **adapters**. Each adapter normalizes the interface so the core engine doesn't care which framework you're using.

### Selenium Adapter

```python
# adapters/selenium_adapter.py
import requests
from selenium.webdriver.remote.webdriver import WebDriver

class SeleniumAdapter:
    def __init__(self, driver: WebDriver):
        self.driver = driver

    def inject_script(self, script_url: str):
        script_content = requests.get(script_url, timeout=30).text
        self.driver.execute_script(script_content)

    def execute_script(self, script: str):
        return self.driver.execute_script(script)

    def get_current_url(self) -> str:
        return self.driver.current_url

    def take_screenshot(self, path: str):
        self.driver.save_screenshot(path)
```

### Playwright Adapter

```python
# adapters/playwright_adapter.py
class PlaywrightAdapter:
    def __init__(self, page):
        self.page = page

    def inject_script(self, script_url: str):
        self.page.add_script_tag(url=script_url)
        self.page.wait_for_function("typeof axe !== 'undefined'")

    def execute_script(self, script: str):
        return self.page.evaluate(script)

    def get_current_url(self) -> str:
        return self.page.url

    def take_screenshot(self, path: str):
        self.page.screenshot(path=path, full_page=True)
```

### Cypress Plugin

```javascript
// adapters/cypress_adapter.js
// Install: npm install cypress-axe axe-core

Cypress.Commands.add('runAccessibilityAgent', (config = {}) => {
  const targetLevel = config.targetLevel || 'AA';
  const wcagVersions = config.wcagVersions || ['2.0', '2.1', '2.2'];

  cy.injectAxe();

  const runOnlyTags = [];
  if (wcagVersions.includes('2.0')) runOnlyTags.push('wcag2a', 'wcag2aa', 'wcag2aaa');
  if (wcagVersions.includes('2.1')) runOnlyTags.push('wcag21a', 'wcag21aa');
  if (wcagVersions.includes('2.2')) runOnlyTags.push('wcag22aa');

  cy.checkA11y(null, {
    runOnly: { type: 'tag', values: runOnlyTags }
  }, (violations) => {
    const result = {
      url: cy.url(),
      totalViolations: violations.length,
      byLevel: { A: [], AA: [], AAA: [] },
      byVersion: { '2.0': [], '2.1': [], '2.2': [] },
    };

    violations.forEach(v => {
      const level = extractLevel(v.tags);
      const version = extractVersion(v.tags);
      result.byLevel[level].push(v);
      result.byVersion[version].push(v);
    });

    return result;
  });
});
```

### Appium Adapter (Mobile)

```python
# adapters/appium_adapter.py
class AppiumAdapter:
    """
    For mobile apps, axe-core doesn't apply directly.
    Uses platform-specific accessibility APIs instead.
    """
    def __init__(self, driver):
        self.driver = driver
        self.platform = driver.capabilities.get('platformName', '').lower()

    def scan_native_accessibility(self) -> dict:
        """
        Uses Appium's accessibility features to check
        native mobile accessibility attributes.
        """
        if self.platform == 'android':
            return self._scan_android()
        elif self.platform == 'ios':
            return self._scan_ios()

    def _scan_android(self) -> dict:
        """Check Android accessibility properties via UIAutomator."""
        elements = self.driver.find_elements_by_xpath("//*")
        violations = []
        for el in elements:
            # Check content-description (alt text equivalent)
            if el.tag_name in ['android.widget.ImageView', 'android.widget.ImageButton']:
                if not el.get_attribute('content-desc'):
                    violations.append({
                        'rule': '1.1.1',
                        'level': 'A',
                        'description': 'Image missing content description',
                        'element': el.get_attribute('resource-id'),
                    })
            # Check touch target size (WCAG 2.5.5 / 2.5.8)
            size = el.size
            if el.is_displayed() and el.is_enabled():
                if size['width'] < 44 or size['height'] < 44:
                    violations.append({
                        'rule': '2.5.8',
                        'level': 'AA',
                        'description': f'Touch target too small: {size["width"]}x{size["height"]}px (min 44x44)',
                        'element': el.get_attribute('resource-id'),
                    })
        return {'violations': violations}

    def _scan_ios(self) -> dict:
        """Check iOS accessibility properties."""
        elements = self.driver.find_elements_by_xpath("//*")
        violations = []
        for el in elements:
            accessible = el.get_attribute('accessible')
            label = el.get_attribute('label')
            if el.tag_name in ['XCUIElementTypeImage', 'XCUIElementTypeButton']:
                if not label:
                    violations.append({
                        'rule': '1.1.1',
                        'level': 'A',
                        'description': 'Element missing accessibility label',
                        'element': el.get_attribute('name'),
                    })
        return {'violations': violations}
```

---

## From Violations to Verdicts: The Regional Compliance Reporter

This is where results get mapped to **regional regulations** — the part that makes this agent truly valuable.

```python
# reporters/compliance_report.py
from dataclasses import dataclass
from typing import List, Dict

@dataclass
class RegionalCompliance:
    region: str
    regulations: List[str]
    required_version: str
    required_level: str
    is_compliant: bool
    total_violations: int
    critical_violations: int
    details: List[dict]

class ComplianceReporter:
    REGIONAL_REQUIREMENTS = {
        "US": {
            "regulations": ["ADA Title III", "Section 508"],
            "required_version": "2.1",
            "required_level": "AA",
        },
        "EU": {
            "regulations": ["EN 301 549", "European Accessibility Act (EAA)",
                           "Web Accessibility Directive"],
            "required_version": "2.1",
            "required_level": "AA",
        },
        "APAC": {
            "regulations": ["DDA (Australia)", "JIS X 8341-3 (Japan)",
                           "KWCAG (South Korea)", "RPwD Act (India)",
                           "Digital Service Standards (Singapore)"],
            "required_version": "2.1",
            "required_level": "AA",
        },
    }

    def generate_report(self, result, regions: List[str] = None) -> Dict[str, RegionalCompliance]:
        """Generate compliance report for specified regions."""
        regions = regions or ["US", "EU", "APAC"]
        report = {}

        for region in regions:
            req = self.REGIONAL_REQUIREMENTS[region]
            relevant_violations = self._get_violations_for_requirement(
                result, req["required_version"], req["required_level"]
            )

            report[region] = RegionalCompliance(
                region=region,
                regulations=req["regulations"],
                required_version=req["required_version"],
                required_level=req["required_level"],
                is_compliant=len(relevant_violations) == 0,
                total_violations=len(relevant_violations),
                critical_violations=sum(
                    1 for v in relevant_violations if v.impact == "critical"
                ),
                details=[self._violation_to_dict(v) for v in relevant_violations],
            )

        return report

    def _get_violations_for_requirement(self, result, version, level):
        """Filter violations relevant to a specific version + level requirement."""
        level_hierarchy = {"A": ["A"], "AA": ["A", "AA"], "AAA": ["A", "AA", "AAA"]}
        required_levels = level_hierarchy[level]

        version_hierarchy = {
            "2.0": ["2.0"],
            "2.1": ["2.0", "2.1"],
            "2.2": ["2.0", "2.1", "2.2"],
        }
        required_versions = version_hierarchy[version]

        relevant = []
        for v_level in required_levels:
            for violation in result.violations_by_level.get(v_level, []):
                if any(ver in required_versions for ver in violation.wcag_versions):
                    relevant.append(violation)

        return relevant

    def print_summary(self, report: Dict[str, RegionalCompliance]):
        """Print a human-readable compliance summary."""
        for region, compliance in report.items():
            status = "✅ COMPLIANT" if compliance.is_compliant else "❌ NON-COMPLIANT"
            print(f"\n{'='*60}")
            print(f"  {region} Compliance: {status}")
            print(f"  Standard: WCAG {compliance.required_version} Level {compliance.required_level}")
            print(f"  Regulations: {', '.join(compliance.regulations)}")
            print(f"  Violations: {compliance.total_violations} "
                  f"(Critical: {compliance.critical_violations})")
            print(f"{'='*60}")
```

---

## Version Delta: What Changes Between WCAG 2.0, 2.1, and 2.2

One of the most powerful features: **see exactly what changes between WCAG versions** for your application.

```python
# core/version_comparator.py
class WCAGVersionComparator:
    """Compare accessibility results across WCAG 2.0, 2.1, and 2.2."""

    def compare(self, result) -> dict:
        """Show which violations are new in each version."""
        only_in_2_0 = set()
        new_in_2_1 = set()
        new_in_2_2 = set()

        for v in result.violations_by_version.get("2.0", []):
            only_in_2_0.add(v.rule_id)

        for v in result.violations_by_version.get("2.1", []):
            if v.rule_id not in only_in_2_0:
                new_in_2_1.add(v.rule_id)

        for v in result.violations_by_version.get("2.2", []):
            if v.rule_id not in only_in_2_0 and v.rule_id not in new_in_2_1:
                new_in_2_2.add(v.rule_id)

        return {
            "wcag_2.0": {
                "total_criteria_checked": len(only_in_2_0),
                "violations": list(only_in_2_0),
            },
            "wcag_2.1": {
                "new_criteria": 17,
                "new_violations": list(new_in_2_1),
                "cumulative_violations": len(only_in_2_0) + len(new_in_2_1),
            },
            "wcag_2.2": {
                "new_criteria": 9,
                "new_violations": list(new_in_2_2),
                "cumulative_violations": len(only_in_2_0) + len(new_in_2_1) + len(new_in_2_2),
            },
        }
```

---

## Putting It All Together: Real-World Usage

### With Playwright

```python
from playwright.sync_api import sync_playwright
from core.accessibility_engine import AccessibilityEngine
from adapters.playwright_adapter import PlaywrightAdapter
from reporters.compliance_report import ComplianceReporter

config = {
    "versions": ["2.0", "2.1", "2.2"],
    "target_level": "AA",
    "check_all_levels": True,
}

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto("https://your-application.com")

    # Create the agent components
    engine = AccessibilityEngine(config)
    adapter = PlaywrightAdapter(page)

    # Run the scan
    result = engine.scan(adapter)

    # Generate compliance reports for all regions
    reporter = ComplianceReporter()
    report = reporter.generate_report(result, regions=["US", "EU", "APAC"])
    reporter.print_summary(report)

    browser.close()
```

### With Selenium

```python
from selenium import webdriver
from core.accessibility_engine import AccessibilityEngine
from adapters.selenium_adapter import SeleniumAdapter
from reporters.compliance_report import ComplianceReporter

driver = webdriver.Chrome()
driver.get("https://your-application.com")

engine = AccessibilityEngine({"versions": ["2.0", "2.1", "2.2"], "target_level": "AA"})
adapter = SeleniumAdapter(driver)

result = engine.scan(adapter)
reporter = ComplianceReporter()
report = reporter.generate_report(result)
reporter.print_summary(report)

driver.quit()
```

### With Cypress

```javascript
// cypress/e2e/accessibility.cy.js
describe('Accessibility Compliance', () => {
  it('should meet WCAG 2.1 AA for US and EU compliance', () => {
    cy.visit('/');
    cy.runAccessibilityAgent({
      targetLevel: 'AA',
      wcagVersions: ['2.0', '2.1', '2.2'],
      regions: ['US', 'EU', 'APAC'],
    });
  });

  it('should check critical user flows', () => {
    const criticalPages = ['/', '/login', '/checkout', '/dashboard'];

    criticalPages.forEach(page => {
      cy.visit(page);
      cy.runAccessibilityAgent({ targetLevel: 'AA' });
    });
  });
});
```

### With Appium (Mobile)

```python
from appium import webdriver as appium_driver
from adapters.appium_adapter import AppiumAdapter
from reporters.compliance_report import ComplianceReporter

caps = {
    "platformName": "Android",
    "deviceName": "Pixel_6",
    "app": "/path/to/your.apk",
    "automationName": "UiAutomator2",
}

driver = appium_driver.Remote("http://localhost:4723/wd/hub", caps)

adapter = AppiumAdapter(driver)
result = adapter.scan_native_accessibility()

reporter = ComplianceReporter()
report = reporter.generate_report(result, regions=["US", "EU", "APAC"])
reporter.print_summary(report)

driver.quit()
```

---

## Sample Output: Compliance Report

Here's what the agent produces after scanning a page:

```
============================================================
  US Compliance: ❌ NON-COMPLIANT
  Standard: WCAG 2.1 Level AA
  Regulations: ADA Title III, Section 508
  Violations: 7 (Critical: 2)
============================================================
  Critical:
    [1.1.1] Level A — 3 images missing alt text
    [1.4.3] Level AA — 2 elements with insufficient color contrast

  Serious:
    [2.4.6] Level AA — Page missing descriptive headings
    [4.1.2] Level A — Form inputs missing accessible names

============================================================
  EU Compliance: ❌ NON-COMPLIANT
  Standard: WCAG 2.1 Level AA
  Regulations: EN 301 549, European Accessibility Act (EAA)
  Violations: 9 (Critical: 3)
============================================================
  (Same as US, plus:)
    [1.4.11] Level AA — UI components lack non-text contrast
    [2.5.3] Level A — Accessible name doesn't match visible label

============================================================
  APAC Compliance: ❌ NON-COMPLIANT
  Standard: WCAG 2.1 Level AA
  Regulations: DDA, JIS X 8341-3, KWCAG
  Violations: 7 (Critical: 2)
============================================================

--- WCAG Version Comparison ---
  WCAG 2.0: 5 violations found
  WCAG 2.1: 5 + 2 new = 7 violations (new: 1.4.11, 2.5.3)
  WCAG 2.2: 7 + 2 new = 9 violations (new: 2.4.11, 3.3.8)
```

---

## CI/CD Integration: Accessibility as a Quality Gate

Drop this into your pipeline as a quality gate:

```yaml
# .github/workflows/accessibility.yml
name: Accessibility Compliance Check

on:
  pull_request:
    branches: [main, develop]

jobs:
  a11y-check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Start application
        run: npm start &

      - name: Run Accessibility Agent
        run: |
          python -m accessibility_agent.run \
            --url http://localhost:3000 \
            --wcag-versions 2.0,2.1,2.2 \
            --level AA \
            --regions US,EU,APAC \
            --fail-on critical \
            --output reports/a11y-report.html

      - name: Upload Report
        uses: actions/upload-artifact@v4
        with:
          name: accessibility-report
          path: reports/a11y-report.html

      - name: Fail if non-compliant
        run: |
          python -m accessibility_agent.check_compliance \
            --report reports/a11y-report.json \
            --require-compliant US,EU
```

---

## The Cheat Sheet: Key WCAG Criteria Every Test Engineer Should Know

| Criterion | Level | Version | What It Checks |
|-----------|-------|---------|----------------|
| 1.1.1 Non-text Content | A | 2.0 | Images have alt text |
| 1.3.1 Info and Relationships | A | 2.0 | Semantic HTML structure |
| 1.4.3 Contrast (Minimum) | AA | 2.0 | 4.5:1 text contrast ratio |
| 1.4.11 Non-text Contrast | AA | 2.1 | 3:1 contrast for UI components |
| 2.1.1 Keyboard | A | 2.0 | All functionality via keyboard |
| 2.4.6 Headings and Labels | AA | 2.0 | Descriptive headings |
| 2.5.8 Target Size (Minimum) | AA | 2.2 | 24x24px minimum touch targets |
| 3.3.7 Redundant Entry | A | 2.2 | Don't ask for same info twice |
| 3.3.8 Accessible Authentication | AA | 2.2 | No cognitive function tests for login |
| 4.1.2 Name, Role, Value | A | 2.0 | ARIA attributes on custom components |

---

## The SDET Mindset: Accessibility Is a Testing Discipline

Accessibility testing isn't a separate activity — it's a **quality dimension** that belongs in your automation suite alongside functional, performance, and security testing.

By building this as an **Agent with Skills**:

1. **Reusability** — Write once, use across Selenium, Playwright, Cypress, or Appium
2. **Consistency** — Same rules, same reporting, regardless of the framework
3. **Compliance clarity** — Know exactly where you stand for US, EU, and APAC regulations
4. **Version tracking** — See how WCAG 2.0 → 2.1 → 2.2 impacts your compliance
5. **Shift-left** — Catch accessibility violations in CI before they reach production

The tools exist. The standards are clear. The legal requirements are tightening. The only question is: **are you testing for it?**

---

*Start with Level AA. Start with one framework. Start with your most critical user flows. Then expand. That's how you build accessibility into your engineering culture — not as an afterthought, but as a first-class quality signal.*
