/* ══════════════════════════════════════════════════════════
   AI Resume Builder — JavaScript Engine
   Uses Google Gemini API (free tier) for AI generation.
   All processing is client-side. No backend needed.
   ══════════════════════════════════════════════════════════ */

// ── State ────────────────────────────────────────────────
let currentStep = 1;
let generatedResume = '';
let generatedCover = '';

// ── API Key Management ───────────────────────────────────

function saveApiKey() {
  const key = document.getElementById('apiKeyInput').value.trim();
  if (!key) {
    showToast('Please enter an API key', 'error');
    return;
  }
  localStorage.setItem('a11y_gemini_key', key);
  document.getElementById('apiBanner').classList.add('saved');
  document.getElementById('apiKeyInput').value = '••••••••' + key.slice(-4);
  showToast('API key saved to browser', 'success');
}

function getApiKey() {
  return localStorage.getItem('a11y_gemini_key') || '';
}

// Restore saved key indicator on load
document.addEventListener('DOMContentLoaded', () => {
  const key = getApiKey();
  if (key) {
    document.getElementById('apiBanner').classList.add('saved');
    document.getElementById('apiKeyInput').value = '••••••••' + key.slice(-4);
  }

  // Setup drag-and-drop
  setupDragDrop();

  // Setup tab switching
  setupTabs();
  setupPreviewTabs();
});

// ── Tab Switching ────────────────────────────────────────

function setupTabs() {
  document.querySelectorAll('.rb-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      switchTab(tab.dataset.tab);
    });
  });
}

function switchTab(tabName) {
  document.querySelectorAll('.rb-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.rb-tab-content').forEach(c => c.classList.remove('active'));

  const activeTab = document.querySelector(`.rb-tab[data-tab="${tabName}"]`);
  const activeContent = document.getElementById(`tab-${tabName}`);
  if (activeTab) activeTab.classList.add('active');
  if (activeContent) activeContent.classList.add('active');
}

function setupPreviewTabs() {
  document.querySelectorAll('.rb-preview-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.rb-preview-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.rb-preview-pane').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(`preview-${tab.dataset.preview}`).classList.add('active');
    });
  });
}

// ── Step Navigation ──────────────────────────────────────

function goToStep(step) {
  // Validate before moving forward
  if (step === 2 && currentStep === 1) {
    if (!getProfileData().hasContent) {
      showToast('Please fill in your profile details or upload a resume', 'error');
      return;
    }
  }

  if (step === 3 && currentStep === 2) {
    showToast('Please click "Generate with AI" first', 'info');
    return;
  }

  // Hide all sections
  document.querySelectorAll('.rb-section').forEach(s => s.classList.add('rb-hidden'));

  // Show target
  document.getElementById(`section${step}`).classList.remove('rb-hidden');

  // Update step indicator
  document.querySelectorAll('.rb-step-dot').forEach(dot => {
    const dotStep = parseInt(dot.dataset.step);
    dot.classList.remove('active', 'completed');
    if (dotStep === step) dot.classList.add('active');
    else if (dotStep < step) dot.classList.add('completed');
  });

  currentStep = step;
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── File Upload & Parsing ────────────────────────────────

function setupDragDrop() {
  const zone = document.getElementById('uploadZone');
  const fileInput = document.getElementById('resumeFile');

  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('dragover');
  });

  zone.addEventListener('dragleave', () => {
    zone.classList.remove('dragover');
  });

  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length) handleFileUpload(files[0]);
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFileUpload(e.target.files[0]);
  });
}

async function handleFileUpload(file) {
  const status = document.getElementById('uploadStatus');
  const maxSize = 10 * 1024 * 1024; // 10MB

  if (file.size > maxSize) {
    status.textContent = 'File too large. Max 10MB.';
    status.className = 'rb-upload-status error';
    return;
  }

  status.textContent = `Parsing ${file.name}...`;
  status.className = 'rb-upload-status success';

  try {
    let text = '';

    if (file.name.endsWith('.txt')) {
      text = await file.text();
    } else if (file.name.endsWith('.pdf')) {
      text = await parsePDF(file);
    } else if (file.name.endsWith('.docx') || file.name.endsWith('.doc')) {
      text = await parseDOCX(file);
    } else {
      throw new Error('Unsupported file format. Use .pdf, .docx, or .txt');
    }

    // Fill the paste fields with extracted text
    fillFieldsFromText(text);
    status.textContent = `Parsed "${file.name}" successfully. Check the "Paste Details" tab — fields have been populated.`;
    status.className = 'rb-upload-status success';
    showToast('Resume parsed! Review your details in "Paste Details" tab.', 'success');

    // Switch to paste tab to show filled fields
    setTimeout(() => switchTab('paste'), 500);
  } catch (err) {
    status.textContent = `Error: ${err.message}`;
    status.className = 'rb-upload-status error';
  }
}

async function parsePDF(file) {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let text = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map(item => item.str).join(' ') + '\n';
  }
  return text;
}

async function parseDOCX(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value;
}

function fillFieldsFromText(text) {
  // Smart extraction — try to find name, email, phone, etc.
  const emailMatch = text.match(/[\w.+-]+@[\w-]+\.[\w.-]+/);
  const phoneMatch = text.match(/[\+]?[\d\s\-().]{10,}/);
  const linkedinMatch = text.match(/linkedin\.com\/in\/[\w-]+/i);
  const githubMatch = text.match(/github\.com\/[\w-]+/i);

  if (emailMatch) document.getElementById('email').value = emailMatch[0];
  if (phoneMatch) document.getElementById('phone').value = phoneMatch[0].trim();
  if (linkedinMatch) document.getElementById('linkedin').value = 'https://' + linkedinMatch[0];
  if (githubMatch) document.getElementById('github').value = 'https://' + githubMatch[0];

  // Put the full text into experience for AI to work with
  document.getElementById('experience').value = text;
}

// ── Collect Profile Data ────────────────────────────────

function getProfileData() {
  const data = {
    fullName: document.getElementById('fullName').value.trim(),
    email: document.getElementById('email').value.trim(),
    phone: document.getElementById('phone').value.trim(),
    location: document.getElementById('location').value.trim(),
    linkedin: document.getElementById('linkedin').value.trim(),
    github: document.getElementById('github').value.trim(),
    summary: document.getElementById('summary').value.trim(),
    experience: document.getElementById('experience').value.trim(),
    skills: document.getElementById('skills').value.trim(),
    education: document.getElementById('education').value.trim(),
    certifications: document.getElementById('certifications').value.trim(),
  };

  data.hasContent = !!(data.fullName || data.experience || data.summary);
  return data;
}

function getJobData() {
  return {
    title: document.getElementById('jobTitle').value.trim(),
    company: document.getElementById('companyName').value.trim(),
    description: document.getElementById('jobDescription').value.trim(),
    notes: document.getElementById('additionalNotes').value.trim(),
  };
}

// ── AI Generation (Gemini API) ───────────────────────────

async function generateResume() {
  const apiKey = getApiKey();
  if (!apiKey) {
    showToast('Please save your Gemini API key first', 'error');
    document.getElementById('apiKeyInput').focus();
    return;
  }

  const profile = getProfileData();
  const job = getJobData();

  if (!job.description) {
    showToast('Please paste the job description', 'error');
    return;
  }

  const btn = document.getElementById('generateBtn');
  btn.querySelector('.rb-btn-text').classList.add('rb-hidden');
  btn.querySelector('.rb-btn-loading').classList.remove('rb-hidden');
  btn.disabled = true;

  try {
    // Generate resume and cover letter in parallel
    const [resumeHtml, coverHtml] = await Promise.all([
      callGemini(apiKey, buildResumePrompt(profile, job)),
      callGemini(apiKey, buildCoverLetterPrompt(profile, job)),
    ]);

    generatedResume = resumeHtml;
    generatedCover = coverHtml;

    // Render into preview
    document.getElementById('resumePreview').innerHTML = sanitizeHTML(resumeHtml);
    document.getElementById('coverPreview').innerHTML = sanitizeHTML(coverHtml);

    // Move to step 3
    goToStepDirect(3);
    showToast('Resume & cover letter generated!', 'success');
  } catch (err) {
    showToast(`AI Error: ${err.message}`, 'error');
  } finally {
    btn.querySelector('.rb-btn-text').classList.remove('rb-hidden');
    btn.querySelector('.rb-btn-loading').classList.add('rb-hidden');
    btn.disabled = false;
  }
}

async function regenerateSection(section) {
  const apiKey = getApiKey();
  if (!apiKey) { showToast('API key required', 'error'); return; }

  const profile = getProfileData();
  const job = getJobData();

  showToast(`Regenerating ${section}...`, 'info');

  try {
    const prompt = section === 'resume'
      ? buildResumePrompt(profile, job)
      : buildCoverLetterPrompt(profile, job);

    const html = await callGemini(apiKey, prompt);
    const previewEl = document.getElementById(section === 'resume' ? 'resumePreview' : 'coverPreview');
    previewEl.innerHTML = sanitizeHTML(html);

    if (section === 'resume') generatedResume = html;
    else generatedCover = html;

    showToast(`${section} regenerated!`, 'success');
  } catch (err) {
    showToast(`Error: ${err.message}`, 'error');
  }
}

function goToStepDirect(step) {
  document.querySelectorAll('.rb-section').forEach(s => s.classList.add('rb-hidden'));
  document.getElementById(`section${step}`).classList.remove('rb-hidden');
  document.querySelectorAll('.rb-step-dot').forEach(dot => {
    const dotStep = parseInt(dot.dataset.step);
    dot.classList.remove('active', 'completed');
    if (dotStep === step) dot.classList.add('active');
    else if (dotStep < step) dot.classList.add('completed');
  });
  currentStep = step;
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

async function callGemini(apiKey, prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(apiKey)}`;

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 4096,
      },
    }),
  });

  if (!resp.ok) {
    const errData = await resp.json().catch(() => ({}));
    throw new Error(errData?.error?.message || `API returned ${resp.status}`);
  }

  const data = await resp.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';

  // Extract HTML from markdown code block if present
  const htmlMatch = text.match(/```html\s*([\s\S]*?)```/);
  return htmlMatch ? htmlMatch[1].trim() : text;
}

// ── Prompt Builders ──────────────────────────────────────

function buildResumePrompt(profile, job) {
  return `You are an expert resume writer. Generate a professional, ATS-friendly resume in clean HTML format.

CANDIDATE INFORMATION:
- Name: ${profile.fullName || 'Not provided'}
- Email: ${profile.email || 'Not provided'}
- Phone: ${profile.phone || 'Not provided'}
- Location: ${profile.location || 'Not provided'}
- LinkedIn: ${profile.linkedin || 'Not provided'}
- GitHub: ${profile.github || 'Not provided'}
- Summary: ${profile.summary || 'Not provided'}
- Experience: ${profile.experience || 'Not provided'}
- Skills: ${profile.skills || 'Not provided'}
- Education: ${profile.education || 'Not provided'}
- Certifications: ${profile.certifications || 'Not provided'}

TARGET ROLE:
- Job Title: ${job.title || 'Not specified'}
- Company: ${job.company || 'Not specified'}
- Job Description: ${job.description}

${job.notes ? `ADDITIONAL INSTRUCTIONS: ${job.notes}` : ''}

REQUIREMENTS:
1. Output ONLY clean HTML (no <html>, <head>, <body> tags — just the resume content).
2. Use a centered header with name, contact info.
3. Use <h2> for section headers: PROFESSIONAL SUMMARY, EXPERIENCE, SKILLS, EDUCATION, CERTIFICATIONS.
4. Use a <div class="resume-header"> for the header section.
5. Use a <p class="resume-contact"> for contact details.
6. Tailor the experience and summary to match the job description.
7. Use strong action verbs and quantified achievements.
8. Highlight keywords from the job description naturally.
9. Keep it to 1-2 pages worth of content.
10. Use <ul> with <li> for bullet points in experience.
11. Do NOT include any CSS or style tags.
12. Do NOT wrap in a code block.`;
}

function buildCoverLetterPrompt(profile, job) {
  return `You are an expert cover letter writer. Generate a professional, compelling cover letter in clean HTML format.

CANDIDATE INFORMATION:
- Name: ${profile.fullName || 'Not provided'}
- Email: ${profile.email || 'Not provided'}
- Phone: ${profile.phone || 'Not provided'}
- Location: ${profile.location || 'Not provided'}
- Summary: ${profile.summary || 'Not provided'}
- Key Experience: ${profile.experience ? profile.experience.substring(0, 2000) : 'Not provided'}
- Skills: ${profile.skills || 'Not provided'}

TARGET ROLE:
- Job Title: ${job.title || 'Not specified'}
- Company: ${job.company || 'Not specified'}
- Job Description: ${job.description}

${job.notes ? `ADDITIONAL INSTRUCTIONS: ${job.notes}` : ''}

REQUIREMENTS:
1. Output ONLY clean HTML (no <html>, <head>, <body> tags — just the letter content).
2. Format as a proper business letter: date, greeting, 3-4 paragraphs, closing.
3. Opening paragraph: Express enthusiasm for the specific role and company.
4. Body paragraphs: Connect candidate's experience to key requirements in the JD. Use specific examples.
5. Closing paragraph: Reiterate interest, call to action.
6. Be genuine, confident, and specific — avoid generic filler.
7. Keep it under 400 words.
8. Use <p> tags for paragraphs.
9. Do NOT include any CSS or style tags.
10. Do NOT wrap in a code block.`;
}

// ── HTML Sanitizer (basic XSS protection) ────────────────

function sanitizeHTML(html) {
  // Allow only safe tags for resume rendering
  const temp = document.createElement('div');
  temp.innerHTML = html;

  // Remove script, style, iframe, object, embed, form elements
  const dangerous = temp.querySelectorAll('script, style, iframe, object, embed, form, input, textarea, select, button, link, meta');
  dangerous.forEach(el => el.remove());

  // Remove on* event attributes
  temp.querySelectorAll('*').forEach(el => {
    Array.from(el.attributes).forEach(attr => {
      if (attr.name.startsWith('on') || attr.name === 'srcdoc') {
        el.removeAttribute(attr.name);
      }
      if (attr.name === 'href' || attr.name === 'src' || attr.name === 'action') {
        const val = attr.value.trim().toLowerCase();
        if (val.startsWith('javascript:') || val.startsWith('data:text/html')) {
          el.removeAttribute(attr.name);
        }
      }
    });
  });

  return temp.innerHTML;
}

// ── Edit Preview ─────────────────────────────────────────

function editPreview(section) {
  const paper = document.getElementById(section === 'resume' ? 'resumePreview' : 'coverPreview');
  const isEditing = paper.contentEditable === 'true';

  if (isEditing) {
    paper.contentEditable = 'false';
    if (section === 'resume') generatedResume = paper.innerHTML;
    else generatedCover = paper.innerHTML;
    showToast('Changes saved', 'success');
  } else {
    paper.contentEditable = 'true';
    paper.focus();
    showToast('Editing mode — click "Edit" again to save', 'info');
  }
}

// ── Export: PDF ──────────────────────────────────────────

function exportPDF() {
  const activePane = document.querySelector('.rb-preview-pane.active');
  const paper = activePane.querySelector('.rb-preview-paper');
  const type = activePane.id.includes('resume') ? 'Resume' : 'Cover_Letter';
  const name = document.getElementById('fullName').value.trim() || 'Document';

  const opt = {
    margin: 0.5,
    filename: `${name.replace(/\s+/g, '_')}_${type}.pdf`,
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' },
  };

  showToast('Generating PDF...', 'info');
  html2pdf().set(opt).from(paper).save().then(() => {
    showToast('PDF downloaded!', 'success');
  });
}

// ── Export: DOCX ─────────────────────────────────────────

function exportDOCX() {
  const activePane = document.querySelector('.rb-preview-pane.active');
  const paper = activePane.querySelector('.rb-preview-paper');
  const type = activePane.id.includes('resume') ? 'Resume' : 'Cover_Letter';
  const name = document.getElementById('fullName').value.trim() || 'Document';

  showToast('Generating DOCX...', 'info');

  try {
    const content = paper.innerText;
    const lines = content.split('\n').filter(l => l.trim());

    const children = [];

    lines.forEach(line => {
      const trimmed = line.trim();
      if (!trimmed) return;

      // Detect headers (ALL CAPS lines or short lines)
      if (trimmed === trimmed.toUpperCase() && trimmed.length < 50 && trimmed.length > 2) {
        children.push(new docx.Paragraph({
          text: trimmed,
          heading: docx.HeadingLevel.HEADING_2,
          spacing: { before: 300, after: 100 },
        }));
      }
      // Detect bullet points
      else if (trimmed.startsWith('•') || trimmed.startsWith('-') || trimmed.startsWith('–')) {
        children.push(new docx.Paragraph({
          text: trimmed.replace(/^[•\-–]\s*/, ''),
          bullet: { level: 0 },
        }));
      }
      // Regular paragraph
      else {
        children.push(new docx.Paragraph({
          text: trimmed,
          spacing: { after: 80 },
        }));
      }
    });

    const doc = new docx.Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 720, right: 720, bottom: 720, left: 720 },
          },
        },
        children: children,
      }],
    });

    docx.Packer.toBlob(doc).then(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${name.replace(/\s+/g, '_')}_${type}.docx`;
      a.click();
      URL.revokeObjectURL(url);
      showToast('DOCX downloaded!', 'success');
    });
  } catch (err) {
    showToast(`DOCX error: ${err.message}`, 'error');
  }
}

// ── Export: Clipboard ────────────────────────────────────

function copyToClipboard() {
  const activePane = document.querySelector('.rb-preview-pane.active');
  const paper = activePane.querySelector('.rb-preview-paper');

  navigator.clipboard.writeText(paper.innerText).then(() => {
    showToast('Copied to clipboard!', 'success');
  }).catch(() => {
    // Fallback
    const range = document.createRange();
    range.selectNodeContents(paper);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('copy');
    sel.removeAllRanges();
    showToast('Copied to clipboard!', 'success');
  });
}

// ── Toast Notifications ──────────────────────────────────

function showToast(message, type = 'info') {
  // Remove existing toasts
  document.querySelectorAll('.rb-toast').forEach(t => t.remove());

  const toast = document.createElement('div');
  toast.className = `rb-toast ${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);

  requestAnimationFrame(() => {
    toast.classList.add('show');
  });

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}
