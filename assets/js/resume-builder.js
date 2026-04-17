/* ══════════════════════════════════════════════════════════
   AI Resume Builder — Multi-Provider Engine
   Supports: HuggingFace (free, no key), Groq (free key),
   Gemini (free key), OpenRouter (free key).
   All processing is client-side. No backend needed.
   ══════════════════════════════════════════════════════════ */

// ── Provider Registry ────────────────────────────────────

const AI_PROVIDERS = {
  // ── FREE: No API key required ──────────────────────────
  'huggingface-free': {
    label: 'HuggingFace Mistral (Free)',
    model: 'mistralai/Mistral-7B-Instruct-v0.3',
    needsKey: false,
    getKeyUrl: null,
    keyStorageId: null,
    call: callHuggingFaceFree,
  },
  'huggingface-qwen-free': {
    label: 'HuggingFace Qwen 2.5 (Free)',
    model: 'Qwen/Qwen2.5-72B-Instruct',
    needsKey: false,
    getKeyUrl: null,
    keyStorageId: null,
    call: callHuggingFaceFree,
  },
  'huggingface-llama-free': {
    label: 'HuggingFace Llama 3.2 (Free)',
    model: 'meta-llama/Llama-3.2-3B-Instruct',
    needsKey: false,
    getKeyUrl: null,
    keyStorageId: null,
    call: callHuggingFaceFree,
  },
  // ── FREE KEY: Requires a free API key ──────────────────
  'groq': {
    label: 'Groq — Llama 3.3 70B (Free Key)',
    model: 'llama-3.3-70b-versatile',
    needsKey: true,
    getKeyUrl: 'https://console.groq.com/keys',
    keyStorageId: 'rb_key_groq',
    call: callGroq,
  },
  'groq-mixtral': {
    label: 'Groq — Mixtral 8x7B (Free Key)',
    model: 'mixtral-8x7b-32768',
    needsKey: true,
    getKeyUrl: 'https://console.groq.com/keys',
    keyStorageId: 'rb_key_groq',
    call: callGroq,
  },
  'groq-deepseek': {
    label: 'Groq — DeepSeek R1 Distill (Free Key)',
    model: 'deepseek-r1-distill-llama-70b',
    needsKey: true,
    getKeyUrl: 'https://console.groq.com/keys',
    keyStorageId: 'rb_key_groq',
    call: callGroq,
  },
  'gemini': {
    label: 'Google Gemini 2.0 Flash (Free Key)',
    model: 'gemini-2.0-flash',
    needsKey: true,
    getKeyUrl: 'https://aistudio.google.com/app/apikey',
    keyStorageId: 'rb_key_gemini',
    call: callGemini,
  },
  'openrouter-free': {
    label: 'OpenRouter — Free Models (Free Key)',
    model: 'meta-llama/llama-3.3-8b-instruct:free',
    needsKey: true,
    getKeyUrl: 'https://openrouter.ai/keys',
    keyStorageId: 'rb_key_openrouter',
    call: callOpenRouter,
  },
};

// ── State ────────────────────────────────────────────────
let currentStep = 1;
let generatedResume = '';
let generatedCover = '';
let currentProvider = 'groq';

// ── DOM Ready ────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  // Restore saved provider
  const saved = localStorage.getItem('rb_provider');
  if (saved && AI_PROVIDERS[saved]) {
    currentProvider = saved;
    const sel = document.getElementById('providerSelect');
    if (sel) sel.value = saved;
  }
  onProviderChange();
  setupDragDrop();
  setupTabs();
  setupPreviewTabs();
  populateRegenSelectors();
});

// ── Provider Selection ───────────────────────────────────

function onProviderChange() {
  const select = document.getElementById('providerSelect');
  currentProvider = select.value;
  const provider = AI_PROVIDERS[currentProvider];
  localStorage.setItem('rb_provider', currentProvider);

  const keyRow = document.getElementById('apiKeyRow');
  const keyInput = document.getElementById('apiKeyInput');
  const getKeyLink = document.getElementById('getKeyLink');
  const modelLabel = document.getElementById('currentModelLabel');

  if (modelLabel) modelLabel.textContent = provider.label;

  if (provider.needsKey) {
    keyRow.classList.remove('rb-hidden');
    getKeyLink.href = provider.getKeyUrl;
    // Restore saved key
    const savedKey = localStorage.getItem(provider.keyStorageId);
    if (savedKey) {
      keyInput.value = '••••••••' + savedKey.slice(-4);
      document.getElementById('apiBanner').classList.add('saved');
    } else {
      keyInput.value = '';
      document.getElementById('apiBanner').classList.remove('saved');
    }
  } else {
    keyRow.classList.add('rb-hidden');
    document.getElementById('apiBanner').classList.add('saved');
  }
}

function saveApiKey() {
  const key = document.getElementById('apiKeyInput').value.trim();
  const provider = AI_PROVIDERS[currentProvider];

  if (!key || key.startsWith('••')) {
    showToast('Please enter a valid API key', 'error');
    return;
  }
  if (!provider.needsKey) {
    showToast('This model does not need an API key', 'info');
    return;
  }

  localStorage.setItem(provider.keyStorageId, key);
  document.getElementById('apiBanner').classList.add('saved');
  document.getElementById('apiKeyInput').value = '••••••••' + key.slice(-4);
  showToast('API key saved for ' + provider.label, 'success');
}

function getProviderKey(providerId) {
  const provider = AI_PROVIDERS[providerId || currentProvider];
  if (!provider || !provider.needsKey) return null;
  return localStorage.getItem(provider.keyStorageId) || '';
}

// ── Populate Regenerate Selectors ────────────────────────

function populateRegenSelectors() {
  const options = Object.entries(AI_PROVIDERS).map(function(entry) {
    var id = entry[0], p = entry[1];
    var tag = p.needsKey ? '' : ' [No Key]';
    return '<option value="' + id + '">' + p.label + tag + '</option>';
  }).join('');

  ['regenModelResume', 'regenModelCover'].forEach(function(selId) {
    var el = document.getElementById(selId);
    if (el) el.innerHTML = options;
  });
}

// ── Tab Switching ────────────────────────────────────────

function setupTabs() {
  document.querySelectorAll('.rb-tab').forEach(tab => {
    tab.addEventListener('click', () => switchTab(tab.dataset.tab));
  });
}

function switchTab(tabName) {
  document.querySelectorAll('.rb-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.rb-tab-content').forEach(c => c.classList.remove('active'));
  var activeTab = document.querySelector('.rb-tab[data-tab="' + tabName + '"]');
  var activeContent = document.getElementById('tab-' + tabName);
  if (activeTab) activeTab.classList.add('active');
  if (activeContent) activeContent.classList.add('active');
}

function setupPreviewTabs() {
  document.querySelectorAll('.rb-preview-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.rb-preview-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.rb-preview-pane').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById('preview-' + tab.dataset.preview).classList.add('active');
    });
  });
}

// ── Step Navigation ──────────────────────────────────────

function goToStep(step) {
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
  goToStepDirect(step);
}

function goToStepDirect(step) {
  document.querySelectorAll('.rb-section').forEach(s => s.classList.add('rb-hidden'));
  document.getElementById('section' + step).classList.remove('rb-hidden');
  document.querySelectorAll('.rb-step-dot').forEach(dot => {
    var dotStep = parseInt(dot.dataset.step);
    dot.classList.remove('active', 'completed');
    if (dotStep === step) dot.classList.add('active');
    else if (dotStep < step) dot.classList.add('completed');
  });
  currentStep = step;
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── File Upload & Parsing ────────────────────────────────

function setupDragDrop() {
  var zone = document.getElementById('uploadZone');
  var fileInput = document.getElementById('resumeFile');

  zone.addEventListener('dragover', function(e) { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', function() { zone.classList.remove('dragover'); });
  zone.addEventListener('drop', function(e) {
    e.preventDefault();
    zone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFileUpload(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', function(e) {
    if (e.target.files.length) handleFileUpload(e.target.files[0]);
  });
}

async function handleFileUpload(file) {
  var status = document.getElementById('uploadStatus');
  if (file.size > 10 * 1024 * 1024) {
    status.textContent = 'File too large. Max 10MB.';
    status.className = 'rb-upload-status error';
    return;
  }

  status.textContent = 'Parsing ' + file.name + '...';
  status.className = 'rb-upload-status success';

  try {
    var text = '';
    if (file.name.endsWith('.txt')) text = await file.text();
    else if (file.name.endsWith('.pdf')) text = await parsePDF(file);
    else if (file.name.endsWith('.docx') || file.name.endsWith('.doc')) text = await parseDOCX(file);
    else throw new Error('Unsupported format. Use .pdf, .docx, or .txt');

    fillFieldsFromText(text);
    status.textContent = 'Parsed "' + file.name + '" — fields populated. Check "Paste Details" tab.';
    status.className = 'rb-upload-status success';
    showToast('Resume parsed! Review in "Paste Details" tab.', 'success');
    setTimeout(function() { switchTab('paste'); }, 500);
  } catch (err) {
    status.textContent = 'Error: ' + err.message;
    status.className = 'rb-upload-status error';
  }
}

async function parsePDF(file) {
  var arrayBuffer = await file.arrayBuffer();
  var pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  var text = '';
  for (var i = 1; i <= pdf.numPages; i++) {
    var page = await pdf.getPage(i);
    var content = await page.getTextContent();
    text += content.items.map(function(item) { return item.str; }).join(' ') + '\n';
  }
  return text;
}

async function parseDOCX(file) {
  var arrayBuffer = await file.arrayBuffer();
  var result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
  return result.value;
}

function fillFieldsFromText(text) {
  var emailMatch = text.match(/[\w.+-]+@[\w-]+\.[\w.-]+/);
  var phoneMatch = text.match(/[\+]?[\d\s\-().]{10,}/);
  var linkedinMatch = text.match(/linkedin\.com\/in\/[\w-]+/i);
  var githubMatch = text.match(/github\.com\/[\w-]+/i);

  if (emailMatch) document.getElementById('email').value = emailMatch[0];
  if (phoneMatch) document.getElementById('phone').value = phoneMatch[0].trim();
  if (linkedinMatch) document.getElementById('linkedin').value = 'https://' + linkedinMatch[0];
  if (githubMatch) document.getElementById('github').value = 'https://' + githubMatch[0];
  document.getElementById('experience').value = text;
}

// ── Collect Data ─────────────────────────────────────────

function getProfileData() {
  var data = {
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

// ── Fetch Job from URL ───────────────────────────────────

async function fetchJobFromURL() {
  var urlInput = document.getElementById('jobUrl');
  var url = urlInput.value.trim();

  if (!url) {
    showToast('Please enter a job posting URL', 'error');
    urlInput.focus();
    return;
  }

  // Basic URL validation
  try { new URL(url); } catch (e) {
    showToast('Please enter a valid URL (starting with https://)', 'error');
    return;
  }

  var btn = document.getElementById('fetchUrlBtn');
  btn.querySelector('.rb-btn-text').classList.add('rb-hidden');
  btn.querySelector('.rb-btn-loading').classList.remove('rb-hidden');
  btn.disabled = true;

  showToast('Fetching job posting...', 'info');

  try {
    // Use a public CORS proxy to fetch the page content
    var proxyUrls = [
      'https://api.allorigins.win/raw?url=' + encodeURIComponent(url),
      'https://corsproxy.io/?' + encodeURIComponent(url),
    ];

    var html = '';
    var fetched = false;
    for (var i = 0; i < proxyUrls.length; i++) {
      try {
        var resp = await fetch(proxyUrls[i], { signal: AbortSignal.timeout(15000) });
        if (resp.ok) {
          html = await resp.text();
          fetched = true;
          break;
        }
      } catch (proxyErr) {
        continue;
      }
    }

    if (!fetched || !html) {
      throw new Error('Could not fetch the page. The site may block external access. Please copy-paste the job description manually.');
    }

    // Parse HTML and extract text content
    var parser = new DOMParser();
    var doc = parser.parseFromString(html, 'text/html');

    // Remove noise elements
    doc.querySelectorAll('script, style, nav, header, footer, iframe, noscript, svg, img, link, meta').forEach(function(el) { el.remove(); });

    // Try to find the main job content by common selectors
    var jobContent = '';
    var selectors = [
      '.job-description', '.jobs-description', '.job-details',
      '.description__text', '.show-more-less-html', '.job-view-layout',
      '[data-job-description]', '.posting-requirements', '.job-posting',
      'article', 'main', '[role="main"]',
    ];

    for (var s = 0; s < selectors.length; s++) {
      var el = doc.querySelector(selectors[s]);
      if (el && el.innerText && el.innerText.trim().length > 100) {
        jobContent = el.innerText.trim();
        break;
      }
    }

    // Fallback: get body text
    if (!jobContent) {
      jobContent = (doc.body && doc.body.innerText) || '';
    }

    // Clean up excessive whitespace
    jobContent = jobContent.replace(/\n{3,}/g, '\n\n').replace(/[ \t]{2,}/g, ' ').trim();

    if (jobContent.length < 50) {
      throw new Error('Could not extract enough content from the page. Please copy-paste the job description manually.');
    }

    // Truncate if extremely long (keep first 8000 chars)
    if (jobContent.length > 8000) {
      jobContent = jobContent.substring(0, 8000) + '\n\n[Content truncated — edit if needed]';
    }

    // Fill the textarea
    document.getElementById('jobDescription').value = jobContent;

    // Try to auto-detect title and company from the page
    var pageTitle = doc.querySelector('title');
    if (pageTitle) {
      var titleText = pageTitle.textContent || '';
      // Common patterns: "Job Title at Company" or "Job Title - Company"
      var titleMatch = titleText.match(/^(.+?)\s*(?:at|[-–|@])\s*(.+?)(?:\s*[-–|]|$)/i);
      if (titleMatch) {
        if (!document.getElementById('jobTitle').value) {
          document.getElementById('jobTitle').value = titleMatch[1].trim();
        }
        if (!document.getElementById('companyName').value) {
          document.getElementById('companyName').value = titleMatch[2].trim();
        }
      }
    }

    showToast('Job posting fetched! Review the extracted content below.', 'success');
  } catch (err) {
    showError('Fetch Failed', err.message);
  } finally {
    btn.querySelector('.rb-btn-text').classList.remove('rb-hidden');
    btn.querySelector('.rb-btn-loading').classList.add('rb-hidden');
    btn.disabled = false;
  }
}

// ══════════════════════════════════════════════════════════
// AI PROVIDER IMPLEMENTATIONS
// ══════════════════════════════════════════════════════════

// ── HuggingFace Free Inference (No key needed) ───────────

async function callHuggingFaceFree(prompt, providerId) {
  var provider = AI_PROVIDERS[providerId];
  // Use the HuggingFace Inference API with router endpoint for better CORS support
  var url = 'https://router.huggingface.co/hf-inference/models/' + provider.model + '/v1/chat/completions';

  var resp;
  try {
    resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: provider.model,
        messages: [
          { role: 'system', content: 'You are an expert resume and cover letter writer. Output clean HTML only. No code blocks.' },
          { role: 'user', content: prompt }
        ],
        max_tokens: 3000,
        temperature: 0.7,
      }),
    });
  } catch (networkErr) {
    throw new Error(
      'Network error connecting to HuggingFace. This may be due to browser restrictions.\n\n' +
      'FIX: Switch to Groq (recommended) — get a free key at console.groq.com/keys'
    );
  }

  if (!resp.ok) {
    var err = await resp.json().catch(function() { return {}; });
    if (resp.status === 503) {
      throw new Error('Model is loading — please wait 30s and try again, or switch to Groq (recommended).');
    }
    if (resp.status === 429) {
      throw new Error('Rate limited by HuggingFace. Switch to Groq for unlimited free usage.');
    }
    throw new Error(err.error || 'HuggingFace returned ' + resp.status + '. Try switching to Groq.');
  }

  var data = await resp.json();
  // OpenAI-compatible response format
  if (data.choices && data.choices[0] && data.choices[0].message) {
    return extractHTML(data.choices[0].message.content);
  }
  // Legacy format fallback
  if (Array.isArray(data) && data[0] && data[0].generated_text) {
    return extractHTML(data[0].generated_text);
  }
  throw new Error('Unexpected response from HuggingFace. Try switching to Groq.');
}

// ── Groq (Free key — Llama, Mixtral, DeepSeek) ──────────

async function callGroq(prompt, providerId) {
  var provider = AI_PROVIDERS[providerId];
  var apiKey = getProviderKey(providerId);
  if (!apiKey) throw new Error('Groq API key required (free, takes 10 seconds).\n\n1. Go to console.groq.com/keys\n2. Sign up with Google/GitHub\n3. Create an API key\n4. Paste it above and click Save Key');

  var resp;
  try {
    resp = await fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + apiKey,
    },
    body: JSON.stringify({
      model: provider.model,
      messages: [
        { role: 'system', content: 'You are an expert resume and cover letter writer. Output clean HTML only. No code blocks.' },
        { role: 'user', content: prompt },
      ],
      temperature: 0.7,
      max_tokens: 4096,
    }),
  });
  } catch (networkErr) {
    throw new Error('Network error connecting to Groq. Check your internet connection and try again.');
  }

  if (!resp.ok) {
    var err = await resp.json().catch(function() { return {}; });
    if (resp.status === 401) {
      throw new Error('Invalid Groq API key. Please check your key and save it again.');
    }
    throw new Error((err.error && err.error.message) || 'Groq returned ' + resp.status);
  }

  var data = await resp.json();
  var text = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
  return extractHTML(text);
}

// ── Google Gemini (Free key) ─────────────────────────────

async function callGemini(prompt, providerId) {
  var apiKey = getProviderKey(providerId);
  if (!apiKey) throw new Error('Gemini API key required. Get one free at aistudio.google.com/app/apikey — or switch to a free model above.');

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + encodeURIComponent(apiKey);

  var resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 4096 },
    }),
  });

  if (!resp.ok) {
    var err = await resp.json().catch(function() { return {}; });
    throw new Error((err.error && err.error.message) || 'Gemini returned ' + resp.status);
  }

  var data = await resp.json();
  var text = (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0] && data.candidates[0].content.parts[0].text) || '';
  return extractHTML(text);
}

// ── OpenRouter (Free key, free models) ───────────────────

async function callOpenRouter(prompt, providerId) {
  var provider = AI_PROVIDERS[providerId];
  var apiKey = getProviderKey(providerId);
  if (!apiKey) throw new Error('OpenRouter API key required. Get one free at openrouter.ai/keys — or switch to a free model above.');

  var resp = await fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + apiKey,
      'HTTP-Referer': window.location.origin,
      'X-Title': 'AI Resume Builder',
    },
    body: JSON.stringify({
      model: provider.model,
      messages: [
        { role: 'system', content: 'You are an expert resume and cover letter writer. Output clean HTML only. No code blocks.' },
        { role: 'user', content: prompt },
      ],
      temperature: 0.7,
      max_tokens: 4096,
    }),
  });

  if (!resp.ok) {
    var err = await resp.json().catch(function() { return {}; });
    throw new Error((err.error && err.error.message) || 'OpenRouter returned ' + resp.status);
  }

  var data = await resp.json();
  var text = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
  return extractHTML(text);
}

// ── Unified AI caller ────────────────────────────────────

async function callAI(prompt, providerId) {
  providerId = providerId || currentProvider;
  var provider = AI_PROVIDERS[providerId];
  if (!provider) throw new Error('Unknown provider: ' + providerId);

  if (provider.needsKey && !getProviderKey(providerId)) {
    throw new Error(
      'API key required for ' + provider.label + '. ' +
      'Get a free key at: ' + provider.getKeyUrl + '\n' +
      'Or switch to a free model (no key needed) in the provider bar above.'
    );
  }

  return provider.call(prompt, providerId);
}

// ── Extract HTML from AI response ────────────────────────

function extractHTML(text) {
  // Try ```html code block first
  var htmlBlock = text.match(/```html\s*([\s\S]*?)```/);
  if (htmlBlock) return htmlBlock[1].trim();

  // Try any code block
  var codeBlock = text.match(/```\s*([\s\S]*?)```/);
  if (codeBlock) return codeBlock[1].trim();

  // If it looks like HTML, return as-is
  if (text.indexOf('<') !== -1 && text.indexOf('>') !== -1) return text.trim();

  // Convert plain text to HTML paragraphs
  return text.split('\n').filter(function(l) { return l.trim(); }).map(function(l) { return '<p>' + l + '</p>'; }).join('\n');
}

// ══════════════════════════════════════════════════════════
// GENERATION
// ══════════════════════════════════════════════════════════

async function generateResume() {
  var profile = getProfileData();
  var job = getJobData();

  if (!job.description) {
    showToast('Please paste the job description', 'error');
    return;
  }

  var btn = document.getElementById('generateBtn');
  btn.querySelector('.rb-btn-text').classList.add('rb-hidden');
  btn.querySelector('.rb-btn-loading').classList.remove('rb-hidden');
  btn.disabled = true;

  var provider = AI_PROVIDERS[currentProvider];
  showToast('Generating with ' + provider.label + '...', 'info');

  hideError();

  try {
    // Generate sequentially to avoid rate limits on free APIs
    var resumeHtml = await callAI(buildResumePrompt(profile, job));
    generatedResume = resumeHtml;
    document.getElementById('resumePreview').innerHTML = sanitizeHTML(resumeHtml);

    var coverHtml = await callAI(buildCoverLetterPrompt(profile, job));
    generatedCover = coverHtml;
    document.getElementById('coverPreview').innerHTML = sanitizeHTML(coverHtml);

    goToStepDirect(3);
    showToast('Generated with ' + provider.label + '! Not happy? Pick another model and regenerate.', 'success');
  } catch (err) {
    showError('Generation Failed', err.message);
  } finally {
    btn.querySelector('.rb-btn-text').classList.remove('rb-hidden');
    btn.querySelector('.rb-btn-loading').classList.add('rb-hidden');
    btn.disabled = false;
  }
}

async function regenerateWithModel(section) {
  var selectId = section === 'resume' ? 'regenModelResume' : 'regenModelCover';
  var providerId = document.getElementById(selectId).value;
  var provider = AI_PROVIDERS[providerId];

  if (!provider) { showToast('Select a model first', 'error'); return; }

  if (provider.needsKey && !getProviderKey(providerId)) {
    showToast('API key required for ' + provider.label + '. Get one free at: ' + provider.getKeyUrl + ' — or pick a free model.', 'error');
    return;
  }

  var profile = getProfileData();
  var job = getJobData();

  showToast('Regenerating ' + section + ' with ' + provider.label + '...', 'info');

  try {
    var prompt = section === 'resume'
      ? buildResumePrompt(profile, job)
      : buildCoverLetterPrompt(profile, job);

    var html = await callAI(prompt, providerId);
    var previewEl = document.getElementById(section === 'resume' ? 'resumePreview' : 'coverPreview');
    previewEl.innerHTML = sanitizeHTML(html);

    if (section === 'resume') generatedResume = html;
    else generatedCover = html;

    showToast(section + ' regenerated with ' + provider.label + '!', 'success');
  } catch (err) {
    showError('Regeneration Failed', err.message);
  }
}

// ── Prompt Builders ──────────────────────────────────────

function buildResumePrompt(profile, job) {
  return 'You are an expert resume writer. Generate a professional, ATS-friendly resume in clean HTML format.\n\n' +
    'CANDIDATE INFORMATION:\n' +
    '- Name: ' + (profile.fullName || 'Not provided') + '\n' +
    '- Email: ' + (profile.email || 'Not provided') + '\n' +
    '- Phone: ' + (profile.phone || 'Not provided') + '\n' +
    '- Location: ' + (profile.location || 'Not provided') + '\n' +
    '- LinkedIn: ' + (profile.linkedin || 'Not provided') + '\n' +
    '- GitHub: ' + (profile.github || 'Not provided') + '\n' +
    '- Summary: ' + (profile.summary || 'Not provided') + '\n' +
    '- Experience: ' + (profile.experience || 'Not provided') + '\n' +
    '- Skills: ' + (profile.skills || 'Not provided') + '\n' +
    '- Education: ' + (profile.education || 'Not provided') + '\n' +
    '- Certifications: ' + (profile.certifications || 'Not provided') + '\n\n' +
    'TARGET ROLE:\n' +
    '- Job Title: ' + (job.title || 'Not specified') + '\n' +
    '- Company: ' + (job.company || 'Not specified') + '\n' +
    '- Job Description: ' + job.description + '\n\n' +
    (job.notes ? 'ADDITIONAL INSTRUCTIONS: ' + job.notes + '\n\n' : '') +
    'REQUIREMENTS:\n' +
    '1. Output ONLY clean HTML (no <html>, <head>, <body> tags — just resume content).\n' +
    '2. Use a centered header with name, contact info in a <div class="resume-header">.\n' +
    '3. Use <p class="resume-contact"> for contact details.\n' +
    '4. Use <h2> for section headers: PROFESSIONAL SUMMARY, EXPERIENCE, SKILLS, EDUCATION, CERTIFICATIONS.\n' +
    '5. Tailor the experience and summary to match the job description.\n' +
    '6. Use strong action verbs and quantified achievements.\n' +
    '7. Highlight keywords from the job description naturally.\n' +
    '8. Keep it to 1-2 pages worth of content.\n' +
    '9. Use <ul> with <li> for bullet points.\n' +
    '10. Do NOT include any CSS, style tags, or code block markers.';
}

function buildCoverLetterPrompt(profile, job) {
  return 'You are an expert cover letter writer. Generate a professional cover letter in clean HTML format.\n\n' +
    'CANDIDATE INFORMATION:\n' +
    '- Name: ' + (profile.fullName || 'Not provided') + '\n' +
    '- Email: ' + (profile.email || 'Not provided') + '\n' +
    '- Phone: ' + (profile.phone || 'Not provided') + '\n' +
    '- Location: ' + (profile.location || 'Not provided') + '\n' +
    '- Summary: ' + (profile.summary || 'Not provided') + '\n' +
    '- Key Experience: ' + (profile.experience ? profile.experience.substring(0, 2000) : 'Not provided') + '\n' +
    '- Skills: ' + (profile.skills || 'Not provided') + '\n\n' +
    'TARGET ROLE:\n' +
    '- Job Title: ' + (job.title || 'Not specified') + '\n' +
    '- Company: ' + (job.company || 'Not specified') + '\n' +
    '- Job Description: ' + job.description + '\n\n' +
    (job.notes ? 'ADDITIONAL INSTRUCTIONS: ' + job.notes + '\n\n' : '') +
    'REQUIREMENTS:\n' +
    '1. Output ONLY clean HTML (no <html>, <head>, <body> tags).\n' +
    '2. Format as a business letter: date, greeting, 3-4 paragraphs, closing.\n' +
    '3. Connect experience to key JD requirements with specific examples.\n' +
    '4. Be genuine, confident, specific — no generic filler.\n' +
    '5. Keep under 400 words. Use <p> tags.\n' +
    '6. Do NOT include any CSS, style tags, or code block markers.';
}

// ── HTML Sanitizer ───────────────────────────────────────

function sanitizeHTML(html) {
  var temp = document.createElement('div');
  temp.innerHTML = html;
  temp.querySelectorAll('script, style, iframe, object, embed, form, input, textarea, select, button, link, meta')
    .forEach(function(el) { el.remove(); });
  temp.querySelectorAll('*').forEach(function(el) {
    Array.from(el.attributes).forEach(function(attr) {
      if (attr.name.startsWith('on') || attr.name === 'srcdoc') el.removeAttribute(attr.name);
      if (attr.name === 'href' || attr.name === 'src' || attr.name === 'action') {
        var val = attr.value.trim().toLowerCase();
        if (val.startsWith('javascript:') || val.startsWith('data:text/html')) el.removeAttribute(attr.name);
      }
    });
  });
  return temp.innerHTML;
}

// ── Edit Preview ─────────────────────────────────────────

function editPreview(section) {
  var paper = document.getElementById(section === 'resume' ? 'resumePreview' : 'coverPreview');
  var isEditing = paper.contentEditable === 'true';
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
  var activePane = document.querySelector('.rb-preview-pane.active');
  var paper = activePane.querySelector('.rb-preview-paper');
  var type = activePane.id.indexOf('resume') !== -1 ? 'Resume' : 'Cover_Letter';
  var name = document.getElementById('fullName').value.trim() || 'Document';

  showToast('Generating PDF...', 'info');
  html2pdf().set({
    margin: 0.5,
    filename: name.replace(/\s+/g, '_') + '_' + type + '.pdf',
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true },
    jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' },
  }).from(paper).save().then(function() { showToast('PDF downloaded!', 'success'); });
}

// ── Export: DOCX ─────────────────────────────────────────

function exportDOCX() {
  var activePane = document.querySelector('.rb-preview-pane.active');
  var paper = activePane.querySelector('.rb-preview-paper');
  var type = activePane.id.indexOf('resume') !== -1 ? 'Resume' : 'Cover_Letter';
  var name = document.getElementById('fullName').value.trim() || 'Document';

  showToast('Generating DOCX...', 'info');
  try {
    var lines = paper.innerText.split('\n').filter(function(l) { return l.trim(); });
    var children = [];
    lines.forEach(function(line) {
      var trimmed = line.trim();
      if (!trimmed) return;
      if (trimmed === trimmed.toUpperCase() && trimmed.length < 50 && trimmed.length > 2) {
        children.push(new docx.Paragraph({ text: trimmed, heading: docx.HeadingLevel.HEADING_2, spacing: { before: 300, after: 100 } }));
      } else if (/^[•\-–]/.test(trimmed)) {
        children.push(new docx.Paragraph({ text: trimmed.replace(/^[•\-–]\s*/, ''), bullet: { level: 0 } }));
      } else {
        children.push(new docx.Paragraph({ text: trimmed, spacing: { after: 80 } }));
      }
    });
    var doc = new docx.Document({
      sections: [{ properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } }, children: children }],
    });
    docx.Packer.toBlob(doc).then(function(blob) {
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = name.replace(/\s+/g, '_') + '_' + type + '.docx';
      a.click();
      URL.revokeObjectURL(url);
      showToast('DOCX downloaded!', 'success');
    });
  } catch (err) {
    showToast('DOCX error: ' + err.message, 'error');
  }
}

// ── Export: Clipboard ────────────────────────────────────

function copyToClipboard() {
  var activePane = document.querySelector('.rb-preview-pane.active');
  var paper = activePane.querySelector('.rb-preview-paper');
  navigator.clipboard.writeText(paper.innerText).then(function() {
    showToast('Copied to clipboard!', 'success');
  }).catch(function() {
    var range = document.createRange();
    range.selectNodeContents(paper);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('copy');
    sel.removeAllRanges();
    showToast('Copied to clipboard!', 'success');
  });
}

// ── Toast ────────────────────────────────────────────────

function showToast(message, type) {
  type = type || 'info';
  document.querySelectorAll('.rb-toast').forEach(function(t) { t.remove(); });
  var toast = document.createElement('div');
  toast.className = 'rb-toast ' + type;
  toast.textContent = message;
  document.body.appendChild(toast);
  requestAnimationFrame(function() { toast.classList.add('show'); });
  var duration = type === 'error' ? 8000 : 4000;
  setTimeout(function() {
    toast.classList.remove('show');
    setTimeout(function() { toast.remove(); }, 300);
  }, duration);
}

// ── Inline Error Banner ──────────────────────────────────

function showError(title, message) {
  var banner = document.getElementById('errorBanner');
  var titleEl = document.getElementById('errorTitle');
  var msgEl = document.getElementById('errorMessage');
  if (!banner) { showToast(title + ': ' + message, 'error'); return; }
  titleEl.textContent = title;
  msgEl.textContent = message;
  banner.classList.remove('rb-hidden');
  banner.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  showToast(title + ' — see details above', 'error');
}

function hideError() {
  var banner = document.getElementById('errorBanner');
  if (banner) banner.classList.add('rb-hidden');
}
