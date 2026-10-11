/**
 * Volunteer Drafting Studio & Template Logic Validator (Smart Field Sync Edition)
 */

// ============================================================================
// 1. CONFIGURATION & STATE
// ============================================================================
const ACCESS_PASSWORD = "Habeas123!";
let currentArrayBuffer = null;
let primaryFormReady;
let activeDocx = null;
let templateReadSerial = 0;
function updateUploadIndicator() {
  const name = document.getElementById('upload-template-name');
  const detail = document.getElementById('upload-template-detail');
  const hint = document.getElementById('upload-template-hint');
  if (!name || !detail || !hint) return;
  name.textContent = activeDocx ? `Loaded: ${activeDocx.name}` : 'Drag & Drop your .docx template here';
  detail.hidden = !activeDocx;
  detail.textContent = activeDocx ? `${activeDocx.fields} ${activeDocx.fields === 1 ? 'field' : 'fields'} detected · Template syntax checked` : '';
  hint.textContent = activeDocx ? 'Click here or drop another Word document to replace it' : 'or click to browse from your computer';
}

function identifyTemplate(snippet = false) {
  updateUploadIndicator();
  const area = document.getElementById('active-template'); if (!area) return;
  area.replaceChildren();
  const title = document.createElement('strong'); title.textContent = activeDocx ? `Active template: ${activeDocx.name}` : 'No active DOCX template.'; area.append(title);
  if (activeDocx) {
    const detail = document.createElement('p'); detail.textContent = `${activeDocx.fields} ${activeDocx.fields === 1 ? 'field' : 'fields'} detected · Template syntax checked · Word output not yet verified.`; area.append(detail);
  }
  if (snippet) { const note = document.createElement('p'); note.textContent = 'Sandbox source: pasted text snippet.'; area.append(note); }
}

// ============================================================================
// 2. AUTHENTICATION & UI EVENTS
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
  const overlay = document.getElementById('auth-overlay');
  const pinInput = document.getElementById('accessPin');
  const unlockBtn = document.getElementById('unlockBtn');
  const togglePwBtn = document.getElementById('togglePwBtn');
  const authError = document.getElementById('authError');
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const renderBtn = document.getElementById('renderBtn');
  const downloadBtn = document.getElementById('downloadBtn');

  const inspectSnippetBtn = document.getElementById('inspectSnippetBtn');
  const snippetInput = document.getElementById('snippetInput');

  // 1. Snippet Input & Button Listener
  if (inspectSnippetBtn && snippetInput) {
    snippetInput.addEventListener('input', () => {
      inspectSnippetBtn.disabled = !snippetInput.value.trim();
    });

    inspectSnippetBtn.addEventListener('click', () => {
      const text = snippetInput.value;
      if (!text.trim()) {
        if (typeof showToast === 'function') {
          showToast("Please paste some text into the box to inspect.", "error");
        }
        return;
      }

      // Reset state for new inspection
      const resultsDiv = document.getElementById('results');
      const previewContainer = document.getElementById('docx-preview-container');
      const sandboxDiv = document.getElementById('sandbox');
      if (resultsDiv) resultsDiv.innerHTML = "";
      if (previewContainer) previewContainer.innerHTML = "";
      if (sandboxDiv) sandboxDiv.style.display = "none";

      ++templateReadSerial;
      inspectTemplate(createDocxFromText(text), "Pasted_Snippet.docx", 'snippet');
    });
  }

  primaryFormReady = loadPrimaryFormFromIndex();

  if (sessionStorage.getItem('drafting_studio_auth') === 'true') {
    if (overlay) overlay.style.display = 'none';
  }

  function verifyAccess() {
    if (!pinInput) return;
    const entered = pinInput.value.trim();
    if (entered.toLowerCase() === ACCESS_PASSWORD.toLowerCase()) {
      if (overlay) overlay.style.display = 'none';
      sessionStorage.setItem('drafting_studio_auth', 'true');
    } else {
      if (authError) authError.style.display = 'block';
    }
  }

  function togglePasswordVisibility() {
    if (!pinInput || !togglePwBtn) return;
    if (pinInput.type === 'password') {
      pinInput.type = 'text';
      togglePwBtn.textContent = '🙈';
    } else {
      pinInput.type = 'password';
      togglePwBtn.textContent = '👁️';
    }
  }

  if (unlockBtn) unlockBtn.addEventListener('click', verifyAccess);
  if (togglePwBtn) togglePwBtn.addEventListener('click', togglePasswordVisibility);
  if (pinInput) {
    pinInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') verifyAccess();
    });
  }

  if (renderBtn) renderBtn.addEventListener('click', () => runTestRender(false));
  if (downloadBtn) downloadBtn.addEventListener('click', () => runTestRender(true));

  // 2. Dropzone & File Input Handlers
  if (dropzone && fileInput) {
    dropzone.addEventListener('click', () => {
      fileInput.value = '';
      fileInput.click();
    });

    dropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        dropzone.click();
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length) {
        handleFile(e.target.files[0]);
      }
    });

    ['dragenter', 'dragover'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'dragend'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');

      let droppedFile = null;
      if (e.dataTransfer) {
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
          droppedFile = e.dataTransfer.files[0];
        } else if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
          const item = e.dataTransfer.items[0];
          if (item.kind === 'file') {
            droppedFile = item.getAsFile();
          }
        }
      }

      if (droppedFile) {
        handleFile(droppedFile);
      } else {
        if (typeof showToast === 'function') {
          showToast("Could not capture file from drag. Click the box to browse instead.", "info");
        }
      }
    });
  }
});

function handleFile(file) {
  if (!file) return;
  const serial = ++templateReadSerial;
  activeDocx = null; currentArrayBuffer = null; identifyTemplate();
  document.getElementById('results').replaceChildren();
  document.getElementById('sandbox').style.display = 'none';

  if (!file.name.toLowerCase().endsWith('.docx')) {
    if (typeof showToast === 'function') {
      showToast('Please select a Microsoft Word .docx document.', 'error');
    } else {
      alert('Please select a Microsoft Word .docx document.');
    }
    return;
  }

  // Clear out snippet text box and reset button
  const snippetInput = document.getElementById('snippetInput');
  const inspectSnippetBtn = document.getElementById('inspectSnippetBtn');
  if (snippetInput) snippetInput.value = '';
  if (inspectSnippetBtn) inspectSnippetBtn.disabled = true;

  const reader = new FileReader();
  reader.onload = function(e) {
    if (serial !== templateReadSerial) return;
    inspectTemplate(e.target.result, file.name);
  };
  reader.onerror = function() {
    if (serial !== templateReadSerial) return;
    activeDocx = null; currentArrayBuffer = null; identifyTemplate();
    if (typeof showToast === 'function') {
      showToast('Failed to read file from disk.', 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================================
// 3. AUTO-SYNC: CLONE PINNED PRIMARY FORM FROM index.html
// ============================================================================
async function loadPrimaryFormFromIndex() {
  const container = document.getElementById('sandbox-form-container');
  if (!container) return;

  try {
    const res = await fetch('index.html');
    if (!res.ok) throw new Error('Could not fetch index.html');
    
    const htmlText = await res.text();
    const doc = new DOMParser().parseFromString(htmlText, 'text/html');

    const primarySection = doc.querySelector('.primary-fields-section') || doc.body;
    const clone = primarySection.cloneNode(true);

    clone.querySelectorAll('button, #docx-preview-container, .preview-section, [type="submit"], h2, #dynamic-fields-section, .header-row').forEach(el => el.remove());

    container.innerHTML = "";
    container.appendChild(clone);

    // Bind Sandbox Jurisdiction Change to Facility Dropdown
    const sandboxJurisdiction = container.querySelector('#jurisdiction');
    const sandboxFacility = container.querySelector('#detention_facility');
    if (sandboxJurisdiction && sandboxFacility) {
      DocxEngine.updateFacilityDropdown(sandboxFacility, sandboxJurisdiction.value);
      sandboxJurisdiction.addEventListener('change', (e) => {
        DocxEngine.updateFacilityDropdown(sandboxFacility, e.target.value);
      });
    }
  } catch (err) {
    console.warn("Auto-sync fallback: Using default primary fields", err);
    container.innerHTML = `
      <div class="sandbox-grid">
        <div class="field"><label>Jurisdiction</label><select id="jurisdiction" data-form-field="Jurisdiction"><option value="Massachusetts">Massachusetts</option><option value="New Hampshire">New Hampshire</option></select></div>
        <div class="field"><label>Detention Facility</label><select id="detention_facility" data-form-field="Detention Facility"><option value="Plymouth">Plymouth</option><option value="Burlington">Burlington</option></select></div>
        <div class="field"><label>Petitioner Name</label><input type="text" id="petitioner_first_name" data-form-field="Petitioner First Name" value="Jane Marie Doe"></div>
        <div class="field"><label>Case Number</label><input type="text" id="case_number" data-form-field="Case Number" value="2026-CV-04321"></div>
      </div>
    `;
  }
}

// ============================================================================
// 4. TEMPLATE LOGIC SCANNER & DYNAMIC FIELD DISCOVERY
// ============================================================================
function inspectTemplate(buffer, filename, source = 'docx') {
  const results = document.getElementById('results'), sandbox = document.getElementById('sandbox');
  results.replaceChildren(); sandbox.style.display = 'none';
  currentArrayBuffer = null;
  if (source === 'docx') activeDocx = null;
  identifyTemplate(source === 'snippet');
  try {
    const model = DocxEngine.inspect(buffer);
    const badge = document.createElement('div'); badge.className = 'badge-pass';
    badge.textContent = `${filename}: template syntax checked. ${model.fields.length} ${model.fields.length === 1 ? 'field' : 'fields'} found. Generate a test document and review it in Microsoft Word before use.`;
    results.append(badge);
    renderDynamicFields(model.fields, document.getElementById('sandbox-dynamic-container'), document.getElementById('sandbox-dynamic-section'), true);
    currentArrayBuffer = buffer;
    if (source === 'docx') activeDocx = { name: filename.split(/[\\/]/).pop(), fields: model.fields.length, buffer };
    identifyTemplate(source === 'snippet');
    sandbox.style.display = 'block';
  } catch (e) {
    const card = document.createElement('div'); card.className = 'error-card';
    const title = document.createElement('h4'); title.textContent = 'Template error:';
    const message = document.createElement('p'); message.textContent = e.message; card.append(title, message);
    if (e.context) { const snippet = document.createElement('div'); snippet.className = 'snippet-box'; snippet.textContent = e.context; card.append(snippet); }
    const fix = document.createElement('p'); fix.textContent = 'If the error identifies a field or instruction, locate it in Word, correct it, and inspect again.'; card.append(fix); results.append(card);
  }
}

// ============================================================================
// 5. LIVE TEST SANDBOX RENDERER (Delegates compilation to DocxEngine)
// ============================================================================
async function runTestRender(download) {
  if (!currentArrayBuffer) { showToast('Upload a document or inspect a snippet first.', 'error'); return; }
  await primaryFormReady;
  return GenerationUI.run(download, { buffer: currentArrayBuffer, collect: getFormData });
}

/**
 * Wraps raw pasted snippet text into an in-memory DOCX ArrayBuffer
 */
function createDocxFromText(rawText) {
  const zip = new PizZip();
  const xmlBody = rawText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .split(/\r\n|\r|\n/)
    .map(line => `<w:p><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`)
    .join('');

  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${xmlBody}<w:sectPr/></w:body></w:document>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

  zip.file("_rels/.rels", rels);
  zip.file("[Content_Types].xml", contentTypes);
  zip.file("word/document.xml", docXml);

  return zip.generate({ type: "arraybuffer" });
}