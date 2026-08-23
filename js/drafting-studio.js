/**
 * Volunteer Drafting Studio & Template Logic Validator (Smart Field Sync Edition)
 */

// ============================================================================
// 1. CONFIGURATION & STATE
// ============================================================================
const ACCESS_PASSWORD = "Habeas123!";
let currentArrayBuffer = null;

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
      const text = snippetInput.value.trim();
      if (!text) {
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

      currentArrayBuffer = createDocxFromText(text);
      inspectTemplate(currentArrayBuffer, "Pasted_Snippet.docx");
    });
  }

  loadPrimaryFormFromIndex();

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
    currentArrayBuffer = e.target.result;
    inspectTemplate(currentArrayBuffer, file.name);
  };
  reader.onerror = function() {
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
        <div class="field"><label>Jurisdiction</label><select id="jurisdiction"><option value="Massachusetts">Massachusetts</option><option value="New Hampshire">New Hampshire</option></select></div>
        <div class="field"><label>Detention Facility</label><select id="detention_facility"><option value="Plymouth">Plymouth</option><option value="Burlington">Burlington</option></select></div>
        <div class="field"><label>Petitioner Name</label><input type="text" id="Petitioner Name" value="Jane Marie Doe"></div>
        <div class="field"><label>Case Number</label><input type="text" id="Case Number" value="2026-CV-04321"></div>
      </div>
    `;
  }
}

// ============================================================================
// 4. TEMPLATE LOGIC SCANNER & DYNAMIC FIELD DISCOVERY
// ============================================================================
function inspectTemplate(buffer, filename) {
  const resultsDiv = document.getElementById('results');
  const sandboxDiv = document.getElementById('sandbox');
  const dynamicSection = document.getElementById('sandbox-dynamic-section');
  const dynamicContainer = document.getElementById('sandbox-dynamic-container');

  resultsDiv.innerHTML = "";
  sandboxDiv.style.display = "none";
  if (dynamicSection) dynamicSection.style.display = "none";

  const previewContainer = document.getElementById('docx-preview-container');
  if (previewContainer) previewContainer.innerHTML = "";

  try {
    const zip = new PizZip(buffer);
    let xmlContent = "";
    
    Object.keys(zip.files).forEach(name => {
      if (name.startsWith("word/") && name.endsWith(".xml")) {
        xmlContent += zip.files[name].asText() + "\n";
      }
    });

    const cleanDocText = xmlContent
      .replace(/<\/w:p>/g, '\n')
      .replace(/\[[^\]]*?\]/g, m => m.replace(/<[^>]+>/g, ''))
      .replace(/<[^>]+>/g, '')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");

    function getSnippet(fullText, index, tagLen) {
      const start = Math.max(0, index - 50);
      const end = Math.min(fullText.length, index + tagLen + 50);
      let before = fullText.substring(start, index).replace(/\s+/g, ' ');
      let tag = fullText.substring(index, index + tagLen);
      let after = fullText.substring(index + tagLen, end).replace(/\s+/g, ' ');
      
      if (start > 0) before = "..." + before;
      if (end < fullText.length) after = after + "...";
      return { before, tag, after };
    }

    const errors = [];
    const stack = [];
    const tagRegex = /\[\s*(IF\s+[^\]]+|ELSE\s+IF\s+[^\]]+|ELSEIF\s+[^\]]+|ELSE|END\s+IF|[^\]]+)\s*\]/gi;
    let match;

    while ((match = tagRegex.exec(cleanDocText)) !== null) {
      const rawTag = match[0];
      const body = match[1].trim();
      const matchIndex = match.index;
      const snippet = getSnippet(cleanDocText, matchIndex, rawTag.length);

      if (!isTagSyntaxValid(body)) {
        errors.push({
          title: "Unbalanced Quotation Marks",
          tag: rawTag,
          snippet: snippet,
          fix: `In Word, search for this tag with <b>Ctrl+F</b> and ensure all quotes are paired. Example: <code>[IF Jurisdiction = "Massachusetts"]</code>.`
        });
      }

      if (/^IF\s+/i.test(body)) {
        const cond = body.replace(/^IF\s+/i, '').trim();
        if (!cond) {
          errors.push({
            title: "Empty [IF] Tag",
            tag: rawTag,
            snippet: snippet,
            fix: `Add a condition inside the bracket, such as <code>[IF Jurisdiction = "Massachusetts"]</code>.`
          });
        }
        stack.push({ tag: rawTag, cond: cond, hasElse: false, snippet: snippet });
      }
      else if (/^(ELSE\s*IF|ELSEIF)\s*/i.test(body)) {
        const cond = body.replace(/^(ELSE\s*IF|ELSEIF)\s*/i, '').trim();
        if (stack.length === 0) {
          errors.push({
            title: "Orphan [ELSE IF] Tag",
            tag: rawTag,
            snippet: snippet,
            fix: `This [ELSE IF] has no opening [IF] statement above it. Add an <code>[IF ...]</code> tag before this line or remove this tag.`
          });
        } else {
          const top = stack[stack.length - 1];
          if (top.hasElse) {
            errors.push({
              title: "Invalid Tag Order ([ELSE IF] after [ELSE])",
              tag: rawTag,
              snippet: snippet,
              fix: `In your Word template, move this <code>${rawTag}</code> tag <b>ABOVE</b> the preceding <code>[ELSE]</code> tag. [ELSE] must always be the final branch before [END IF].`
            });
          }
          if (!cond) {
            errors.push({
              title: "Empty [ELSE IF] Condition",
              tag: rawTag,
              snippet: snippet,
              fix: `Add a condition to this [ELSE IF] tag, such as <code>[ELSE IF Jurisdiction = "Rhode Island"]</code>.`
            });
          }
        }
      }
      else if (/^ELSE$/i.test(body)) {
        if (stack.length === 0) {
          errors.push({
            title: "Orphan [ELSE] Tag",
            tag: rawTag,
            snippet: snippet,
            fix: `Found an [ELSE] with no opening [IF] tag above it. Ensure it is placed between an <code>[IF ...]</code> and <code>[END IF]</code>.`
          });
        } else {
          const top = stack[stack.length - 1];
          if (top.hasElse) {
            errors.push({
              title: "Duplicate [ELSE] Statement",
              tag: rawTag,
              snippet: snippet,
              fix: `There is already an [ELSE] in this block. An IF block can only have one default [ELSE]. Change this to <code>[ELSE IF ...]</code> or remove the extra [ELSE].`
            });
          }
          top.hasElse = true;
        }
      }
      else if (/^END\s+IF$/i.test(body)) {
        if (stack.length === 0) {
          errors.push({
            title: "Orphan [END IF] Tag",
            tag: rawTag,
            snippet: snippet,
            fix: `In Word, search for this <code>[END IF]</code> (Ctrl+F) and delete it, or add the missing <code>[IF ...]</code> tag above this paragraph.`
          });
        } else {
          stack.pop();
        }
      }
      else if (/^ELSE\s+/i.test(body)) {
        errors.push({
          title: "Malformed [ELSE] Tag with Condition",
          tag: rawTag,
          snippet: snippet,
          fix: `<code>[ELSE]</code> cannot take a condition. If you want a specific condition, change it to <b><code>[ELSE IF ${body.replace(/^ELSE\s+/i, '')}]</code></b>. If it is the default fallback, simply write <b><code>[ELSE]</code></b>.`
        });
      }
    }

    while (stack.length > 0) {
      const unclosed = stack.pop();
      errors.push({
        title: "Missing [END IF] (Unclosed Conditional Block)",
        tag: unclosed.tag,
        snippet: unclosed.snippet,
        fix: `In Word, search for <code>${unclosed.tag}</code> and add a closing <b><code>[END IF]</code></b> tag on its own line below the conditional text.`
      });
    }

    if (errors.length === 0) {
      resultsDiv.innerHTML = `
        <div class="badge-pass">
          ✅ <b>Ready for Production:</b> "${filename}" passed all logic and tag validations with 0 errors!
        </div>
      `;

      // Safely discover and inject dynamic custom fields
      let discovered = [];
      if (typeof DocxEngine !== 'undefined' && typeof DocxEngine.extractCustomFields === 'function') {
        discovered = DocxEngine.extractCustomFields(xmlContent);
      }

      if (discovered.length > 0 && dynamicContainer && dynamicSection) {
        dynamicContainer.innerHTML = "";
        discovered.forEach(tag => {
          const fieldWrapper = document.createElement('div');
          fieldWrapper.className = 'field';

          const label = document.createElement('label');
          label.textContent = tag + ':';

          const input = document.createElement('input');
          input.type = 'text';
          input.setAttribute('data-dynamic-tag', tag);

          if (tag.toLowerCase().includes('date')) {
            input.placeholder = "e.g., May 15, 2026 or on or about May 2024...";
            input.value = "May 12, 2024";
          } else {
            input.placeholder = `Enter ${tag}...`;
            input.value = tag.toLowerCase().includes('facility') 
              ? 'Regional Detention Center' 
              : (tag.toLowerCase().includes('citizenship') 
                ? 'Guatemala' 
                : (tag.toLowerCase().includes('location') ? 'Nogales, Arizona' : 'Sample Value'));
          }

          fieldWrapper.appendChild(label);
          fieldWrapper.appendChild(input);
          dynamicContainer.appendChild(fieldWrapper);
        });
        dynamicSection.style.display = "block";
      }

      sandboxDiv.style.display = "block";
    } else {
      let errHtml = `
        <div class="error-summary-header">
          <h3>❌ Found ${errors.length} Logic Error(s) in "${filename}":</h3>
          <span class="error-summary-hint">Use <b>Ctrl + F</b> in Word to find the highlighted phrases below</span>
        </div>
      `;

      errors.forEach((e, idx) => {
        errHtml += `
          <div class="error-card">
            <div class="error-card-header">
              <h4>⚠️ Error #${idx + 1}: ${e.title}</h4>
              <code class="error-tag-badge">${escapeHtml(e.tag)}</code>
            </div>
            
            <div class="snippet-box">
              <span class="snippet-context">${escapeHtml(e.snippet.before)}</span>
              <mark class="snippet-highlight">${escapeHtml(e.snippet.tag)}</mark>
              <span class="snippet-context">${escapeHtml(e.snippet.after)}</span>
            </div>

            <div class="fix-box">
              <strong>👉 How to fix in Word:</strong> ${e.fix}
            </div>
          </div>
        `;
      });

      resultsDiv.innerHTML = errHtml;
    }

  } catch (err) {
    resultsDiv.innerHTML = `
      <div class="error-card">
        <h4>File Read Error</h4>
        <p>${escapeHtml(err.message)}</p>
      </div>
    `;
  }
}

function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// ============================================================================
// 5. LIVE TEST SANDBOX RENDERER (Delegates compilation to DocxEngine)
// ============================================================================
async function runTestRender(shouldDownload) {
  if (!currentArrayBuffer) {
    if (typeof showToast === 'function') {
      showToast("No document or snippet loaded to render.", "error");
    }
    return;
  }
  
  const testData = {};
  const container = document.getElementById('sandbox-form-container');

  function setFieldVariations(baseKey, val) {
    if (!baseKey) return;
    testData[baseKey] = val;
    const spaced = baseKey.replace(/[-_]/g, ' ').trim();
    testData[spaced] = val;
  }

  // 1. Gather Pinned Primary Data
  let fullName = "Draft";
  if (container) {
    container.querySelectorAll('input, select, textarea').forEach(el => {
      if (el.type === 'radio' && !el.checked) return;

      const val = (el.type === 'checkbox') ? el.checked : el.value;
      
      if (el.id) setFieldVariations(el.id, val);
      if (el.name) setFieldVariations(el.name, val);
      if (el.getAttribute('data-tag')) setFieldVariations(el.getAttribute('data-tag'), val);

      const label = el.closest('.field')?.querySelector('label') || el.previousElementSibling;
      if (label && label.tagName === 'LABEL') {
        const labelText = label.innerText.replace(/[*:]/g, '').trim();
        setFieldVariations(labelText, val);
      }
    });

    const first = testData['first name'] || testData['petitioner_first_name'] || testData['petitioner first name'] || '';
    const middle = testData['middle name'] || testData['petitioner_middle_name'] || testData['petitioner middle name'] || '';
    const last = testData['last name'] || testData['petitioner_last_name'] || testData['petitioner last name'] || '';
    const rawSuffix = testData['suffix'] || testData['petitioner_suffix'] || testData['petitioner suffix'] || '';

    const cleanSuffix = rawSuffix.replace(/^[,\s]+/, '').trim();
    const baseName = [first, middle, last].filter(Boolean).join(' ');
    fullName = cleanSuffix 
      ? (baseName ? `${baseName}, ${cleanSuffix}` : cleanSuffix)
      : (baseName || "Draft");

    if (first || last) {
      setFieldVariations('Petitioner Name', fullName);
      setFieldVariations('Petitioner First Name', first);
      setFieldVariations('Petitioner Middle Name', middle);
      setFieldVariations('Petitioner Last Name', last);
      setFieldVariations('Petitioner Suffix', cleanSuffix);
      setFieldVariations('Plaintiff Name', fullName);
    }

    const pronounChoice = testData['petitioner_pronouns'] || testData['petitioner pronouns'] || testData['pronouns'] || 'male';
    const pronounData = DocxEngine.getPronounData(pronounChoice);
    Object.assign(testData, pronounData);

    const facility = testData['detention_facility'] || testData['detention facility'] || testData['facility'] || '';
    if (facility) {
      setFieldVariations('Detention Facility', facility);
      setFieldVariations('Facility', facility);
    }

    const def = testData['defendant'] || testData['defendant name'] || testData['respondent name'];
    if (def) {
      setFieldVariations('Defendant Name', def);
      setFieldVariations('Respondent Name', def);
    }

    const cNum = testData['case_number'] || testData['case number'] || testData['docket number'];
    if (cNum) {
      setFieldVariations('Case Number', cNum);
      setFieldVariations('Docket Number', cNum);
    }
  }

  // 2. Gather Discovered Dynamic Fields
  document.querySelectorAll('#sandbox-dynamic-container input[data-dynamic-tag]').forEach(input => {
    const tag = input.getAttribute('data-dynamic-tag');
    const val = input.value.trim();
    setFieldVariations(tag, val);
  });

  // 3. Delegate directly to DocxEngine (replaces the ~100 duplicate lines)
  await DocxEngine.generate({
    buffer: currentArrayBuffer,
    data: testData,
    outputFilename: `Habeas Petition ${fullName}.docx`,
    selectElementForAliases: true,
    download: shouldDownload,
    previewContainerId: 'docx-preview-container'
  });
}

function isTagSyntaxValid(tagText) {
  if (!tagText) return true;
  
  const upperText = tagText.toUpperCase().trim();
  const logicalKeywords = ["IF ", "ELSE", "FOR ", "EACH", "ENDIF", "="];
  const isLogicalStatement = logicalKeywords.some(keyword => upperText.indexOf(keyword) !== -1);
  
  if (!isLogicalStatement) return true;
  
  const normalizedText = tagText.replace(/[“”]/g, '"');
  const doubleQuoteCount = (normalizedText.match(/"/g) || []).length;
  
  return doubleQuoteCount % 2 === 0;
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
    .split('\n')
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