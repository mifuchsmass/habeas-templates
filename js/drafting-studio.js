/**
 * Volunteer Drafting Studio & Template Logic Validator (Smart Field Sync Edition)
 */

// ============================================================================
// 1. CONFIGURATION & STATE
// ============================================================================
const ACCESS_PASSWORD = "Habeas123!";
let currentArrayBuffer = null;

/**
 * Derives comprehensive grammatical pronoun tags with distinct casing
 */
function getPronounData(choice) {
  const map = {
    male: {
      he: 'he', He: 'He', HE: 'HE',
      him: 'him', Him: 'Him', HIM: 'HIM',
      his: 'his', His: 'His', HIS: 'HIS',
      hers: 'his', Hers: 'His', HERS: 'HIS',
      himself: 'himself', Himself: 'Himself', HIMSELF: 'HIMSELF',
      is_are: 'is', was_were: 'was', has_have: 'has'
    },
    female: {
      he: 'she', He: 'She', HE: 'SHE',
      him: 'her', Him: 'Her', HIM: 'HER',
      his: 'her', His: 'Her', HIS: 'HER',
      hers: 'hers', Hers: 'Hers', HERS: 'HERS',
      himself: 'herself', Himself: 'Herself', HIMSELF: 'HERSELF',
      is_are: 'is', was_were: 'was', has_have: 'has'
    },
    nonbinary: {
      he: 'they', He: 'They', HE: 'THEY',
      him: 'them', Him: 'Them', HIM: 'THEM',
      his: 'their', His: 'Their', HIS: 'THEIR',
      hers: 'theirs', Hers: 'Theirs', HERS: 'THEIRS',
      himself: 'themselves', Himself: 'Themselves', HIMSELF: 'THEMSELVES',
      is_are: 'are', was_were: 'were', has_have: 'have'
    },
    petitioner: {
      he: 'Petitioner', He: 'Petitioner', HE: 'PETITIONER',
      him: 'Petitioner', Him: 'Petitioner', HIM: 'PETITIONER',
      his: "Petitioner's", His: "Petitioner's", HIS: "PETITIONER'S",
      hers: "Petitioner's", Hers: "Petitioner's", HERS: "PETITIONER'S",
      himself: 'Petitioner', Himself: 'Petitioner', HIMSELF: 'PETITIONER',
      is_are: 'is', was_were: 'was', has_have: 'has'
    }
  };

  const p = map[choice] || map.male;

  return {
    // Lowercase
    "he": p.he,
    "him": p.him,
    "his": p.his,
    "hers": p.hers,
    "himself": p.himself,
    "he/she": p.he,
    "his/her": p.his,
    "him/her": p.him,
    "himself/herself": p.himself,
    "his/hers": p.hers,
    "pronoun subject": p.he,
    "pronoun object": p.him,
    "pronoun possessive": p.his,
    "pronoun reflexive": p.himself,

    // Capitalized / Title Case
    "He": p.He,
    "Him": p.Him,
    "His": p.His,
    "Hers": p.Hers,
    "Himself": p.Himself,
    "He/She": p.He,
    "His/Her": p.His,
    "Him/Her": p.Him,
    "Himself/Herself": p.Himself,
    "Pronoun Subject": p.He,
    "Pronoun Object": p.Him,
    "Pronoun Possessive": p.His,
    "Pronoun Reflexive": p.Himself,

    // ALL CAPS
    "HE": p.HE,
    "HIM": p.HIM,
    "HIS": p.HIS,
    "HERS": p.HERS,
    "HIMSELF": p.HIMSELF,
    "HE/SHE": p.HE,
    "HIS/HER": p.HIS,
    "HIM/HER": p.HIM,
    "HIMSELF/HERSELF": p.HIMSELF,

    // Verbs
    "is/are": p.is_are,
    "was/were": p.was_were,
    "has/have": p.has_have,

    // Logical conditions
    "Pronouns": choice,
    "Gender": choice
  };
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

  loadFormFromIndex();

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

  if (renderBtn) {
    renderBtn.addEventListener('click', () => runTestRender(false));
  }
  if (downloadBtn) {
    downloadBtn.addEventListener('click', () => runTestRender(true));
  }

  // Dropzone Handlers
  if (dropzone && fileInput) {
    dropzone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });

    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length) {
        handleFile(e.dataTransfer.files[0]);
      }
    });
  }
});

function handleFile(file) {
  if (!file || !file.name.toLowerCase().endsWith('.docx')) {
    alert('Please select a Microsoft Word .docx document.');
    return;
  }
  const reader = new FileReader();
  reader.onload = function(e) {
    currentArrayBuffer = e.target.result;
    inspectTemplate(currentArrayBuffer, file.name);
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================================
// 3. AUTO-SYNC: CLONE FORM FROM index.html & POPULATE TEST DATA
// ============================================================================
async function loadFormFromIndex() {
  const container = document.getElementById('sandbox-form-container');
  if (!container) return;

  try {
    const res = await fetch('index.html');
    if (!res.ok) throw new Error('Could not fetch index.html');
    
    const htmlText = await res.text();
    const doc = new DOMParser().parseFromString(htmlText, 'text/html');

    const formElement = doc.querySelector('form') || doc.querySelector('.form-container') || doc.querySelector('main') || doc.body;
    const clone = formElement.cloneNode(true);

    clone.querySelectorAll('button, #docx-preview-container, .preview-section, [type="submit"], h2').forEach(el => el.remove());

    container.innerHTML = "";
    container.appendChild(clone);

    populateSampleTestValues(container);

  } catch (err) {
    console.warn("Auto-sync fallback: Using default fields", err);
    container.innerHTML = `
      <div class="sandbox-grid">
        <div class="field"><label>Petitioner Name</label><input type="text" id="Petitioner Name" value="Jane Marie Doe"></div>
        <div class="field"><label>Case Number</label><input type="text" id="Case Number" value="1:26-cv-10042"></div>
        <div class="field"><label>Jurisdiction</label><select id="Jurisdiction"><option value="Massachusetts">Massachusetts</option><option value="Rhode Island">Rhode Island</option><option value="New Hampshire">New Hampshire</option></select></div>
        <div class="field"><label>Type of Habeas Petition</label><select id="Type of Habeas Petition"><option value="GO Class">GO Class</option><option value="Individual">Individual</option></select></div>
      </div>
    `;
  }
}

function populateSampleTestValues(container) {
  container.querySelectorAll('input[type="text"], input:not([type]), textarea').forEach(input => {
    if (!input.value) {
      const id = (input.id || '').toLowerCase();
      const name = (input.name || '').toLowerCase();
      const label = (input.closest('.field')?.querySelector('label')?.innerText || '').toLowerCase();
      const key = id + ' ' + name + ' ' + label;

      if (key.includes('first')) input.value = 'John';
      else if (key.includes('middle')) input.value = 'David';
      else if (key.includes('last')) input.value = 'Doe';
      else if (key.includes('suffix')) input.value = '';
      else if (key.includes('petitioner')) input.value = 'John David Doe';
      else if (key.includes('defendant') || key.includes('respondent')) input.value = 'Department of Homeland Security';
      else if (key.includes('case') || key.includes('docket')) input.value = '2026-CV-04321';
      else input.value = 'Sample Value';
    }
  });
}

// ============================================================================
// 4. TEMPLATE LOGIC SCANNER
// ============================================================================
function inspectTemplate(buffer, filename) {
  const resultsDiv = document.getElementById('results');
  const sandboxDiv = document.getElementById('sandbox');
  resultsDiv.innerHTML = "";
  sandboxDiv.style.display = "none";

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
// 5. LIVE TEST SANDBOX RENDERER
// ============================================================================
async function runTestRender(shouldDownload) {
  if (!currentArrayBuffer) return;
  
  const testData = {};
  const container = document.getElementById('sandbox-form-container');

  function setFieldVariations(baseKey, val) {
    if (!baseKey) return;
    testData[baseKey] = val;
    testData[baseKey.toLowerCase()] = val;
    testData[baseKey.toUpperCase()] = val;
    const spaced = baseKey.replace(/[-_]/g, ' ').trim();
    testData[spaced] = val;
    testData[spaced.toLowerCase()] = val;
    testData[spaced.toUpperCase()] = val;
  }

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

    const first = testData['first name'] || testData['first_name'] || testData['petitioner first name'] || '';
    const middle = testData['middle name'] || testData['middle_name'] || testData['petitioner middle name'] || '';
    const last = testData['last name'] || testData['last_name'] || testData['petitioner last name'] || '';
    const suffix = testData['suffix'] || testData['petitioner suffix'] || '';

    if (first || last) {
      const fullName = [first, middle, last, suffix].filter(Boolean).join(' ');
      setFieldVariations('Petitioner Name', fullName);
      setFieldVariations('Petitioner First Name', first);
      setFieldVariations('Petitioner Middle Name', middle);
      setFieldVariations('Petitioner Last Name', last);
      setFieldVariations('Petitioner Suffix', suffix);
      setFieldVariations('Plaintiff Name', fullName);
    }

    // Direct pronoun assignment (preserves exact casing)
    const pronounChoice = testData['petitioner pronouns'] || testData['petitioner_pronouns'] || testData['pronouns'] || 'male';
    const pronounData = getPronounData(pronounChoice);
    Object.assign(testData, pronounData);

    const def = testData['defendant name'] || testData['defendant_name'] || testData['respondent name'] || testData['defendant'];
    if (def) {
      setFieldVariations('Defendant Name', def);
      setFieldVariations('Respondent Name', def);
    }

    const cNum = testData['case number'] || testData['case_number'] || testData['docket number'] || testData['case'];
    if (cNum) {
      setFieldVariations('Case Number', cNum);
      setFieldVariations('Docket Number', cNum);
    }
  }

  const optionAliases = {};
  document.querySelectorAll('#sandbox-form-container select option').forEach(opt => {
    const val = opt.value;
    const fullText = opt.text.trim();
    const textWithoutParens = fullText.replace(/\s*\([^)]*\)\s*$/, '').trim();
    if (val) {
      optionAliases[fullText.toLowerCase()] = val;
      optionAliases[textWithoutParens.toLowerCase()] = val;
    }
  });

  try {
    const zip = new PizZip(currentArrayBuffer);
    
    Object.keys(zip.files).forEach(filename => {
      if (filename.startsWith("word/") && filename.endsWith(".xml")) {
        let xml = zip.files[filename].asText().replace(/\[[^\]]*?\]/g, match => match.replace(/<[^>]+>/g, ''));
        const tagRegex = /\[\s*(IF\s+[^\]]+|ELSE\s+IF\s+[^\]]+|ELSEIF\s+[^\]]+|ELSE|END\s+IF)\s*\]/gi;
        const stack = [];
        
        xml = xml.replace(tagRegex, (m, tagBody) => {
          const trimmed = tagBody.trim();
          const sanitize = (c) => c.trim()
            .replace(/[\u201C\u201D]/g, '"')
            .replace(/[\u2018\u2019]/g, "'")
            .replace(/\band\b/gi, '&&')
            .replace(/\bor\b/gi, '||')
            .replace(/\s+=\s+/g, ' == ');
          
          if (/^IF\s+/i.test(trimmed)) {
            const cond = sanitize(trimmed.replace(/^IF\s+/i, ''));
            stack.push({ conditions: [cond], hasElse: false });
            return `[#${cond}]`;
          }
          if (/^ELSE\s*IF\s+/i.test(trimmed) || /^ELSEIF\s+/i.test(trimmed)) {
            if (!stack.length) return m;
            const f = stack[stack.length - 1];
            if (f.hasElse) return m;
            const cond = sanitize(trimmed.replace(/^(ELSE\s*IF|ELSEIF)\s+/i, ''));
            const prev = f.conditions[f.conditions.length - 1];
            f.conditions.push(cond);
            return `[/${prev}][^${prev}][#${cond}]`;
          }
          if (/^ELSE$/i.test(trimmed)) {
            if (!stack.length) return m;
            const f = stack[stack.length - 1];
            if (f.hasElse) return m;
            f.hasElse = true;
            const prev = f.conditions[f.conditions.length - 1];
            return `[/${prev}][^${prev}]`;
          }
          if (/^END\s+IF$/i.test(trimmed)) {
            if (!stack.length) return m;
            const f = stack.pop();
            let cl = "";
            for (let i = f.conditions.length - 1; i >= 0; i--) cl += `[/${f.conditions[i]}]`;
            return cl;
          }
          return m;
        });
        zip.file(filename, xml);
      }
    });

    const doc = new window.docxtemplater(zip, {
      delimiters: { start: '[', end: ']' },
      paragraphLoop: true,
      linebreaks: true,
      parser: (tag) => {
        const cleanTag = tag.replace(/^[#\/\^]/, '').trim();
        return {
          get(scope) {
            if (!cleanTag) return "";

            // 1. EXACT CASE MATCH FIRST (Preserves distinct [he] vs [He] vs [HE])
            if (scope[cleanTag] !== undefined && scope[cleanTag] !== null && scope[cleanTag] !== "") {
              return scope[cleanTag];
            }

            // 2. Direct Key Match (case-insensitive fallback)
            const target = cleanTag.toLowerCase().replace(/[-_\s]/g, '');
            for (const k of Object.keys(scope)) {
              if (k.toLowerCase().replace(/[-_\s]/g, '') === target) {
                const val = scope[k];
                if (val === "" || val === undefined || val === null) {
                  return "*** MISSING DATA ***";
                }
                if (typeof val === "string" && cleanTag === cleanTag.toUpperCase() && /[A-Z]/.test(cleanTag)) {
                  return val.toUpperCase();
                }
                return val;
              }
            }

            const isCondition = /[=!<>]|\b(and|or)\b|&&|\|\|/i.test(cleanTag);
            if (!isCondition) {
              return "*** MISSING DATA ***";
            }

            try {
              let expr = cleanTag;
              expr = expr.replace(/(["'])(.*?)\1/g, (m, q, text) => {
                const lower = text.trim().toLowerCase();
                return optionAliases[lower] ? `"${optionAliases[lower]}"` : m;
              });

              Object.keys(scope).sort((a, b) => b.length - a.length).forEach(v => {
                const escaped = v.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                expr = expr.replace(new RegExp('\\b' + escaped + '\\b', 'gi'), 'scope["' + v + '"]');
              });

              return new Function("scope", "return Boolean(" + expr + ");")(scope);
            } catch (e) {
              return false;
            }
          }
        };
      }
    });

    doc.render(testData);
    const out = doc.getZip().generate({ type: "blob" });

    if (shouldDownload === true) {
      const link = document.createElement("a");
      link.href = URL.createObjectURL(out);
      link.download = "Drafting_Studio_Output.docx";
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      return;
    }

    const previewContainer = document.getElementById('docx-preview-container');
    previewContainer.innerHTML = "";
    await docx.renderAsync(out, previewContainer);
  } catch (err) {
    document.getElementById('docx-preview-container').innerHTML = `
      <p class="preview-error">Render Error: ${escapeHtml(err.message)}</p>
    `;
  }
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