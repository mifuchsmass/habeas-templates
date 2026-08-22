/**
 * Volunteer Drafting Studio & Template Logic Validator (Smart Field Sync Edition)
 */

// ============================================================================
// 1. CONFIGURATION & STATE
// ============================================================================
const ACCESS_PASSWORD = "Habeas123!";
let currentArrayBuffer = null;

/**
 * Derives comprehensive grammatical pronoun and party tags
 */
function getPronounData(choice) {
  const map = {
    male: {
      he: 'he', He: 'He', HE: 'HE',
      him: 'him', Him: 'Him', HIM: 'HIM',
      his: 'his', His: 'His', HIS: 'HIS',
      hers: 'his', Hers: 'His', HERS: 'HIS',
      himself: 'himself', Himself: 'Himself', HIMSELF: 'HIMSELF',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'is', was_were: 'was', has_have: 'has',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    },
    female: {
      he: 'she', He: 'She', HE: 'SHE',
      him: 'her', Him: 'Her', HIM: 'HER',
      his: 'her', His: 'Her', HIS: 'HER',
      hers: 'hers', Hers: 'Hers', HERS: 'HERS',
      himself: 'herself', Himself: 'Herself', HIMSELF: 'HERSELF',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'is', was_were: 'was', has_have: 'has',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    },
    nonbinary: {
      he: 'they', He: 'They', HE: 'THEY',
      him: 'them', Him: 'Them', HIM: 'THEM',
      his: 'their', His: 'Their', HIS: 'THEIR',
      hers: 'theirs', Hers: 'Theirs', HERS: 'THEIRS',
      himself: 'themselves', Himself: 'Themselves', HIMSELF: 'THEMSELVES',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'are', was_were: 'were', has_have: 'have',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    },
    petitioner: {
      he: 'Petitioner', He: 'Petitioner', HE: 'PETITIONER',
      him: 'Petitioner', Him: 'Petitioner', HIM: 'PETITIONER',
      his: "Petitioner's", His: "Petitioner's", HIS: "PETITIONER'S",
      hers: "Petitioner's", Hers: "Petitioner's", HERS: "PETITIONER'S",
      himself: 'Petitioner', Himself: 'Petitioner', HIMSELF: 'PETITIONER',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'is', was_were: 'was', has_have: 'has',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    }
  };

  const p = map[choice] || map.male;

  return {
    "he": p.he, "He": p.He, "HE": p.HE,
    "him": p.him, "Him": p.Him, "HIM": p.HIM,
    "his": p.his, "His": p.His, "HIS": p.HIS,
    "hers": p.hers, "Hers": p.Hers, "HERS": p.HERS,
    "himself": p.himself, "Himself": p.Himself, "HIMSELF": p.HIMSELF,
    "he/she": p.he, "He/She": p.He, "HE/SHE": p.HE,
    "his/her": p.his, "His/Her": p.His, "HIS/HER": p.HIS,
    "him/her": p.him, "Him/Her": p.Him, "HIM/HER": p.HIM,
    "himself/herself": p.himself, "Himself/Herself": p.Himself, "HIMSELF/HERSELF": p.HIMSELF,
    "his/hers": p.hers,

    "Petitioner": p.petitioner,
    "petitioner": p.petitioner_lower,
    "PETITIONER": p.PETITIONER,
    "Petitioner's": p.petitioner_possessive,
    "petitioner's": p.petitioner_possessive.toLowerCase(),
    "PETITIONER'S": p.PETITIONER_POSSESSIVE,
    "Petitioner/Petitioners": p.petitioner,
    "petitioner/petitioners": p.petitioner_lower,
    "PETITIONER/PETITIONERS": p.PETITIONER,
    "Petitioner's/Petitioners'": p.petitioner_possessive,
    "petitioner's/petitioners'": p.petitioner_possessive.toLowerCase(),

    "Pronoun Subject": p.He, "pronoun subject": p.he,
    "Pronoun Object": p.Him, "pronoun object": p.him,
    "Pronoun Possessive": p.His, "pronoun possessive": p.his,
    "Pronoun Reflexive": p.Himself, "pronoun reflexive": p.himself,

    "individual/individuals": p.individual,
    "an individual/individuals": p.individual,

    "is/are": p.is_are, "was/were": p.was_were, "has/have": p.has_have,
    "brings/bring": p.brings_bring, "seeks/seek": p.seeks_seek, "contends/contend": p.contends_contend,
    "s": p.s_suffix,

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

    clone.querySelectorAll('button, #docx-preview-container, .preview-section, [type="submit"], h2, #dynamic-fields-section').forEach(el => el.remove());

    container.innerHTML = "";
    container.appendChild(clone);
  } catch (err) {
    console.warn("Auto-sync fallback: Using default primary fields", err);
    container.innerHTML = `
      <div class="sandbox-grid">
        <div class="field"><label>Petitioner Name</label><input type="text" id="Petitioner Name" value="Jane Marie Doe"></div>
        <div class="field"><label>Case Number</label><input type="text" id="Case Number" value="2026-CV-04321"></div>
        <div class="field"><label>Jurisdiction</label><select id="Jurisdiction"><option value="Massachusetts">Massachusetts</option><option value="Rhode Island">Rhode Island</option></select></div>
        <div class="field"><label>Type of Habeas Petition</label><select id="Type of Habeas Petition"><option value="GO Class">GO Class</option><option value="Non-GO">Non-GO</option></select></div>
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

      // Discover and inject dynamic custom fields
      const discovered = DocxEngine.extractCustomFields(xmlContent);
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
          input.value = tag.toLowerCase().includes('date') ? 'May 12, 2024' : (tag.toLowerCase().includes('facility') ? 'Regional Detention Center' : (tag.toLowerCase().includes('citizenship') ? 'Guatemala' : (tag.toLowerCase().includes('location') ? 'Nogales, Arizona' : 'Sample Value')));

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
// 5. LIVE TEST SANDBOX RENDERER
// ============================================================================
async function runTestRender(shouldDownload) {
  if (!currentArrayBuffer) return;
  
  const testData = {};
  const container = document.getElementById('sandbox-form-container');

  // Helper to record base key and space variations without polluting uppercase keys
  function setFieldVariations(baseKey, val) {
    if (!baseKey) return;
    testData[baseKey] = val;
    const spaced = baseKey.replace(/[-_]/g, ' ').trim();
    testData[spaced] = val;
  }

  // 1. Gather Pinned Primary Data
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
    const suffix = testData['suffix'] || testData['petitioner_suffix'] || testData['petitioner suffix'] || '';

    if (first || last) {
      const fullName = [first, middle, last, suffix].filter(Boolean).join(' ');
      setFieldVariations('Petitioner Name', fullName);
      setFieldVariations('Petitioner First Name', first);
      setFieldVariations('Petitioner Middle Name', middle);
      setFieldVariations('Petitioner Last Name', last);
      setFieldVariations('Petitioner Suffix', suffix);
      setFieldVariations('Plaintiff Name', fullName);
    }

    const pronounChoice = testData['petitioner_pronouns'] || testData['petitioner pronouns'] || testData['pronouns'] || 'male';
    const pronounData = getPronounData(pronounChoice);
    Object.assign(testData, pronounData);

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

            // 1. EXACT CASE MATCH FIRST
            if (scope[cleanTag] !== undefined && scope[cleanTag] !== null && scope[cleanTag] !== "") {
              const val = scope[cleanTag];
              // Smart Caps: If tag was typed in [ALL CAPS], output in ALL CAPS
              if (typeof val === "string" && cleanTag === cleanTag.toUpperCase() && /[A-Z]/.test(cleanTag)) {
                return val.toUpperCase();
              }
              return val;
            }

            // 2. Direct Key Match (case-insensitive fallback)
            const target = cleanTag.toLowerCase().replace(/[-_\s]/g, '');
            for (const k of Object.keys(scope)) {
              if (k.toLowerCase().replace(/[-_\s]/g, '') === target) {
                const val = scope[k];
                if (val === "" || val === undefined || val === null) {
                  return "*** MISSING DATA ***";
                }
                // Smart Caps: If tag was typed in [ALL CAPS], output in ALL CAPS
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