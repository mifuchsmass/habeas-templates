/**
 * Universal DOCX Template Engine
 */
const DocxEngine = {
  templateCache: {},

  async loadTemplate(templatePath) {
    if (this.templateCache[templatePath]) {
      return this.templateCache[templatePath].slice(0);
    }
    const response = await fetch(templatePath);
    if (!response.ok) throw new Error(`Failed to load template: ${templatePath}`);
    const buffer = await response.arrayBuffer();
    this.templateCache[templatePath] = buffer;
    return buffer.slice(0);
  },

  sanitizeCondition(rawCond) {
    return rawCond.trim()
      .replace(/[\u201C\u201D]/g, '"')
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/\band\b/gi, '&&')
      .replace(/\bor\b/gi, '||')
      .replace(/\s+=\s+/g, ' == ');
  },

  parseIfTags(xmlContent) {
    xmlContent = xmlContent.replace(/\[[^\]]*?\]/g, match => match.replace(/<[^>]+>/g, ''));
    const tagRegex = /\[\s*(IF\s+[^\]]+|ELSE\s+IF\s+[^\]]+|ELSEIF\s+[^\]]+|ELSE|END\s+IF)\s*\]/gi;
    const stack = [];

    return xmlContent.replace(tagRegex, (match, tagBody) => {
      const trimmed = tagBody.trim();

      if (/^IF\s+/i.test(trimmed)) {
        const condition = this.sanitizeCondition(trimmed.replace(/^IF\s+/i, ''));
        stack.push({ conditions: [condition], hasElse: false });
        return `[#${condition}]`;
      }

      if (/^ELSE\s*IF\s+/i.test(trimmed) || /^ELSEIF\s+/i.test(trimmed)) {
        if (stack.length === 0) return match;
        const currentFrame = stack[stack.length - 1];
        if (currentFrame.hasElse) return match;

        const rawCond = trimmed.replace(/^(ELSE\s*IF|ELSEIF)\s+/i, '');
        const newCondition = this.sanitizeCondition(rawCond);
        const prevCondition = currentFrame.conditions[currentFrame.conditions.length - 1];

        currentFrame.conditions.push(newCondition);
        return `[/${prevCondition}][^${prevCondition}][#${newCondition}]`;
      }

      if (/^ELSE$/i.test(trimmed)) {
        if (stack.length === 0) return match;
        const currentFrame = stack[stack.length - 1];
        if (currentFrame.hasElse) return match;

        currentFrame.hasElse = true;
        const prevCondition = currentFrame.conditions[currentFrame.conditions.length - 1];
        return `[/${prevCondition}][^${prevCondition}]`;
      }

      if (/^END\s+IF$/i.test(trimmed)) {
        if (stack.length === 0) return match;
        const frame = stack.pop();

        let closingTags = "";
        for (let i = frame.conditions.length - 1; i >= 0; i--) {
          closingTags += `[/${frame.conditions[i]}]`;
        }
        return closingTags;
      }

      return match;
    });
  },

  extractCustomFields(xmlOrBuffer) {
    let xml = "";
    if (typeof xmlOrBuffer === "string") {
      xml = xmlOrBuffer;
    } else {
      const zip = new PizZip(xmlOrBuffer);
      Object.keys(zip.files).forEach(name => {
        if (name.startsWith("word/") && name.endsWith(".xml")) {
          xml += zip.files[name].asText() + "\n";
        }
      });
    }

    const cleanDocText = xml
      .replace(/\[[^\]]*?\]/g, match => match.replace(/<[^>]+>/g, ''))
      .replace(/<[^>]+>/g, '')
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&apos;/g, "'");

    const tagRegex = /\[\s*([^\]]+?)\s*\]/g;
    const discovered = [];
    const seen = new Set();

    const reservedTags = new Set([
      "jurisdiction", "petition type", "type of habeas petition", "type of habeas petition full",
      "petitioner name", "petitioner first name", "petitioner middle name", "petitioner last name", "petitioner suffix",
      "plaintiff name", "defendant name", "respondent name", "case number", "docket number",
      "he", "she", "they", "him", "her", "them", "his", "hers", "their", "theirs", "himself", "herself", "themselves",
      "he/she", "his/her", "him/her", "himself/herself", "his/hers", 
      "pronoun subject", "pronoun object", "pronoun possessive", "pronoun reflexive",
      "pronouns", "gender", "is plural",
      "petitioner", "petitioners", "petitioner's", "petitioners'", "petitioner/petitioners", "petitioner's/petitioners'",
      "individual/individuals", "an individual/individuals",
      "is/are", "was/were", "has/have", "brings/bring", "seeks/seek", "contends/contend", "s",
      "else", "endif", "end if"
    ]);

    let match;
    while ((match = tagRegex.exec(cleanDocText)) !== null) {
      const rawTag = match[1].trim();

      if (/^(if\s+|else\s*if\s+|elseif\s+|else$|end\s+if$|for\s+|each\s+)/i.test(rawTag)) continue;
      if (/[=!<>]|\b(and|or)\b|&&|\|\|/i.test(rawTag)) continue;
      if (/^(note|instruction|instructions|todo|guidance|comment)\b/i.test(rawTag)) continue;

      const lookupKey = rawTag.toLowerCase().replace(/[-_\s]/g, '');
      let isReserved = false;
      for (const r of reservedTags) {
        if (r.replace(/[-_\s]/g, '') === lookupKey) {
          isReserved = true;
          break;
        }
      }
      if (isReserved) continue;

      if (!seen.has(lookupKey)) {
        seen.add(lookupKey);
        discovered.push(rawTag);
      }
    }

    return discovered;
  },

  async generate({ templatePath, data, outputFilename, selectElementForAliases, download = false }) {
    const container = document.getElementById('docx-preview-container');

    try {
      const buffer = await this.loadTemplate(templatePath);
      const zip = new PizZip(buffer);

      Object.keys(zip.files).forEach(filename => {
        if (filename.startsWith("word/") && filename.endsWith(".xml")) {
          zip.file(filename, this.parseIfTags(zip.files[filename].asText()));
        }
      });

      const optionAliases = {};
      if (selectElementForAliases) {
        document.querySelectorAll('select option').forEach(opt => {
          const val = opt.value;
          const fullText = opt.text.trim();
          const textWithoutParens = fullText.replace(/\s*\([^)]*\)\s*$/, '').trim();
          if (val) {
            optionAliases[fullText.toLowerCase()] = val;
            optionAliases[textWithoutParens.toLowerCase()] = val;
          }
        });
      }

      const customParser = (tag) => {
        const cleanTag = tag.replace(/^[#\/\^]/, '').trim();
        return {
          get(scope) {
            if (!cleanTag) return "";

            // 1. EXACT CASE MATCH FIRST
            if (scope[cleanTag] !== undefined && scope[cleanTag] !== null && scope[cleanTag] !== "") {
              const val = scope[cleanTag];
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

              Object.keys(scope).sort((a, b) => b.length - a.length).forEach(varName => {
                const escaped = varName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                expr = expr.replace(new RegExp('\\b' + escaped + '\\b', 'gi'), 'scope["' + varName + '"]');
              });

              return new Function("scope", `return Boolean(${expr});`)(scope);
            } catch (e) {
              return false;
            }
          }
        };
      };

      const doc = new window.docxtemplater(zip, {
        delimiters: { start: '[', end: ']' },
        paragraphLoop: true,
        linebreaks: true,
        parser: customParser
      });

      doc.render(data);

      const out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });

      if (download === true) {
        if (typeof saveAs === "function") {
          saveAs(out, outputFilename || "Document_Output.docx");
        } else {
          const link = document.createElement("a");
          link.href = URL.createObjectURL(out);
          link.download = outputFilename || "Document_Output.docx";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
        }
        return;
      }

      if (container) {
        if (window.docx && typeof window.docx.renderAsync === "function") {
          container.innerHTML = "";
          await window.docx.renderAsync(out, container);
        } else {
          container.innerHTML = `<p class="preview-error">docx-preview rendering library not loaded.</p>`;
        }
      }
    } catch (error) {
      if (container) {
        container.innerHTML = `<p class="preview-error">Render Error: ${error.message}</p>`;
      }
      if (download) {
        alert("Error generating document: " + error.message);
      }
      console.error(error);
    }
  }
};

// ============================================================================
// GLOBAL TOAST NOTIFICATIONS (Explicitly attached to window)
// ============================================================================
window.showToast = function(message, type = 'success', duration = 3000) {
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }

  const toast = document.createElement('div');
  toast.className = `toast-notification toast-${type}`;
  
  const icon = type === 'success' ? '✅' : (type === 'error' ? '⚠️' : 'ℹ️');
  toast.innerHTML = `<span>${icon}</span> <span>${message}</span>`;
  
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(-10px)';
    setTimeout(() => toast.remove(), 300);
  }, duration);
};

// ============================================================================
// GLOBAL FORM STORAGE (Consistent Dynamic File Naming)
// ============================================================================
window.FormStorage = {
  save(prefix = 'Habeas_Data') {
    // 1. Extract Name Parts (Supports both index.html and Studio Sandbox)
    let first = document.querySelector('#petitioner_first_name')?.value.trim() || '';
    let middle = document.querySelector('#petitioner_middle_name')?.value.trim() || '';
    let last = document.querySelector('#petitioner_last_name')?.value.trim() || '';
    let suffix = document.querySelector('#petitioner_suffix')?.value.trim() || '';

    // Fallback if full name is in a single field
    if (!first && !last) {
      const fullInput = document.querySelector('#Petitioner\\ Name') || 
                        document.querySelector('[name="Petitioner Name"]');
      if (fullInput) first = fullInput.value.trim();
    }

    const cleanName = [first, middle, last, suffix]
      .filter(Boolean)
      .join('_')
      .replace(/[^a-zA-Z0-9_-]/g, '_');

    const filename = `${prefix}_${cleanName || 'Draft'}.json`;

    // 2. Gather All Form Data
    const data = {};
    document.querySelectorAll('input, select, textarea').forEach(el => {
      const key = el.getAttribute('data-dynamic-tag') || el.name || el.id;
      if (!key) return;

      if (el.type === 'radio') {
        if (el.checked) data[el.name] = el.value;
      } else if (el.type === 'checkbox') {
        data[key] = el.checked;
      } else {
        data[key] = el.value;
      }
    });

    // 3. Trigger Download
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    window.showToast(`Saved as ${filename}!`, "success");
  },

  openFileDialog() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json,.txt';
    input.style.display = 'none';
    document.body.appendChild(input);

    input.onchange = (e) => {
      if (e.target.files && e.target.files.length) {
        this.load(e.target.files[0]);
      }
      document.body.removeChild(input);
    };
    input.click();
  },

  load(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        Object.keys(data).forEach(key => {
          // 1. Radio Buttons (Pronouns)
          const radio = document.querySelector(`input[type="radio"][name="${key}"][value="${data[key]}"]`);
          if (radio) {
            radio.checked = true;
            return;
          }

          // 2. Text, Select, Checkboxes & Dynamic Fields
          const el = document.querySelector(`[data-dynamic-tag="${key}"]`) || 
                     document.getElementById(key) || 
                     document.querySelector(`[name="${key}"]`);
          if (el) {
            if (el.type === 'checkbox') el.checked = Boolean(data[key]);
            else el.value = data[key];
          }
        });

        window.showToast("Form data loaded successfully!", "success");
      } catch (err) {
        window.showToast("Error loading file: Invalid format.", "error");
      }
    };
    reader.readAsText(file);
  }
};