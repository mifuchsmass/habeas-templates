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

            // 3. Condition vs Missing Placeholder Check
            const isCondition = /[=!<>]|\b(and|or)\b|&&|\|\|/i.test(cleanTag);
            if (!isCondition) {
              return "*** MISSING DATA ***";
            }

            // 4. Condition Evaluation
            try {
              let expr = cleanTag;
              expr = expr.replace(/(["'])(.*?)\1/g, (m, q, text) => {
                const lower = text.trim().toLowerCase();
                return optionAliases[lower] ? `"${optionAliases[lower]}"` : m;
              });

              Object.keys(scope).sort((a, b) => b.length - a.length).forEach(varName => {
                const escaped = varName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                const regex = new RegExp(`\\b${escaped}\\b`, 'gi');
                expr = expr.replace(regex, `scope["${varName}"]`);
              });

              const fn = new Function("scope", `return Boolean(${expr});`);
              return fn(scope);
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

      // File Download (Skips preview)
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

      // Render In-Browser Preview
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