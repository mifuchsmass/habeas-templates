/**
 * Universal DOCX Template Engine
 */
const DocxEngine = {
  // Helper to sanitize condition strings
  sanitizeCondition(rawCond) {
    return rawCond.trim()
      .replace(/[\u201C\u201D]/g, '"')
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/\band\b/gi, '&&')
      .replace(/\bor\b/gi, '||')
      .replace(/\s+=\s+/g, ' == ');
  },

  // 1. Repair split XML and translate [IF ...] / [ELSE IF ...] / [ELSE] / [END IF]
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

  // 2. Main Generation Function
  async generate({ templatePath, data, outputFilename, selectElementForAliases, showPreview }) {
    try {
      const response = await fetch(templatePath);
      if (!response.ok) throw new Error(`Failed to load template: ${templatePath}`);
      const buffer = await response.arrayBuffer();
      const zip = new PizZip(buffer);

      // Pre-process XML files
      Object.keys(zip.files).forEach(filename => {
        if (filename.startsWith("word/") && filename.endsWith(".xml")) {
          zip.file(filename, this.parseIfTags(zip.files[filename].asText()));
        }
      });

      // Build Option Aliases if select elements are provided
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

      // Parser configuration
      const customParser = (tag) => {
        const cleanTag = tag.replace(/^[#\/\^]/, '').trim();
        return {
          get(scope) {
            if (!cleanTag) return "";

            // Direct Key match (case-insensitive)
            const target = cleanTag.toLowerCase();
            for (const k of Object.keys(scope)) {
              if (k.toLowerCase() === target) {
                const val = scope[k];
                if (val === "" || val === undefined || val === null) {
                  return "*** MISSING DATA ***";
                }
                
                // If tag was typed in [ALL CAPS] in the template, output in ALL CAPS:
                if (typeof val === "string" && cleanTag === cleanTag.toUpperCase() && /[A-Z]/.test(cleanTag)) {
                  return val.toUpperCase();
                }

                return val;
              }
            }
            // Condition evaluation
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

      // Render In-Browser Preview only if checkbox is selected
      const container = document.getElementById('docx-preview-container');
      if (container) {
        if (showPreview && window.docx && typeof window.docx.renderAsync === "function") {
          container.innerHTML = "";
          await window.docx.renderAsync(out, container);
        } else {
          container.innerHTML = `<p style="color: #888; font-style: italic; padding: 10px;">Preview skipped.</p>`;
        }
      }

      saveAs(out, outputFilename || "Document_Output.docx");
    } catch (error) {
      alert("Error generating document: " + error.message);
      console.error(error);
    }
  }
};