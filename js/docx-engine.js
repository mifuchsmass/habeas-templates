/**
 * Universal DOCX Template Engine
 */
const DocxEngine = {
  // 1. Repair split XML and translate [IF ...] / [END IF]
  parseIfTags(xmlContent) {
    xmlContent = xmlContent.replace(/\[[^\]]*?\]/g, match => match.replace(/<[^>]+>/g, ''));
    const tagRegex = /\[\s*(IF\s+[^\]]+|END\s+IF)\s*\]/gi;
    const stack = [];

    return xmlContent.replace(tagRegex, (match, tagBody) => {
      const trimmed = tagBody.trim();
      if (/^IF\s+/i.test(trimmed)) {
        let condition = trimmed.replace(/^IF\s+/i, '').trim()
          .replace(/[\u201C\u201D]/g, '"')
          .replace(/[\u2018\u2019]/g, "'")
          .replace(/\band\b/gi, '&&')
          .replace(/\bor\b/gi, '||')
          .replace(/\s+=\s+/g, ' == ');

        stack.push(condition);
        return `[#${condition}]`;
      } else if (/^END\s+IF$/i.test(trimmed)) {
        const lastCondition = stack.pop();
        return lastCondition ? `[/${lastCondition}]` : match;
      }
      return match;
    });
  },

  // 2. Main Generation Function
  async generate({ templatePath, data, outputFilename, selectElementForAliases }) {
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
                return (val === "" || val === undefined || val === null) ? "*** MISSING DATA ***" : val;
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

      saveAs(out, outputFilename || "Document_Output.docx");
    } catch (error) {
      alert("Error generating document: " + error.message);
      console.error(error);
    }
  }
};