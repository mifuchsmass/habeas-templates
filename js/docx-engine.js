/** Word adapter: edit visible text in its existing runs; keep untouched package parts. */
const DocxEngine = {
  templateCache: {},
  STATE_FACILITIES: {
    Massachusetts: [{ value: 'Plymouth', text: 'Plymouth (Plymouth County Correctional Facility)' }, { value: 'Burlington', text: 'Burlington (Burlington ICE Office)' }, { value: 'Boston', text: 'Boston (Boston Field Office)' }],
    'New Hampshire': [{ value: 'Berlin', text: 'Berlin (FCI Berlin)' }, { value: 'Strafford', text: 'Strafford (Strafford County House of Corrections)' }],
    'Rhode Island': [{ value: 'Wyatt', text: 'Wyatt (Donald W. Wyatt Detention Facility)' }],
    Vermont: [{ value: 'Northwest', text: 'Northwest (Northwest State Correctional Facility)' }]
  },
  updateFacilityDropdown(el, state) {
    if (!el) return;
    const choices = this.STATE_FACILITIES[state] || [];
    const otherOnly = state === 'Connecticut' || state === 'Maine';
    el.replaceChildren(...(otherOnly ? [] : [new Option('Select detention facility…', '')]), ...choices.map(o => new Option(o.text, o.value)));
    if (state) el.add(new Option('Other detention location…', '__other__'));
    if (state === 'Connecticut' || state === 'Maine') el.value = '__other__';
    this.updateOtherLocationVisibility(el, state);
  },
  updateOtherLocationVisibility(el = document.getElementById('detention_facility'), state = document.getElementById('jurisdiction')?.value) {
    if (el && ['Connecticut', 'Maine'].includes(state)) el.value = '__other__';
    const field = document.getElementById('other-detention-location-field');
    if (field) field.hidden = el?.value !== '__other__';
  },
  bindFacilityControls(jurisdiction, facility) {
    this.updateFacilityDropdown(facility, jurisdiction.value);
    jurisdiction.addEventListener('change', () => this.updateFacilityDropdown(facility, jurisdiction.value));
    facility.addEventListener('change', () => this.updateOtherLocationVisibility(facility));
  },
  async loadTemplate(path) {
    if (!this.templateCache[path]) {
      const response = await fetch(path);
      if (!response.ok) throw Error(`Failed to load template: ${path}`);
      this.templateCache[path] = await response.arrayBuffer();
    }
    return this.templateCache[path].slice(0);
  },
  ns: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  xml(raw, label) {
    const doc = new DOMParser().parseFromString(raw, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw Error(`Invalid Word XML: ${label}`);
    return doc;
  },
  /** Read-only checks for the two known preview sinks; not a DOCX sanitizer. */
  validatePreview(zip) {
    const blocked = message => { throw Error(`Preview blocked: ${message} Download remains available.`); };
    // The pinned renderer reads attributes by local name, irrespective of prefix.
    const attr = (el, name) => {
      const matches = [...el.attributes].filter(a => a.localName === name);
      if (matches.length > 1) blocked('The document contains ambiguous symbol or link information.');
      return matches[0]?.value ?? null;
    };
    const docs = new Map(), relationships = new Map(), mainParts = new Set(['word/document.xml']), renderParts = new Set(['word/document.xml']);
    const read = name => {
      if (!docs.has(name)) {
        try { docs.set(name, this.xml(zip.file(name).asText(), name)); }
        catch { blocked('The document contains unreadable XML information.'); }
      }
      return docs.get(name);
    };
    const resolve = (owner, target) => {
      try {
        // Match docx-preview's package-path resolution, including encoded names.
        return new URL(target, `http://docx/${owner}`).toString().slice('http://docx/'.length);
      } catch { return null; }
    };
    for (const name of Object.keys(zip.files).filter(n => n.endsWith('.rels'))) {
      const doc = read(name), root = doc.documentElement;
      if (root.localName !== 'Relationships') blocked('The document contains malformed link relationships.');
      const owner = name === '_rels/.rels' ? '' : name.replace(/(^|\/)_rels\//, '$1').replace(/\.rels$/, '');
      const ids = new Set(), records = [];
      for (const el of root.children) {
        if (el.localName !== 'Relationship') blocked('The document contains malformed link relationships.');
        const id = attr(el, 'Id'), type = attr(el, 'Type'), target = attr(el, 'Target'), mode = attr(el, 'TargetMode');
        if (!id || !type || !target || ids.has(id) || mode && !['Internal', 'External'].includes(mode)) blocked('The document contains incomplete or ambiguous link relationships.');
        ids.add(id); records.push({ id, type, target, mode });
        // Include renderer-loaded XML even when its filename has no .xml suffix.
        if (/\/(officeDocument|header|footer|footnotes|endnotes|comments)$/.test(type)) {
          const path = owner === '' && zip.file(target) ? target : resolve(owner, target);
          if (!path || !zip.file(path)) {
            if (mode === 'External') continue; // No package XML is loaded for this resource.
            blocked('The document contains an invalid document-part relationship.');
          }
          renderParts.add(path);
          if (/\/officeDocument$/.test(type)) mainParts.add(path);
          read(path);
        }
      }
      relationships.set(owner, records);
    }
    for (const name of Object.keys(zip.files).filter(n => n.endsWith('.xml'))) read(name);
    const links = new Map(), allLinkIds = new Set();
    for (const [name, doc] of docs) {
      if (!renderParts.has(name) && !name.startsWith('word/') && !['document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments'].includes(doc.documentElement.localName)) continue;
      const ids = new Set(); links.set(name, ids);
      for (const el of doc.getElementsByTagName('*')) {
        if (el.localName === 'sym' && !/^[0-9a-fA-F]{4}$/.test(attr(el, 'char') ?? '')) blocked('A Word symbol has an invalid character code.');
        if (el.localName === 'hyperlink') {
          const id = attr(el, 'id');
          if (id) { ids.add(id); allLinkIds.add(id); }
        }
      }
    }
    for (const [owner, records] of relationships) for (const rel of records) {
      // The renderer can resolve hyperlink IDs from main-document relationships
      // without checking their Type; do not let a mislabeled link bypass checks.
      const usedAsLink = links.get(owner)?.has(rel.id) || mainParts.has(owner) && allLinkIds.has(rel.id);
      if (rel.mode !== 'External' || !(/\/hyperlink$/.test(rel.type) || usedAsLink)) continue;
      const target = rel.target;
      if (/[\u0000-\u001f\u007f]/.test(target) || /%(?![0-9a-fA-F]{2})/.test(target)) blocked('A document link has a malformed address.');
      const trimmed = target.trim();
      if (trimmed.startsWith('#') && trimmed.length > 1) continue;
      let url;
      try { url = new URL(target); } catch { blocked('A document link has an unsupported or malformed address.'); }
      if (!['http:', 'https:', 'mailto:'].includes(url.protocol) ||
          (url.protocol === 'mailto:' ? !url.pathname : !url.hostname)) blocked('A document link uses an unsupported address type.');
    }
  },
  descendants(el, local) { return [...el.getElementsByTagNameNS(this.ns, local)]; },
  ancestor(el, local) {
    while (el && el.nodeType === 1) { if (el.namespaceURI === this.ns && el.localName === local) return el; el = el.parentNode; }
    return null;
  },
  visible(el) {
    return [...el.getElementsByTagName('*')].filter(n => ['t', 'delText', 'instrText', 'br', 'cr', 'tab'].includes(n.localName))
      .map(n => ['br', 'cr'].includes(n.localName) ? '\u000b' : n.localName === 'tab' ? '\t' : n.textContent).join('');
  },
  automation(text) { return /<[^<>]*>|\{\s*(?:IF|ELSE(?:IF)?|END(?:IF)?)\b/i.test(text); },
  inspect(buffer) {
    const zip = new PizZip(buffer), parts = [], fields = new Map();
    if (!zip.file('word/document.xml')) throw Error('The file is not a supported Word DOCX.');
    for (const name of Object.keys(zip.files)) {
      if (!name.startsWith('word/') || !name.endsWith('.xml')) continue;
      const raw = zip.file(name).asText(), doc = this.xml(raw, name);
      if (!['word/document.xml', 'word/footnotes.xml'].includes(name)) {
        if (this.automation(this.visible(doc))) throw Error(`${name}: template fields/instructions in this Word part are unsupported.`);
        continue;
      }
      const roots = name === 'word/footnotes.xml' ? this.descendants(doc, 'footnote') : [this.descendants(doc, 'body')[0]];
      const stories = [];
      for (const root of roots) {
        if (!root) throw Error(`${name}: missing document body.`);
        const id = root.getAttributeNS(this.ns, 'id');
        const location = name === 'word/footnotes.xml' ? `Footnote ${id}` : 'Main document';
        // Complex structures can pass through only without active automation.
        for (const el of [...root.getElementsByTagNameNS(this.ns, '*')]) {
          const nestedTable = el.localName === 'tbl' && this.ancestor(el.parentNode, 'tbl');
          if ((nestedTable || ['txbxContent', 'drawing', 'pict', 'object', 'ins', 'del', 'moveFrom', 'moveTo', 'sdt', 'fldSimple', 'hyperlink'].includes(el.localName)) && this.automation(this.visible(el))) {
            throw Error(`${location}: automation inside ${nestedTable ? 'nested tables' : el.localName} is unsupported.`);
          }
        }
        const paragraphs = this.descendants(root, 'p').filter(p => !this.ancestor(p.parentNode, 'txbxContent') && !this.ancestor(p.parentNode, 'p'));
        let text = ''; const records = [];
        for (const p of paragraphs) {
          const start = text.length, spans = [];
          if ((this.descendants(p, 'instrText').length || this.descendants(p, 'fldChar').length) && this.automation(this.visible(p))) throw Error(`${location}: automation in a paragraph with Word field codes is unsupported.`);
          for (const n of [...p.getElementsByTagNameNS(this.ns, '*')]) {
            if (n.localName === 'footnoteReference') { spans.push({ node: n, start: text.length, end: text.length, value: '', reference: true }); continue; }
            if (!['t', 'br', 'cr', 'tab'].includes(n.localName)) continue;
            const value = n.localName === 't' ? n.textContent : ['br', 'cr'].includes(n.localName) ? '\u000b' : '\t';
            spans.push({ node: n, start: text.length, end: text.length + value.length, value }); text += value;
          }
          records.push({ node: p, start, end: text.length, spans }); text += '\n';
        }
        const parsed = TemplateLogic.parse(text, location);
        const story = { root, id, location, paragraphs: records, parsed };
        this.validateStructures(story);
        for (const f of parsed.fields) if (!fields.has(TemplateLogic.key(f))) fields.set(TemplateLogic.key(f), f);
        stories.push(story);
      }
      parts.push({ name, raw, doc, stories });
    }
    return { zip, parts, fields: [...fields.values()] };
  },
  validateStructures(story) {
    const locate = pos => story.paragraphs.find(p => p.start <= pos && p.end >= pos);
    for (const t of story.parsed.tokens) {
      const p = locate(t.start);
      if (!p || t.end > p.end) throw Error(`${story.location}: instruction or field crosses paragraphs.`);
      const unsafe = this.descendants(p.node, 'instrText').length || this.descendants(p.node, 'fldChar').length;
      if (unsafe) throw Error(`${story.location}: automation in a paragraph with Word field codes is unsupported.`);
    }
    function blocks(nodes) { return nodes.flatMap(n => n.kind === 'block' ? [n, ...n.branches.flatMap(b => blocks(b.nodes))] : []); }
    for (const block of blocks(story.parsed.nodes)) {
      const tagParas = [...block.branches.map(b => locate(b.tag.start)), locate(block.close.start)];
      const cells = tagParas.map(p => this.ancestor(p.node, 'tc'));
      if (cells.some(c => c !== cells[0])) {
        // Whole-row controls: marker rows contain only that instruction; ordinary paragraphs outside tables also work.
        for (let i = 0; i < tagParas.length; i++) {
          const row = this.ancestor(tagParas[i].node, 'tr');
          const tag = i === tagParas.length - 1 ? block.close : block.branches[i].tag;
          if (row && this.visible(row).trim() !== tag.raw.trim()) throw Error(`${tag.location}: crossing table cells requires instructions on their own otherwise empty rows.`);
        }
      }
      const startP = locate(block.start), endP = locate(block.end - 1);
      for (const table of this.descendants(story.root, 'tbl')) {
        if (!this.descendants(table, 'vMerge').length && !this.descendants(table, 'hMerge').length) continue;
        const ps = story.paragraphs.filter(p => this.ancestor(p.node, 'tbl') === table);
        if (!ps.length || ps.at(-1).end < block.start || ps[0].start > block.end) continue;
        if (block.start > ps[0].start || block.end < ps.at(-1).end) throw Error(`${story.location}: conditional control of part of a table with merged rows/cells is unsupported.`);
      }
      for (const p of story.paragraphs.filter(p => p.start <= block.end && p.end >= block.start)) {
        if (['drawing', 'pict', 'object', 'sym'].some(n => this.descendants(p.node, n).length) && (p.start < block.start || p.end > block.end)) {
          throw Error(`${story.location}: partial conditional control in a paragraph containing drawings or symbols is unsupported.`);
        }
      }
      // Reject crossing structures whose range semantics cannot safely follow conditional deletion.
      for (const p of story.paragraphs.filter(p => p.start >= startP.start && p.end <= endP.end)) {
        if (['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'ins', 'del', 'moveFrom', 'moveTo', 'sdt', 'sectPr'].some(n => this.descendants(p.node, n).length)) {
          throw Error(`${story.location}: conditional block intersects bookmarks, comments, tracked changes, content controls, or a section break.`);
        }
      }
    }
  },
  extractCustomFields(buffer) {
    if (typeof buffer === 'string') return TemplateLogic.parse(buffer).fields;
    return this.inspect(buffer).fields;
  },
  apply(story, plan) {
    const edits = plan.edits;
    if (!edits.length) return false;
    const removed = new Set(), retainedMarkers = [];
    const rowParagraphs = new Map(this.descendants(story.root, 'tr').map(row => [row, story.paragraphs.filter(p => this.ancestor(p.node, 'tr') === row)]));
    let markerParagraph;
    for (const p of story.paragraphs) {
      const relevant = edits.filter(e => e.start <= p.end && e.end >= p.start);
      if (!relevant.length) continue;
      const deletions = relevant.filter(e => !e.replacement);
      let leftover = story.parsed.text.slice(p.start, p.end);
      for (const e of [...deletions].reverse()) leftover = leftover.slice(0, Math.max(0, e.start - p.start)) + leftover.slice(Math.max(0, e.end - p.start));
      const structural = relevant.some(e => ['exclude', 'instruction'].includes(e.type));
      const whollyExcluded = relevant.some(e => e.type === 'exclude' && e.start <= p.start && e.end >= p.end);
      const substantive = ['drawing', 'pict', 'object', 'footnoteReference', 'endnoteReference', 'sym'].some(n => this.descendants(p.node, n).length);
      const fullyRemoved = whollyExcluded || structural && !leftover.trim() && !substantive;
      if (fullyRemoved) {
        for (const marker of this.descendants(p.node, 'footnoteRef')) {
          const run = this.ancestor(marker, 'r').cloneNode(true);
          for (const child of [...run.children]) if (!['rPr', 'footnoteRef'].includes(child.localName)) child.remove();
          retainedMarkers.push(run); markerParagraph = p.node;
        }
        if (this.descendants(p.node, 'sectPr').length) throw Error(`${story.location}: cannot remove a section break.`);
        removed.add(p.node); continue;
      }
      if (structural && ['drawing', 'pict', 'object'].some(n => this.descendants(p.node, n).length)) throw Error(`${story.location}: partial conditional deletion in a paragraph containing a drawing is unsupported.`);
      for (const span of p.spans) {
        if (span.reference) {
          if (relevant.some(e => e.type === 'exclude' && e.start <= span.start && e.end >= span.start)) span.node.remove();
          continue;
        }
        let value = span.value;
        const touching = relevant.filter(e => e.start < span.end && e.end > span.start || e.start === span.start && span.end === span.start);
        for (const e of [...touching].reverse()) {
          const a = Math.max(0, e.start - span.start), b = Math.min(span.value.length, e.end - span.start);
          const replacement = e.start >= span.start && e.start < span.end ? e.replacement : '';
          value = value.slice(0, a) + replacement + value.slice(b);
        }
        if (value === span.value) continue;
        if (span.node.localName !== 't') {
          if (!value) span.node.remove();
          else { const t = span.node.ownerDocument.createElementNS(this.ns, 'w:t'); t.textContent = value; t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); span.node.replaceWith(t); }
        } else {
          const chunks = value.split(/(\r\n|\r|\n|\t)/);
          span.node.textContent = chunks.shift(); span.node.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
          let previous = span.node;
          for (const chunk of chunks) {
            const el = span.node.ownerDocument.createElementNS(this.ns, chunk === '\t' ? 'w:tab' : /^[\r\n]+$/.test(chunk) ? 'w:br' : 'w:t');
            if (el.localName === 't') { el.textContent = chunk; el.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); }
            previous.after(el); previous = el;
          }
        }
      }
    }
    for (const p of removed) p.remove();
    if (retainedMarkers.length) {
      let first = this.descendants(story.root, 'p')[0];
      if (!first) { first = story.root.ownerDocument.createElementNS(this.ns, 'w:p'); story.root.append(first); }
      const properties = markerParagraph && [...markerParagraph.children].find(n => n.localName === 'pPr');
      if (properties && ![...first.children].some(n => n.localName === 'pPr')) first.prepend(properties.cloneNode(true));
      const anchor = [...first.children].find(n => n.localName !== 'pPr') || null;
      for (const run of retainedMarkers) first.insertBefore(run, anchor);
    }
    // Remove fully excluded rows; retain a required empty paragraph in surviving empty cells.
    for (const row of this.descendants(story.root, 'tr')) {
      const original = rowParagraphs.get(row) || [];
      if (original.length && original.every(p => removed.has(p.node))) row.remove();
    }
    for (const table of this.descendants(story.root, 'tbl')) if (!this.descendants(table, 'tr').length) table.remove();
    for (const cell of this.descendants(story.root, 'tc')) if (!this.descendants(cell, 'p').length) cell.appendChild(cell.ownerDocument.createElementNS(this.ns, 'w:p'));
    return true;
  },
  compile(buffer, data) {
    const model = this.inspect(buffer), values = TemplateLogic.values(data), missing = new Map(), unresolved = [];
    const warnings = plan => {
      for (const f of plan.missing) missing.set(TemplateLogic.key(f), f);
      unresolved.push(...plan.unresolved);
    };
    // Validate every story first; evaluate facts only in surviving body text/footnotes.
    const body = model.parts.find(p => p.name === 'word/document.xml');
    for (const story of body.stories) { const plan = TemplateLogic.plan(story.parsed, values); warnings(plan); body.changed = this.apply(story, plan); }
    const referenced = new Set(this.descendants(body.doc, 'footnoteReference').map(n => n.getAttributeNS(this.ns, 'id')));
    for (const name of Object.keys(model.zip.files).filter(n => /^word\/(?:header|footer).*\.xml$/.test(n))) {
      const d = this.xml(model.zip.file(name).asText(), name);
      for (const n of this.descendants(d, 'footnoteReference')) referenced.add(n.getAttributeNS(this.ns, 'id'));
    }
    const foot = model.parts.find(p => p.name === 'word/footnotes.xml');
    if (foot) {
      for (const story of foot.stories) {
        if (story.root.hasAttributeNS(this.ns, 'type') && story.root.getAttributeNS(this.ns, 'type') !== 'normal') continue;
        if (!referenced.has(story.id)) { story.root.remove(); foot.changed = true; continue; }
        const plan = TemplateLogic.plan(story.parsed, values); warnings(plan); foot.changed = this.apply(story, plan) || foot.changed;
      }
    }
    const allIds = foot ? this.descendants(foot.doc, 'footnote').map(n => n.getAttributeNS(this.ns, 'id')) : [];
    const ids = new Set(allIds);
    if (ids.size !== allIds.length) throw Error('Duplicate footnote IDs in the document.');
    for (const id of referenced) if (!ids.has(id)) throw Error(`Broken footnote reference ${id}.`);
    for (const part of model.parts) if (part.changed) model.zip.file(part.name, new XMLSerializer().serializeToString(part.doc));
    const blob = model.zip.generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    return { blob, zip: model.zip, warnings: { missing: [...missing.values()], unresolved } };
  },
  async generate({ templatePath, buffer, data }) {
    return this.compile(buffer || await this.loadTemplate(templatePath), data);
  },
  download(blob, filename) {
    if (typeof saveAs === 'function') { saveAs(blob, filename); return; }
    const url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = filename; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
};

window.showToast = function(message, type = 'success', duration = 3000) {
  let container = document.getElementById('toast-container');
  if (!container) { container = document.createElement('div'); container.id = 'toast-container'; document.body.append(container); }
  const toast = document.createElement('div'); toast.className = `toast-notification toast-${type}`; toast.textContent = message; container.append(toast);
  setTimeout(() => toast.remove(), duration);
};

/** Preview and Download share status ownership, but independently compile current data. */
window.GenerationUI = {
  serial: 0,
  panel() { return document.getElementById('generation-status'); },
  status(text, result) {
    const panel = this.panel(); if (!panel) return;
    if (result) {
      panel.replaceChildren();
      const summary = document.createElement('p'); summary.className = 'generation-summary'; panel.append(summary);
      const w = result.warnings;
      const count = (n, singular, plural) => `${n} ${n === 1 ? singular : plural}`;
      panel.classList.toggle('review-required', !!(w.unresolved.length || w.missing.length));
      summary.textContent = w.unresolved.length
        ? `Attorney review required — ${count(w.unresolved.length, 'condition', 'conditions')} could not be resolved`
        : w.missing.length
          ? `Document created — ${count(w.missing.length, 'blank field remains', 'blank fields remain')}`
          : 'Document created — no unresolved conditions or blank fields.';
      const explanation = document.createElement('p');
      explanation.textContent = w.unresolved.length
        ? 'The conditional instructions and alternatives are preserved in the document for attorney review.'
        : w.missing.length ? 'Blank fields remain as <Field Name> placeholders in the document.' : 'Document creation does not establish readiness for filing; attorney review is still required.';
      panel.append(explanation);
      if (w.unresolved.length && w.missing.length) {
        const note = document.createElement('p'); note.textContent = 'Blank fields remain as <Field Name> placeholders in the document.'; panel.append(note);
      }
      if (w.unresolved.length || w.missing.length) {
        const note = document.createElement('p'); note.textContent = 'Document generated. Review it carefully before filing.'; panel.append(note);
      }
      for (const [label, details, expanded] of [
        [count(w.unresolved.length, 'unresolved condition', 'unresolved conditions'), w.unresolved.map(u => `${u.location}: ${u.instruction}${u.reasons.length ? ' — ' + u.reasons.join('; ') : ''}${u.unevaluated.length ? '\n' + u.unevaluated.join('\n') : ''}`), true],
        [count(w.missing.length, 'blank field', 'blank fields'), w.missing.map(f => `<${f}>`), false]
      ]) {
        if (!details.length) continue;
        const d = document.createElement('details'), s = document.createElement('summary'), list = document.createElement('ul');
        d.open = expanded; s.textContent = label; list.className = 'warning-details';
        for (const text of details) { const li = document.createElement('li'); li.textContent = text; list.append(li); }
        d.append(s, list); panel.append(d);
      }
    }
    let state = panel.querySelector('.generation-state');
    if (!state) { state = document.createElement('p'); state.className = 'generation-state'; panel.prepend(state); }
    state.textContent = text;
  },
  async run(download, options) {
    const serial = ++this.serial, isLatest = () => serial === this.serial;
    if (!download) document.getElementById('docx-preview-container')?.replaceChildren();
    this.status(download ? 'Preparing Word document for download…' : 'Generating current form preview…');
    try {
      // Read form synchronously at start, so later edits belong to later attempts.
      const { formData, fullName } = options.collect();
      const result = await DocxEngine.generate({ ...options, data: formData });
      if (download) {
        DocxEngine.download(result.blob, `Habeas Petition ${fullName || 'Draft'}.docx`);
        if (isLatest()) this.status('Word document download prepared.', result);
      } else {
        if (isLatest()) this.status('Word document created; preparing preview…', result);
        const staging = document.createElement('div');
        try {
          DocxEngine.validatePreview(result.zip);
          if (!window.docx?.renderAsync) throw Error('Preview library is not loaded.');
          await window.docx.renderAsync(result.blob, staging);
          if (isLatest()) { document.getElementById('docx-preview-container').replaceChildren(...staging.childNodes); this.status('Preview ready.'); }
        } catch (e) {
          if (isLatest()) {
            document.getElementById('docx-preview-container')?.replaceChildren();
            this.status(`Word document created, but the preview could not be displayed: ${e.message}`);
          }
        }
      }
      return result;
    } catch (e) {
      if (isLatest()) { const panel = this.panel(); if (panel) panel.replaceChildren(); this.status(`Generation failed: ${e.message}`); }
      return null;
    }
  }
};

window.FormStorage = {
  requireReady() {
    if (typeof requirePleadingReady === 'function') requirePleadingReady();
    if (typeof requireStudioReady === 'function') requireStudioReady();
  },
  controls() {
    const studio = document.getElementById('sandbox');
    return [...(studio || document).querySelectorAll('input, select, textarea')].filter(el => el.type !== 'file' && el.type !== 'password' && el.id !== 'snippetInput');
  },
  name(el) { return el.getAttribute('data-dynamic-tag') || el.getAttribute('data-form-field') || el.name || el.id; },
  snapshot() {
    this.requireReady();
    const data = Object.create(null);
    for (const el of this.controls()) {
      const k = this.name(el); if (!k || el.type === 'radio' && !el.checked) continue;
      const value = el.type === 'checkbox' ? el.checked : el.value;
      if (Object.hasOwn(data, k) && !Object.is(data[k], value)) throw Error(`Conflicting saved values for <${k}>.`);
      data[k] = value;
    }
    TemplateLogic.values(data); return { format: 'habeas-form-v3', values: data };
  },
  restore(saved) {
    this.requireReady();
    if (saved?.format !== 'habeas-form-v3' || !saved.values || Array.isArray(saved.values) || typeof saved.values !== 'object') throw Error('This is not a form data file saved by the current version of the Habeas application.');
    const data = TemplateLogic.values(saved.values), controls = this.controls();
    for (const [name, value] of data) if (typeof value === 'number' && (!Number.isFinite(value) || Math.abs(value) > Number.MAX_SAFE_INTEGER)) throw Error(`Saved number for <${name}> is unreliable; save exact text instead.`);
    const jurisdiction = data.get(TemplateLogic.key('jurisdiction'));
    if (jurisdiction !== undefined) DocxEngine.updateFacilityDropdown(document.getElementById('detention_facility'), jurisdiction);
    for (const el of controls) {
      const name = this.name(el); if (!name || !data.has(TemplateLogic.key(name))) continue;
      const value = data.get(TemplateLogic.key(name));
      if (el.type === 'radio') el.checked = el.value === value;
      else if (el.type === 'checkbox') el.checked = value === true;
      else el.value = value == null ? '' : String(value);
    }
    const facility = document.getElementById('detention_facility');
    // Legacy CT/ME exports used N/A, which is not an actual detention location.
    if (facility && data.get(TemplateLogic.key('Detention Facility')) === 'N/A' && ['Connecticut', 'Maine'].includes(jurisdiction)) {
      facility.value = '__other__';
      const custom = document.getElementById('other_detention_location');
      if (custom) custom.value = '';
    }
    DocxEngine.updateOtherLocationVisibility(facility);
  },
  save(prefix = 'Habeas_Data') {
    try {
      const data = this.snapshot(), base = document.getElementById('petitioner_first_name')?.value.trim() || 'Draft';
      const filename = `${prefix.replace(/\.json$/i, '')}_${base.replace(/[^a-zA-Z0-9_-]/g, '_')}.json`;
      DocxEngine.download(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), filename);
      showToast(`Saved ${filename}`);
    } catch (e) { showToast(e.message, 'error'); }
  },
  openFileDialog() {
    try { this.requireReady(); }
    catch (e) { showToast(e.message, 'error'); return; }
    const input = document.createElement('input'); input.type = 'file'; input.accept = '.json';
    input.onchange = () => { if (input.files[0]) this.load(input.files[0]); }; input.click();
  },
  async load(file) {
    try { this.requireReady(); this.restore(JSON.parse(await file.text())); showToast('Form data loaded.'); }
    catch (e) { showToast(`Could not load form data: ${e.message}`, 'error'); }
  }
};
