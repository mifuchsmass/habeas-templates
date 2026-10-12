/** Shared, non-executable template language. Positions refer to visible story text. */
const TemplateLogic = (() => {
  const blank = v => v == null || (typeof v === 'string' && !v.trim());
  function key(name) {
    const n = name.trim();
    if (!n || /[<>\{\}\\\u0000-\u001f\u007f]/u.test(n) || n.length > 256) throw Error(`Unsupported field name: ${name}`);
    return n.toLowerCase(); // Locale-independent; spaces, hyphens and underscores remain distinct.
  }
  function values(data) {
    const map = new Map();
    for (const [name, value] of Object.entries(data)) {
      const k = key(name);
      if (value != null && !['string', 'number', 'boolean'].includes(typeof value)) throw Error(`Unsupported value for <${name}>: use text, a number, or true/false.`);
      if (map.has(k) && !Object.is(map.get(k), value)) throw Error(`Conflicting values for field <${name}>.`);
      map.set(k, value);
    }
    return map;
  }
  function decimal(raw) {
    const s = String(raw).trim();
    if (s.length > 1000 || !/^-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(s)) throw Error('Use a decimal number with correctly grouped thousands commas (maximum 1,000 characters).');
    const [whole, frac = ''] = s.replace(/,/g, '').split('.');
    return { integer: BigInt(whole + frac), scale: frac.length };
  }
  function numberCompare(a, b) {
    const scale = Math.max(a.scale, b.scale);
    const x = a.integer * 10n ** BigInt(scale - a.scale), y = b.integer * 10n ** BigInt(scale - b.scale);
    return x < y ? -1 : x > y ? 1 : 0;
  }
  function expression(source) {
    const tokens = []; let i = 0;
    while (i < source.length) {
      if (/\s/.test(source[i])) { i++; continue; }
      const start = i, c = source[i];
      if (tokens.at(-1)?.type !== 'value' && c === '<' && !/\s|=/.test(source[i + 1] || ' ')) {
        const end = source.indexOf('>', i + 1);
        if (end < 0) throw Error('Unclosed field reference in condition.');
        const name = source.slice(i + 1, end); key(name);
        tokens.push({ type: 'value', field: name.trim() }); i = end + 1; continue;
      }
      // A field may include leading spaces: distinguish by its complete closing bracket.
      if (tokens.at(-1)?.type !== 'value' && c === '<' && source[i + 1] !== '=' && /^<\s+[^<>]+>/.test(source.slice(i))) {
        const end = source.indexOf('>', i + 1), name = source.slice(i + 1, end); key(name);
        tokens.push({ type: 'value', field: name.trim() }); i = end + 1; continue;
      }
      if (c === '"' || c === '“') {
        const close = c === '“' ? '”' : '"'; let text = ''; i++;
        while (i < source.length && source[i] !== close) {
          if (source[i] === '\\') {
            i++; if (!['\\', '"', '”'].includes(source[i])) throw Error('Only escaped quotes and backslashes are supported inside quoted text.');
          }
          text += source[i++];
        }
        if (source[i] !== close) throw Error('Unclosed quoted text.');
        i++; tokens.push({ type: 'value', text }); continue;
      }
      const op = /^(?:!=|<=|>=|=|<|>)/.exec(source.slice(i));
      if (op) {
        if (!/\s/.test(source[i - 1] || '') || !/\s/.test(source[i + op[0].length] || '')) throw Error('Put spaces on both sides of each comparison symbol.');
        tokens.push({ type: 'compare', op: op[0] }); i += op[0].length; continue;
      }
      if (c === '(' || c === ')') { tokens.push({ type: c }); i++; continue; }
      const word = /^(AND|OR|NOT)\b/i.exec(source.slice(i));
      if (word) { tokens.push({ type: word[0].toUpperCase() }); i += word[0].length; continue; }
      const num = /^-?[\d][\d,.]*(?![\w.])/.exec(source.slice(i));
      if (num) {
        const n = decimal(num[0]); tokens.push({ type: 'value', number: n }); i += num[0].length; continue;
      }
      throw Error(`Invalid condition near “${source.slice(start, start + 30)}”. Text must be quoted; fields need angle brackets.`);
    }
    let pos = 0;
    const take = type => tokens[pos]?.type === type ? tokens[pos++] : null;
    function comparison() {
      if (take('(')) { const n = or(); if (!take(')')) throw Error('Missing closing parenthesis.'); return n; }
      const a = take('value'), op = take('compare'), b = take('value');
      if (!a || !op || !b) throw Error('A condition needs two values and a comparison.');
      const numeric = ['<', '>', '<=', '>='].includes(op.op) || a.number !== undefined || b.number !== undefined;
      if (numeric && (a.text !== undefined || b.text !== undefined)) throw Error('Numerical comparisons require fields or unquoted numbers.');
      return { kind: 'compare', a, b, op: op.op, numeric };
    }
    function not() { return take('NOT') ? { kind: 'NOT', child: not() } : comparison(); }
    function and() { let n = not(); while (take('AND')) n = { kind: 'AND', left: n, right: not() }; return n; }
    function or() { let n = and(); while (take('OR')) n = { kind: 'OR', left: n, right: and() }; return n; }
    const result = or(); if (pos !== tokens.length) throw Error('Unexpected text after condition.');
    return result;
  }
  function evaluate(node, data) {
    if (node.kind === 'NOT') { const r = evaluate(node.child, data); return { ...r, value: r.value === null ? null : !r.value }; }
    if (node.kind === 'AND' || node.kind === 'OR') {
      const a = evaluate(node.left, data), decisive = node.kind === 'AND' ? false : true;
      if (a.value === decisive) return { value: decisive, reasons: [] };
      const b = evaluate(node.right, data);
      if (b.value === decisive) return { value: decisive, reasons: [] };
      return { value: a.value === null || b.value === null ? null : !decisive, reasons: [...a.reasons, ...b.reasons] };
    }
    const reasons = [];
    const get = v => {
      if (!v.field) return v.number ?? v.text;
      const value = data.get(key(v.field));
      if (blank(value)) { reasons.push(`Missing <${v.field}>`); return null; }
      if (!node.numeric) return String(value);
      if (typeof value === 'number' && (!Number.isFinite(value) || Math.abs(value) > Number.MAX_SAFE_INTEGER)) { reasons.push(`Unreliable numeric value in <${v.field}>: enter it as exact text.`); return null; }
      try { return decimal(value); } catch (e) { reasons.push(`Invalid number in <${v.field}>: ${e.message}`); return null; }
    };
    const a = get(node.a), b = get(node.b);
    if (a === null || b === null) return { value: null, reasons };
    const norm = v => String(v).trim().toLowerCase();
    const cmp = node.numeric ? numberCompare(a, b) : norm(a) === norm(b) ? 0 : norm(a) < norm(b) ? -1 : 1;
    return { value: ({ '=': cmp === 0, '!=': cmp !== 0, '<': cmp < 0, '>': cmp > 0, '<=': cmp <= 0, '>=': cmp >= 0 })[node.op], reasons: [] };
  }
  function parse(text, location = 'Text') {
    const tokens = [], fields = new Map(); let i = 0;
    const addExpressionFields = n => {
      if (n.kind === 'compare') for (const v of [n.a, n.b]) { if (v.field) fields.set(key(v.field), v.field); }
      else if (n.child) addExpressionFields(n.child);
      else { addExpressionFields(n.left); addExpressionFields(n.right); }
    };
    const fail = (message, at = i) => {
      const error = Error(`${location}, paragraph ${text.slice(0, at).split('\n').length}: ${message}`);
      error.context = text.slice(Math.max(0, at - 50), Math.min(text.length, at + 150));
      throw error;
    };
    while (i < text.length) {
      const start = i, escaped = text[i] === '\\' && (text[i + 1] === '<' || /^\{\s*(?:IF|ELSE(?:IF)?|END(?:IF)?)\b/i.test(text.slice(i + 1)));
      if (escaped) i++;
      if (/^\{\s*(?:IF|ELSE(?:IF)?|END(?:IF)?)\b/i.test(text.slice(i))) {
        // Closing braces inside quoted literals are allowed.
        let end = i + 1, quote = null;
        for (; end < text.length; end++) {
          const c = text[end];
          if (!quote && c === '<') {
            const field = /^<[^<>{}\r\n]+>/.exec(text.slice(end));
            if (field) { end += field[0].length - 1; continue; }
          }
          if (quote && c === '\\') { end++; continue; }
          if (quote && c === quote) quote = null;
          else if (!quote && (c === '"' || c === '“')) quote = c === '“' ? '”' : c;
          else if (!quote && c === '}') break;
        }
        if (end === text.length || text.slice(i, end).includes('\n')) fail('Each complete instruction must be in one paragraph.', start);
        const raw = text.slice(i, end + 1), body = raw.slice(1, -1).trim();
        if (escaped) tokens.push({ kind: 'escape', start, end: end + 1, replacement: raw });
        else {
          let kind, expr;
          if (/^IF\s+/i.test(body)) { kind = 'if'; expr = body.replace(/^IF\s+/i, ''); }
          else if (/^ELSE\s+IF\s+/i.test(body)) { kind = 'elseif'; expr = body.replace(/^ELSE\s+IF\s+/i, ''); }
          else if (/^ELSE$/i.test(body)) kind = 'else';
          else if (/^END\s+IF$/i.test(body)) kind = 'end';
          else fail('Malformed conditional instruction.', start);
          let condition;
          if (expr) { try { condition = expression(expr); addExpressionFields(condition); } catch (e) { fail(e.message, start); } }
          tokens.push({ kind, condition, raw, start, end: end + 1, location: `${location}, paragraph ${text.slice(0, start).split('\n').length}` });
        }
        i = end + 1; continue;
      }
      if (text[i] === '<') {
        const end = text.indexOf('>', i + 1);
        if (end < 0 || text.slice(i, end).includes('\n')) fail('Unclosed or paragraph-split field reference.', start);
        const raw = text.slice(i, end + 1), name = raw.slice(1, -1).trim();
        if (escaped) tokens.push({ kind: 'escape', start, end: end + 1, replacement: raw });
        else { try { fields.set(key(name), name); } catch (e) { fail(e.message, start); } tokens.push({ kind: 'field', name, raw, start, end: end + 1 }); }
        i = end + 1; continue;
      }
      if (escaped) fail('Invalid escape.', start);
      i++;
    }
    let p = 0;
    function sequence(inBlock = false) {
      const nodes = [];
      while (p < tokens.length) {
        const t = tokens[p];
        if (['elseif', 'else', 'end'].includes(t.kind)) { if (!inBlock) fail(`Orphan ${t.raw}.`, t.start); break; }
        p++;
        if (t.kind !== 'if') { nodes.push(t); continue; }
        const block = { kind: 'block', start: t.start, branches: [{ tag: t, nodes: sequence(true) }] }; let hasElse = false;
        while (tokens[p] && ['elseif', 'else'].includes(tokens[p].kind)) {
          const tag = tokens[p++];
          if (hasElse) fail('ELSE IF or duplicate ELSE after ELSE.', tag.start);
          hasElse = tag.kind === 'else'; block.branches.push({ tag, nodes: sequence(true) });
        }
        const end = tokens[p++]; if (!end || end.kind !== 'end') fail('IF is missing END IF.', t.start);
        block.close = end; block.end = end.end; nodes.push(block);
      }
      return nodes;
    }
    return { text, nodes: sequence(), fields: [...fields.values()], tokens, location };
  }
  function plan(parsed, data) {
    const edits = [], missing = new Map(), unresolved = [];
    const replace = (start, end, replacement = '', type = 'text') => edits.push({ start, end, replacement, type });
    function ordinary(n) {
      if (n.kind === 'escape') replace(n.start, n.end, n.replacement);
      if (n.kind === 'field') {
        const value = data.get(key(n.name));
        if (blank(value)) missing.set(key(n.name), n.name);
        else {
          const text = String(value);
          if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]|[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/u.test(text)) throw Error(`Value for <${n.name}> contains characters that cannot be safely written to Word XML.`);
          replace(n.start, n.end, text);
        }
      }
    }
    function preserve(nodes, parent, detail) {
      for (const n of nodes) {
        if (n.kind !== 'block') ordinary(n);
        else {
          detail.unevaluated.push(`${n.branches[0].tag.location}: ${n.branches[0].tag.raw} (not evaluated because ${parent} is unresolved)`);
          for (const b of n.branches) preserve(b.nodes, parent, detail);
        }
      }
    }
    function visit(nodes) {
      for (const n of nodes) {
        if (n.kind !== 'block') { ordinary(n); continue; }
        let selected = -1, unknown = null;
        for (let b = 0; b < n.branches.length; b++) {
          const tag = n.branches[b].tag;
          const r = tag.kind === 'else' ? { value: true } : evaluate(tag.condition, data);
          if (r.value === null) { unknown = { instruction: tag.raw, location: tag.location, reasons: r.reasons, unevaluated: [] }; break; }
          if (r.value) { selected = b; break; }
        }
        if (unknown) { unresolved.push(unknown); for (const b of n.branches) preserve(b.nodes, unknown.instruction, unknown); continue; }
        for (let b = 0; b < n.branches.length; b++) {
          const branch = n.branches[b], end = n.branches[b + 1]?.tag.start ?? n.close.start;
          replace(branch.tag.start, branch.tag.end, '', 'instruction');
          if (b === selected) visit(branch.nodes);
          else replace(branch.tag.end, end, '', 'exclude');
        }
        replace(n.close.start, n.close.end, '', 'instruction');
      }
    }
    visit(parsed.nodes);
    return { edits: edits.sort((a, b) => a.start - b.start), missing: [...missing.values()], unresolved };
  }
  return { key, blank, values, decimal, expression, evaluate, parse, plan };
})();
if (typeof module !== 'undefined') module.exports = TemplateLogic;
