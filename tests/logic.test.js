const assert = require('node:assert/strict');
const L = require('../js/template-logic.js');
let checks = 0;
function test(name, fn) { try { fn(); checks++; } catch (e) { e.message = `${name}: ${e.message}`; throw e; } }
const evalCondition = (s, data = {}) => L.evaluate(L.expression(s), L.values(data));
const render = (s, data = {}) => {
  const p = L.plan(L.parse(s), L.values(data)); let text = s;
  for (const e of [...p.edits].reverse()) text = text.slice(0, e.start) + e.replacement + text.slice(e.end);
  return { text, ...p };
};
test('case and meaningful separators', () => assert.deepEqual(L.parse('< Facility > <FACILITY> <Date of Entry> <Date-of-Entry> <Date_of_Entry>').fields, ['FACILITY', 'Date of Entry', 'Date-of-Entry', 'Date_of_Entry']));
test('conflict', () => assert.throws(() => L.values({ Facility: 'A', FACILITY: 'B' }), /Conflicting/));
test('prototype and punctuation fields', () => assert.equal(render('<__proto__> <A "quote">', { ['__proto__']: 'safe', 'A "quote"': 'Q' }).text, 'safe Q'));
test('unusual names reject safely', () => assert.throws(() => L.parse('<A\\B>'), /Unsupported/));
test('condition-only fields', () => assert.deepEqual(L.parse('{IF <Only> = "x"}ok{END IF}').fields, ['Only']));
test('text normalization only for comparisons', () => { assert.equal(evalCondition('<F> = “Plymouth”', { F: ' plyMouth ' }).value, true); assert.equal(render('<F>', { F: ' plyMouth ' }).text, ' plyMouth '); });
test('field-field equals is text', () => assert.equal(evalCondition('<A> = <B>', { A: '0180', B: '180' }).value, false));
test('field-field less is numeric', () => assert.equal(evalCondition('<A> < <B>', { A: '180', B: '181' }).value, true));
for (const [value, expected] of [['180', true], ['180.5', true], ['-25', false], ['0', false], ['1,200', true], ['12,345.67', true], ['12,34', null], ['1e3', null], ['two hundred', null], ['approximately 220', null]]) {
 test(`numeric ${value}`, () => assert.equal(evalCondition('<D> >= 180', { D: value }).value, expected));
}
test('huge precision', () => assert.equal(evalCondition('<D> > 9007199254740993.12345678901234567890', { D: '9007199254740993.12345678901234567891' }).value, true));
test('negative fraction', () => assert.equal(evalCondition('<D> < 0', { D: '-0.001' }).value, true));
test('oversized attorney number unknown', () => assert.equal(evalCondition('<D> >= 0', { D: '1'.repeat(1001) }).value, null));
for (const syntax of ['<D> >= 1e3', '<D> = 12,34', '<D>=180', '<D> = window.alert(1)', '<D> = "x', '<D> > "180"', '<D> = "x" garbage', '<D> = "bad\\q"']) test(`invalid ${syntax}`, () => assert.throws(() => L.expression(syntax)));
for (const [s, expected] of [
 ['<T> = "yes" OR <U> = "yes"', true], ['<F> = "yes" AND <U> = "yes"', false], ['<T> = "yes" AND <U> = "yes"', null], ['<F> = "yes" OR <U> = "yes"', null], ['NOT (<U> = "yes")', null], ['<T> = "yes" AND NOT (<F> = "yes")', true], ['<T> = "yes" OR <F> = "yes" AND <U> = "yes"', true]
]) test(`tri-state ${s}`, () => { const r = evalCondition(s, { T: 'yes', F: 'no' }); assert.equal(r.value, expected); if (expected !== null) assert.deepEqual(r.reasons, []); });
const unresolved = '{IF <D> >= 180}\n<N> long <Missing>\n{ELSE}\n<N> short\n{END IF}';
test('unknown complete original structure', () => { const r = render(unresolved, { N: 'Alex' }); assert.equal(r.text, unresolved.replaceAll('<N>', 'Alex')); assert.deepEqual(r.missing, ['Missing']); assert.equal(r.unresolved.length, 1); });
test('earlier unknown prevents later true', () => assert.equal(render('{IF <X> = "yes"}A{ELSE IF <Y> = "yes"}B{ELSE}C{END IF}', { Y: 'yes' }).unresolved.length, 1));
test('preserved inner not evaluated', () => { const r = render('{IF <X> = "yes"}{IF <Y> = "yes"}<N>{ELSE}<N>{END IF}{END IF}', { N: 'Alex', Y: 'yes' }); assert.equal(r.unresolved.length, 1); assert.equal(r.unresolved[0].unevaluated.length, 1); assert(r.text.includes('{IF <Y> = "yes"}Alex{ELSE}Alex')); });
test('outer resolved inner unknown', () => { const r = render('{IF <X> = "yes"}{IF <Y> = "yes"}A{ELSE}B{END IF}{ELSE}C{END IF}', { X: 'yes' }); assert.equal(r.text, '{IF <Y> = "yes"}A{ELSE}B{END IF}'); });
test('missing fields output only, once', () => { const r = render('{IF <X> = "yes"}<Excluded>{ELSE}<M> <m>{END IF}', { X: 'no' }); assert.equal(r.missing.length, 1); assert(!r.missing.includes('Excluded')); });
test('blank zero false', () => assert.equal(render('<A>|<B>|<C>', { A: 0, B: false, C: '  ' }).text, '0|false|<C>'));
test('inserted data never parsed', () => assert.equal(render('<N>', { N: '{IF <Bad> = "yes"}<Other>{END IF}' }).text, '{IF <Bad> = "yes"}<Other>{END IF}'));
test('literal square brackets and escapes', () => assert.equal(render('[sic] [T]he claim[s] \\<user@example.com> \\{IF <D> >= 180} C:\\file').text, '[sic] [T]he claim[s] <user@example.com> {IF <D> >= 180} C:\\file'));
test('escaped quotes, slashes and braces', () => assert.equal(evalCondition('<N> = "A\\"B\\\\C}"', { N: 'A"B\\C}' }).value, true));
for (const s of ['{ELSE}', '{IF <X> = "yes"}', '{IF <X> = "yes"}{ELSE}{ELSE}{END IF}', '{IF <X> = "yes"}{ELSE}{ELSE IF <Y> = "yes"}{END IF}', '{IF <X>\n= "yes"}A{END IF}', '{IF }A{END IF}']) test('structure ' + s, () => assert.throws(() => L.parse(s)));
for (const s of ['{ENDIF}', '{ELSEIF <X> = \"yes\"}']) test('malformed command ' + s, () => assert.throws(() => L.parse(s)));
test('unsafe XML characters reject while Unicode emoji works', () => { assert.throws(() => render('<N>', { N: String.fromCharCode(0) })); assert.equal(render('<N>', { N: '🙂' }).text, '🙂'); });
test('nonscalar data rejects', () => assert.throws(() => L.values({ N: { unexpected: true } })));
test('quotes in condition field names are ordinary name characters', () => assert.equal(render('{IF <A\"Field> = \"x\"}YES{END IF}', { 'A\"Field': 'x' }).text, 'YES'));
test('native unsafe numbers cannot silently select a branch', () => assert.equal(evalCondition('<N> > 0', { N: Number.MAX_SAFE_INTEGER + 1 }).value, null));
test('mixed case keywords', () => assert.equal(render('{if <X> = "yes"}A{else}B{end if}', { X: 'yes' }).text, 'A'));
console.log(`${checks} language regression checks passed.`);
