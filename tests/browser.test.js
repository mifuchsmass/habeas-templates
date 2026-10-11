/* Run in a real browser with the pinned PizZip and application scripts. */
(async () => {
  const results = [], artifacts = {};
  const assert = (value, message = 'assertion failed') => { if (!value) throw Error(message); };
  const equal = (a, b) => assert(JSON.stringify(a) === JSON.stringify(b), `${JSON.stringify(a)} !== ${JSON.stringify(b)}`);
  const test = async (name, fn) => { try { await fn(); results.push({ name, passed: true }); } catch (e) { results.push({ name, passed: false, error: e.stack }); } };
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const p = s => `<w:p><w:r><w:t xml:space="preserve">${esc(s)}</w:t></w:r></w:p>`;
  const numP = s => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${esc(s)}</w:t></w:r></w:p>`;
  const row = (a, b) => `<w:tr><w:tc><w:tcPr>${b === undefined ? '<w:gridSpan w:val="2"/>' : ''}</w:tcPr>${p(a)}</w:tc>${b === undefined ? '' : `<w:tc><w:tcPr/>${p(b)}</w:tc>`}</w:tr>`;
  const table = rows => `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>${rows}</w:tbl>`;
  const ns = DocxEngine.ns, ref = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
  function fixture(body, footnotes = null) {
    const zip = new PizZip(Uint8Array.from(atob(template64), c => c.charCodeAt(0)));
    zip.file('word/document.xml', `<w:document xmlns:w="${ns}"><w:body>${body}<w:sectPr/></w:body></w:document>`);
    if (footnotes !== null) {
      zip.file('word/styles.xml', zip.file('word/styles.xml').asText().replace('</w:styles>', '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style><w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style></w:styles>'));
      zip.file('word/footnotes.xml', `<w:footnotes xmlns:w="${ns}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${footnotes}</w:footnotes>`);
      zip.file('[Content_Types].xml', zip.file('[Content_Types].xml').asText().replace('</Types>', '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'));
      zip.file('word/_rels/document.xml.rels', zip.file('word/_rels/document.xml.rels').asText().replace('</Relationships>', '<Relationship Id="rIdTestFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>'));
    }
    return zip.generate({ type: 'arraybuffer' });
  }
  const foot = content => `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>&lt;Na</w:t></w:r><w:r><w:t>me&gt;</w:t></w:r></w:p>${content}</w:footnote>`;
  const getText = (result, file = 'word/document.xml') => DocxEngine.visible(DocxEngine.xml(result.zip.file(file).asText(), file)).replace(/\u000b/g, '\n');
  const save = (name, result) => { artifacts[name] = result.zip.generate({ type: 'base64' }); };
  const throws = fn => { let threw = false; try { fn(); } catch (e) { threw = true; } assert(threw, 'expected an explicit error'); };
  const conditional = ['{IF <Days> >= 180}', 'LONG <Name> <Missing>', '{ELSE IF <Days> >= 90}', 'MEDIUM <Name>', '{ELSE}', 'SHORT <Name>', '{END IF}'].map(p).join('');
  await test('resolved / else-if / unknown real DOCX', () => {
    const buf = fixture(conditional);
    let out = DocxEngine.compile(buf, { Days: '180', Name: ' Alex ' }); equal(getText(out), 'LONG  Alex  <Missing>'); equal(out.warnings.missing, ['Missing']); save('01-resolved', out);
    out = DocxEngine.compile(buf, { Days: '100', Name: 'Alex' }); equal(getText(out), 'MEDIUM Alex'); equal(out.warnings.missing, []);
    out = DocxEngine.compile(buf, { Name: 'Alex' }); assert(getText(out).includes('{IF <Days> >= 180}LONG Alex <Missing>{ELSE IF <Days> >= 90}MEDIUM Alex{ELSE}SHORT Alex{END IF}')); equal(out.warnings.unresolved.length, 1); save('02-unresolved', out);
  });
  await test('nested uncertainty and selected outer', () => {
    const source = '{IF <Outer> = "yes"}{IF <Inner> = "yes"}<Name> YES{ELSE}<Name> NO{END IF}{ELSE}OUTSIDE{END IF}';
    const buf = fixture(p(source)); let out = DocxEngine.compile(buf, { Name: 'Alex', Inner: 'yes' });
    equal(getText(out), source.replaceAll('<Name>', 'Alex')); equal(out.warnings.unresolved.length, 1); equal(out.warnings.unresolved[0].unevaluated.length, 1);
    out = DocxEngine.compile(buf, { Name: 'Alex', Outer: 'yes' }); equal(getText(out), '{IF <Inner> = "yes"}Alex YES{ELSE}Alex NO{END IF}'); save('03-nested', out);
  });
  await test('styles / split runs / numbering / table / package relationships', () => {
    const body = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>&lt;Na</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>me&gt;</w:t></w:r></w:p>` + p('{IF <Show> = "yes"}') + numP('Included <Name>') + table(row('Cell <Name>', 'Other')) + p('{END IF}') + numP('After');
    const buf = fixture(body), input = new PizZip(buf), out = DocxEngine.compile(buf, { Name: 'Alex', Show: 'yes' });
    const xml = out.zip.file('word/document.xml').asText(); assert(xml.includes('<w:b')); assert(xml.includes('<w:i')); assert(xml.includes('w:numPr')); assert(xml.includes('w:tbl')); assert(!getText(out).includes('{IF'));
    for (const n of Object.keys(input.files)) if (n !== 'word/document.xml' && !input.files[n].dir) equal(out.zip.file(n).asText(), input.file(n).asText());
    save('04-numbering-table', out);
    const excluded = DocxEngine.compile(buf, { Show: 'no', Name: 'Alex' }); assert(!excluded.zip.file('word/document.xml').asText().includes('w:tbl')); equal(getText(excluded), 'AlexAfter'); equal(excluded.warnings.missing, []);
  });
  await test('conditional spanning rows removes whole rows', () => {
    const buf = fixture(table(row('{IF <Show> = "yes"}') + row('Included <Name>', 'Second cell') + row('{END IF}') + row('After', 'Still here')));
    const out = DocxEngine.compile(buf, { Show: 'no' }); equal(getText(out), 'AfterStill here'); equal(DocxEngine.descendants(DocxEngine.xml(out.zip.file('word/document.xml').asText()), 'tr').length, 1);
    save('05-table-row-excluded', out);
  });
  await test('multi-paragraph inline boundaries retain adjacent prose', () => {
    const buf = fixture(p('Before {IF <Show> = "yes"}chosen') + p('middle') + p('end{ELSE}fallback{END IF} After'));
    const out = DocxEngine.compile(buf, { Show: 'no' }); equal(getText(out), 'Before fallback After');
  });
  const noteCondition = ['{IF <Days> >= 180}', 'LONG <Name>', '{IF <Inner> = "yes"}', 'INNER <Name>', '{ELSE}', 'INNER ELSE <Name>', '{END IF}', '{ELSE}', 'SHORT <Name> <Note Missing>', '{END IF}'].map(p).join('');
  const noteBuf = fixture(`<w:p><w:r><w:t>Body &lt;Name&gt;</w:t></w:r>${ref}</w:p>`, foot(noteCondition));
  await test('footnotes resolved nested / UNKNOWN / styles / numbering markers', () => {
    const out = DocxEngine.compile(noteBuf, { Name: 'Alex', Days: '180', Inner: 'yes' }); equal(getText(out, 'word/footnotes.xml'), 'AlexLONG AlexINNER Alex');
    const xml = out.zip.file('word/footnotes.xml').asText(); for (const marker of ['FootnoteText', 'FootnoteReference', 'footnoteRef', 'w:b', 'w:i', 'separator', 'continuationSeparator']) assert(xml.includes(marker), marker);
    const original = new PizZip(noteBuf); equal(out.zip.file('word/_rels/document.xml.rels').asText(), original.file('word/_rels/document.xml.rels').asText()); save('06-footnotes-resolved', out);
    const unknown = DocxEngine.compile(noteBuf, { Name: 'Alex', Inner: 'yes' }); assert(getText(unknown, 'word/footnotes.xml').includes('{IF <Inner> = "yes"}INNER Alex{ELSE}INNER ELSE Alex{END IF}')); equal(unknown.warnings.unresolved.length, 1); equal(unknown.warnings.unresolved[0].unevaluated.length, 1); equal(unknown.warnings.missing, ['Note Missing']); save('07-footnotes-unresolved', unknown);
  });
  await test('excluded reference removes unused footnote and its warnings', () => {
    const buf = fixture(p('{IF <Show> = "yes"}') + `<w:p><w:r><w:t>Referenced body</w:t></w:r>${ref}</w:p>` + p('{END IF}'), foot(noteCondition));
    const out = DocxEngine.compile(buf, { Show: 'no' }); assert(!out.zip.file('word/document.xml').asText().includes('footnoteReference')); assert(!out.zip.file('word/footnotes.xml').asText().includes('w:id="1"')); equal(out.warnings.missing, []); equal(out.warnings.unresolved, []); save('08-excluded-reference', out);
  });
  await test('inline excluded reference does not survive', () => {
    const buf = fixture(`<w:p><w:r><w:t>Before {IF &lt;Show&gt; = "yes"}included</w:t></w:r>${ref}<w:r><w:t>{END IF} After</w:t></w:r></w:p>`, foot(p('note')));
    const out = DocxEngine.compile(buf, { Show: 'no' }); equal(getText(out), 'Before  After'); assert(!out.zip.file('word/document.xml').asText().includes('footnoteReference'));
  });
  await test('footnote marker survives removal of its instruction paragraph', () => {
    const note = `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t>{IF &lt;Show&gt; = "yes"}</w:t></w:r></w:p>${p('YES')}${p('{ELSE}')}${p('NO')}${p('{END IF}')}</w:footnote>`;
    const out = DocxEngine.compile(fixture(`<w:p>${ref}</w:p>`, note), { Show: 'no' }); equal(getText(out, 'word/footnotes.xml'), 'NO'); assert(out.zip.file('word/footnotes.xml').asText().includes('footnoteRef')); save('09-footnote-marker', out);
  });
  await test('unreferenced invalid footnote syntax still fails validation', () => throws(() => DocxEngine.compile(fixture(p('No reference'), foot(p('{IF <X> = "yes"}'))), {})));
  await test('footnotes cannot balance each other', () => throws(() => DocxEngine.inspect(fixture(`<w:p>${ref}</w:p>`, `<w:footnote w:id="1">${p('{IF <X> = "yes"}')}</w:footnote><w:footnote w:id="2">${p('{END IF}')}</w:footnote>`))));
  await test('main and footnote cannot balance each other', () => throws(() => DocxEngine.inspect(fixture(p('{IF <X> = "yes"}'), foot(p('{END IF}'))))));
  await test('instruction split across paragraphs or table cells fails', () => {
    throws(() => DocxEngine.inspect(fixture(p('{IF <X>') + p('= "yes"}A{END IF}'))));
    throws(() => DocxEngine.inspect(fixture(table(row('{IF <X>', '= "yes"}A{END IF}')))));
  });
  await test('instruction split across runs and Word line break succeeds', () => {
    const body = '<w:p><w:r><w:t>{IF &lt;Da</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ys&gt;</w:t><w:br/><w:t> &gt;= 180}YES{ELSE}NO{END IF}</w:t></w:r></w:p>';
    equal(getText(DocxEngine.compile(fixture(body), { Days: '180' })), 'YES');
  });
  await test('unsafe inactive branch fails rather than executes', () => throws(() => DocxEngine.compile(fixture(p('{IF <X> = "yes"}yes{ELSE}{IF window.alert(1)}unsafe{END IF}{END IF}')), { X: 'yes' })));
  await test('safe pass-through complex structure, unsupported automation rejects', () => {
    const literal = '<w:p><w:r><w:drawing><w:txbxContent>' + p('Literal box') + '</w:txbxContent></w:drawing></w:r></w:p>';
    const out = DocxEngine.compile(fixture(literal + p('<Name>')), { Name: 'Alex' }); assert(getText(out).includes('Literal box'));
    throws(() => DocxEngine.inspect(fixture(literal.replace('Literal box', '&lt;Name&gt;'))));
    throws(() => DocxEngine.inspect(fixture(p('{IF <X> = "yes"}') + '<w:p><w:bookmarkStart w:id="1" w:name="B"/><w:r><w:t>X</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' + p('{END IF}'))));
  });
  await test('literal escapes and field punctuation in XML', () => {
    const out = DocxEngine.compile(fixture(p('[sic] [T]he claim[s] \\<mail@example.com> \\{IF <Days> >= 180} <A "quote"> <__proto__> <Name>')), { 'A "quote"': 'Q', ['__proto__']: 'P', Name: '<Other>{IF <X> = "yes"}' });
    equal(getText(out), '[sic] [T]he claim[s] <mail@example.com> {IF <Days> >= 180} Q P <Other>{IF <X> = "yes"}'); equal(out.warnings.missing, []);
  });
  await test('template field discovery across all footnote stories', () => {
    const fields = DocxEngine.extractCustomFields(noteBuf); assert(fields.includes('Note Missing')); assert(fields.includes('Inner')); assert(fields.includes('Days'));
  });
  await test('broken reference is an explicit generation failure', () => throws(() => DocxEngine.compile(fixture(`<w:p>${ref}</w:p>`), {})));
  await test('migrated sample generates without old syntax or grammar tags', () => {
    const input = Uint8Array.from(atob(template64), c => c.charCodeAt(0));
    const out = DocxEngine.compile(input, { Jurisdiction: 'Massachusetts', 'Petitioner Name': 'Alex Morgan', 'Case Number': 'TEST-001', 'Defendant Name': 'Respondent', 'Type of Habeas Petition': 'Non-GO', 'Country of Citizenship': 'Example', 'Continuous Residence Date': 'May 2020', 'Date of Entry': 'May 2020', 'Entry Location': 'Example', 'Date of Arrest': 'May 2026', 'Detention Facility': 'Plymouth' });
    equal(out.warnings.unresolved, []); equal(out.warnings.missing, []); const text = getText(out); assert(!text.includes('[IF')); assert(!text.includes('[He/She]')); assert(text.includes("Petitioner is being held")); save('10-sample-petition', out);
    const missing = DocxEngine.compile(input, { Jurisdiction: 'Massachusetts', 'Type of Habeas Petition': 'Non-GO' }); assert(getText(missing).includes('<Petitioner Name>')); assert(!getText(missing).includes('Draft')); save('11-sample-missing-fields', missing);
  });

  await test('case, overlap and meaningful field separators in real DOCX', () => {
    const out = DocxEngine.compile(fixture(p('< Facility >|<FACILITY>|<Detention Facility>|<Date of Entry>|<Date_of_Entry>|<Date-of-Entry>')), { facility: ' F ', 'Detention Facility': 'D', 'Date of Entry': 'S', Date_of_Entry: 'U', 'Date-of-Entry': 'H' });
    equal(getText(out), ' F | F |D|S|U|H');
    throws(() => DocxEngine.compile(fixture(p('<Facility>')), { Facility: 'A', FACILITY: 'B' }));
  });
  await test('invalid numeric data UNKNOWN and exact large decimal choice in Word', () => {
    const buf = fixture(p('{IF <N> > 9007199254740993.12345678901234567890}HIGH{ELSE}LOW{END IF}'));
    equal(getText(DocxEngine.compile(buf, { N: '9007199254740993.12345678901234567891' })), 'HIGH');
    const out = DocxEngine.compile(buf, { N: '12,34' }); assert(getText(out).startsWith('{IF')); assert(out.warnings.unresolved[0].reasons[0].includes('Invalid number')); save('12-invalid-number', out);
  });
  await test('definite AND/OR results suppress irrelevant missing-condition warnings', () => {
    const buf = fixture(p('{IF <X> = "yes" OR <Missing> = "yes"}YES{ELSE}NO{END IF}'));
    const out = DocxEngine.compile(buf, { X: 'yes' }); equal(getText(out), 'YES'); equal(out.warnings.unresolved, []); equal(out.warnings.missing, []);
  });
  await test('multiple footnotes keep IDs and only retained references create warnings', () => {
    const reference = id => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
    const body = `<w:p><w:r><w:t>First</w:t></w:r>${reference(1)}</w:p>` + p('{IF <Show> = "yes"}') + `<w:p>${reference(2)}</w:p>` + p('{END IF}') + `<w:p>${reference(3)}</w:p>`;
    const notes = [1,2,3].map(id => `<w:footnote w:id="${id}" w:type="normal"><w:p><w:r><w:footnoteRef/><w:t>Note ${id} &lt;F${id}&gt;</w:t></w:r></w:p></w:footnote>`).join('');
    const out = DocxEngine.compile(fixture(body, notes), { Show: 'no', F1: 'One', F3: 'Three' });
    equal(out.warnings.missing, []); equal(getText(out, 'word/footnotes.xml'), 'Note 1 OneNote 3 Three');
    assert(!out.zip.file('word/footnotes.xml').asText().includes('w:id="2"')); save('13-multiple-footnotes', out);
  });
  await test('marker and instruction in same footnote run retain marker only', () => {
    const notes = `<w:footnote w:id="1"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/><w:t>{IF &lt;X&gt; = "yes"}</w:t></w:r></w:p>${p('YES')}${p('{END IF}')}</w:footnote>`;
    const out = DocxEngine.compile(fixture(`<w:p>${ref}</w:p>`, notes), { X: 'no' }); equal(getText(out, 'word/footnotes.xml'), ''); assert(out.zip.file('word/footnotes.xml').asText().includes('footnoteRef')); save('14-empty-footnote', out);
  });
  await test('non-text reference on selected instruction paragraph remains', () => {
    const body = `<w:p>${ref}<w:r><w:t>{IF &lt;X&gt; = "yes"}</w:t></w:r></w:p>${p('YES')}${p('{END IF}')}`;
    const out = DocxEngine.compile(fixture(body, foot(p('Footnote'))), { X: 'yes', Name: 'Alex' }); assert(out.zip.file('word/document.xml').asText().includes('footnoteReference'));
  });
  await test('multiline and tab values are Word breaks and tabs without reinterpretation', () => {
    const out = DocxEngine.compile(fixture(p('<Name>')), { Name: 'Line1\nLine2\t<Literal>' });
    // Actual newline/tab values, rather than two-character escaped sequences.
    const actual = DocxEngine.compile(fixture(p('<Name>')), { Name: 'Line1' + String.fromCharCode(10) + 'Line2' + String.fromCharCode(9) + '<Literal>' });
    equal(getText(actual), 'Line1' + String.fromCharCode(10) + 'Line2' + String.fromCharCode(9) + '<Literal>'); const xml = actual.zip.file('word/document.xml').asText(); assert(xml.includes('w:br')); assert(xml.includes('w:tab')); equal(actual.warnings.missing, []);
  });
  await test('unsupported text in drawing or header produces explicit error', () => {
    throws(() => DocxEngine.inspect(fixture('<w:p><w:r><w:drawing xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:t>&lt;Name&gt;</a:t></w:drawing></w:r></w:p>')));
    const zip = new PizZip(fixture(p('body'))); zip.file('word/header1.xml', `<w:hdr xmlns:w="${ns}">${p('<Name>')}</w:hdr>`);
    throws(() => DocxEngine.inspect(zip.generate({ type: 'arraybuffer' })));
  });
  await test('blank fields remain exact original placeholders once, zero and false insert', () => {
    const out = DocxEngine.compile(fixture(p('< Missing >|<MISSING>|<Zero>|<False>')), { Missing: '   ', Zero: 0, False: false }); equal(getText(out), '< Missing >|<MISSING>|0|false'); equal(out.warnings.missing.length, 1);
  });


  await test('bulleted list and numbered paragraph formatting remain', () => {
    const bullet = numP('Bullet <Name>').replace('w:numId w:val="1"', 'w:numId w:val="99"');
    const source = new PizZip(fixture(p('{IF <X> = "yes"}') + bullet + p('{END IF}') + numP('Numbered <Name>')));
    const originalNumbering = source.file('word/numbering.xml').asText();
    source.file('word/numbering.xml', originalNumbering
      .replace(/<w:num\b/, '<w:abstractNum w:abstractNumId="99"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num')
      .replace('</w:numbering>', '<w:num w:numId="99"><w:abstractNumId w:val="99"/></w:num></w:numbering>'));
    const out = DocxEngine.compile(source.generate({ type: 'arraybuffer' }), { X: 'yes', Name: 'Alex' }); equal(getText(out), 'Bullet AlexNumbered Alex'); assert(out.zip.file('word/document.xml').asText().includes('w:numId w:val="99"')); equal(out.zip.file('word/numbering.xml').asText(), source.file('word/numbering.xml').asText());
    const numbering = DocxEngine.xml(out.zip.file('word/numbering.xml').asText());
    const definitions = DocxEngine.descendants(numbering, 'abstractNum');
    const instances = DocxEngine.descendants(numbering, 'num');
    const children = [...numbering.documentElement.children];
    const firstInstance = children.findIndex(n => n.namespaceURI === ns && n.localName === 'num');
    assert(firstInstance >= 0);
    assert(children.every((n, index) => n.localName !== 'abstractNum' || index < firstInstance), 'Abstract definitions must precede numbering instances');
    equal(definitions.map(n => n.getAttributeNS(ns, 'abstractNumId')), ['0', '1', '2', '99']);
    equal(instances.map(n => n.getAttributeNS(ns, 'numId')), ['1', '2', '3', '99']);
    const definitionIds = definitions.map(n => n.getAttributeNS(ns, 'abstractNumId'));
    const instanceIds = instances.map(n => n.getAttributeNS(ns, 'numId'));
    equal(new Set(definitionIds).size, definitionIds.length);
    equal(new Set(instanceIds).size, instanceIds.length);
    for (const instance of instances) {
      const links = DocxEngine.descendants(instance, 'abstractNumId');
      equal(links.length, 1); assert(definitionIds.includes(links[0].getAttributeNS(ns, 'val')), 'Broken abstract numbering reference');
    }
    const original = DocxEngine.xml(originalNumbering);
    for (const local of ['abstractNum', 'num']) {
      equal(DocxEngine.descendants(numbering, local).filter(n => n.getAttributeNS(ns, local === 'num' ? 'numId' : 'abstractNumId') !== '99').map(n => new XMLSerializer().serializeToString(n)),
        DocxEngine.descendants(original, local).map(n => new XMLSerializer().serializeToString(n)));
    }
    const bulletDefinition = definitions.find(n => n.getAttributeNS(ns, 'abstractNumId') === '99');
    const bulletLevel = DocxEngine.descendants(bulletDefinition, 'lvl').find(n => n.getAttributeNS(ns, 'ilvl') === '0');
    equal(DocxEngine.descendants(bulletLevel, 'numFmt')[0].getAttributeNS(ns, 'val'), 'bullet');
    // Construct the expected marker independently of the HTML source's decoding.
    equal(DocxEngine.descendants(bulletLevel, 'lvlText')[0].getAttributeNS(ns, 'val'), String.fromCodePoint(0x2022));
    const decimalReference = DocxEngine.descendants(numbering, 'num').find(n => n.getAttributeNS(ns, 'numId') === '1');
    const decimalId = DocxEngine.descendants(decimalReference, 'abstractNumId')[0].getAttributeNS(ns, 'val');
    const decimalDefinition = definitions.find(n => n.getAttributeNS(ns, 'abstractNumId') === decimalId);
    const decimalLevel = DocxEngine.descendants(decimalDefinition, 'lvl').find(n => n.getAttributeNS(ns, 'ilvl') === '0');
    equal(DocxEngine.descendants(decimalLevel, 'numFmt')[0].getAttributeNS(ns, 'val'), 'decimal');
    equal(DocxEngine.descendants(decimalLevel, 'lvlText')[0].getAttributeNS(ns, 'val'), '%1)');
    const paragraphs = DocxEngine.descendants(DocxEngine.xml(out.zip.file('word/document.xml').asText()), 'p');
    equal(paragraphs.map(n => DocxEngine.descendants(n, 'numId')[0]?.getAttributeNS(ns, 'val')), ['99', '1']);
    equal(paragraphs.map(n => DocxEngine.descendants(n, 'ilvl')[0]?.getAttributeNS(ns, 'val')), ['0', '0']);
    equal(paragraphs.map(n => DocxEngine.visible(n)), ['Bullet Alex', 'Numbered Alex']);
    for (const [index, paragraph] of paragraphs.entries()) {
      const id = DocxEngine.descendants(paragraph, 'numId')[0].getAttributeNS(ns, 'val');
      const instance = instances.find(n => n.getAttributeNS(ns, 'numId') === id); assert(instance, 'Broken paragraph numbering reference');
      const abstractId = DocxEngine.descendants(instance, 'abstractNumId')[0].getAttributeNS(ns, 'val');
      const definition = definitions.find(n => n.getAttributeNS(ns, 'abstractNumId') === abstractId); assert(definition);
      const level = DocxEngine.descendants(definition, 'lvl').find(n => n.getAttributeNS(ns, 'ilvl') === '0'); assert(level);
      equal(DocxEngine.descendants(level, 'numFmt')[0].getAttributeNS(ns, 'val'), index === 0 ? 'bullet' : 'decimal');
      equal(DocxEngine.descendants(level, 'lvlText')[0].getAttributeNS(ns, 'val'), index === 0 ? String.fromCodePoint(0x2022) : '%1)');
      equal(DocxEngine.descendants(instance, 'lvlOverride').length, 0);
    }
    equal(document.characterSet.toUpperCase(), 'UTF-8');
    save('15-bullets-and-numbering-final-corrected', out);
  });
  await test('Word field-code automation cannot silently pass', () => {
    throws(() => DocxEngine.inspect(fixture('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>{IF &lt;X&gt; = "yes"}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')));
  });
  await test('partial merged-row table control is rejected', () => {
    const body = table(row('{IF <X> = "yes"}') + row('Merged').replace('<w:tcPr>', '<w:tcPr><w:vMerge w:val="restart"/>') + row('{END IF}') + row('Outside merged continuation').replace('<w:tcPr>', '<w:tcPr><w:vMerge/>'));
    throws(() => DocxEngine.inspect(fixture(body)));
  });

  // Real controls for form, status, saved-data, and generation concurrency tests.
  const controls = document.createElement('div'); controls.innerHTML = '<div id="generation-status"></div><div id="docx-preview-container"></div><select id="jurisdiction"><option>Massachusetts</option><option>New Hampshire</option></select><select id="detention_facility"></select><select id="petition_type"><option value="Non-GO">Non-GO full text</option></select><input id="petitioner_first_name"><input id="petitioner_middle_name"><input id="petitioner_last_name"><input id="petitioner_suffix"><input id="defendant"><input id="case_number"><input data-dynamic-tag="A &quot;quote&quot;"><input data-dynamic-tag="Date of Entry"><input data-dynamic-tag="Date_of_Entry">'; document.body.append(controls);
  await test('name construction preserves parts, missing name blank, dropdown code explicit', () => {
    document.getElementById('petitioner_first_name').value = ' Alex '; document.getElementById('petitioner_last_name').value = ' Morgan '; document.getElementById('petitioner_suffix').value = ' , Jr. ';
    const r = getFormData(); equal(r.fullName, 'Alex Morgan, Jr.'); equal(r.formData['Petitioner First Name'], ' Alex '); equal(r.formData['Petitioner Suffix'], ' , Jr. '); equal(r.formData['Type of Habeas Petition'], 'Non-GO'); equal(r.formData['Type of Habeas Petition Full'], 'Non-GO full text');
    for (const id of ['petitioner_first_name', 'petitioner_last_name', 'petitioner_suffix']) document.getElementById(id).value = '';
    equal(getFormData().fullName, '');
  });
  await test('new saved-data matching is safe, case-insensitive and distinguishes separators', () => {
    const inputs = controls.querySelectorAll('[data-dynamic-tag]'); inputs[0].value = '  exact "quoted" text  '; inputs[1].value = 'space'; inputs[2].value = 'underscore';
    const saved = FormStorage.snapshot(); assert(!Object.keys(saved.values).some(k => k.includes('pronoun')));
    const converted = Object.fromEntries(Object.entries(saved.values).map(([k,v]) => [k.toUpperCase(), v]));
    inputs[0].value = inputs[1].value = inputs[2].value = ''; FormStorage.restore({ format: 'habeas-form-v3', values: converted });
    equal(inputs[0].value, '  exact "quoted" text  '); equal(inputs[1].value, 'space'); equal(inputs[2].value, 'underscore');
    throws(() => FormStorage.restore({ format: 'habeas-form-v3', values: { F: 'a', f: 'b' } }));
  });

  await test('saved primary IDs cannot collide with distinct underscore field names', () => {
    const primary = document.getElementById('petitioner_first_name'); primary.setAttribute('data-form-field', 'Petitioner First Name'); primary.value = 'Primary';
    const dynamic = document.createElement('input'); dynamic.setAttribute('data-dynamic-tag', 'petitioner_first_name'); dynamic.value = 'Separate underscore field'; controls.append(dynamic);
    const saved = FormStorage.snapshot(); equal(saved.values['Petitioner First Name'], 'Primary'); equal(saved.values.petitioner_first_name, 'Separate underscore field');
    primary.value = dynamic.value = ''; FormStorage.restore(saved); equal(primary.value, 'Primary'); equal(dynamic.value, 'Separate underscore field'); dynamic.remove(); primary.value = '';
  });

  await test('independent Download and warnings reset, Save/Load leave warning status alone', async () => {
    const oldDownload = DocxEngine.download; let downloads = 0; DocxEngine.download = () => downloads++;
    const buf = fixture(p('<Missing>')); await GenerationUI.run(true, { buffer: buf, collect: () => ({ formData: {}, fullName: '' }) }); equal(downloads, 1); assert(GenerationUI.panel().textContent.includes('1 blank field'));
    const before = GenerationUI.panel().textContent; const saved = FormStorage.snapshot(); FormStorage.save(); await FormStorage.load({ text: async () => JSON.stringify(saved) }); equal(GenerationUI.panel().textContent, before);
    await GenerationUI.run(true, { buffer: buf, collect: () => ({ formData: { Missing: 'Known' }, fullName: '' }) }); assert(!GenerationUI.panel().textContent.includes('1 blank field')); assert(!GenerationUI.panel().querySelector('details'));
    DocxEngine.download = oldDownload;
  });
  await test('overlapping Preview and Download preserve latest warnings and preview', async () => {
    const oldGenerate = DocxEngine.generate, oldDownload = DocxEngine.download; DocxEngine.download = () => {};
    let finishOld, finishNew, releasePreview;
    const oldResult = DocxEngine.compile(fixture(p('<Old Missing>')), {}), newResult = DocxEngine.compile(fixture(p('New result')), {});
    let calls = 0; DocxEngine.generate = () => new Promise(resolve => { if (++calls === 1) finishOld = resolve; else finishNew = resolve; });
    window.docx = { renderAsync: async (_, container) => { await new Promise(r => releasePreview = r); container.textContent = 'OLD PREVIEW'; } };
    document.getElementById('docx-preview-container').textContent = 'Previous completed preview';
    const old = GenerationUI.run(false, { collect: () => ({ formData: {}, fullName: '' }) }); finishOld(oldResult); await Promise.resolve(); await Promise.resolve();
    const newer = GenerationUI.run(true, { collect: () => ({ formData: {}, fullName: '' }) }); assert(GenerationUI.panel().textContent.includes('Preparing Word document for download…')); finishNew(newResult); await newer;
    releasePreview(); await old; assert(!GenerationUI.panel().textContent.includes('Old Missing')); equal(document.getElementById('docx-preview-container').textContent, 'Previous completed preview');
    DocxEngine.generate = oldGenerate; DocxEngine.download = oldDownload;
  });
  await test('late generation completion cannot replace newer status', async () => {
    const oldGenerate = DocxEngine.generate, oldDownload = DocxEngine.download; DocxEngine.download = () => {}; let a, b, count = 0;
    DocxEngine.generate = () => new Promise(r => { if (++count === 1) a = r; else b = r; });
    const opts = { collect: () => ({ formData: {}, fullName: '' }) }, first = GenerationUI.run(true, opts), second = GenerationUI.run(true, opts);
    b(DocxEngine.compile(fixture(p('New')), {})); await second; const status = GenerationUI.panel().textContent;
    a(DocxEngine.compile(fixture(p('<Old>')), {})); await first; equal(GenerationUI.panel().textContent, status);
    DocxEngine.generate = oldGenerate; DocxEngine.download = oldDownload;
  });
  await test('preview failure differs from generation failure', async () => {
    window.docx = { renderAsync: async () => { throw Error('simulated preview problem'); } };
    await GenerationUI.run(false, { buffer: fixture(p('<M>')), collect: () => ({ formData: {}, fullName: '' }) });
    assert(GenerationUI.panel().textContent.includes('Word document created, but the preview could not be displayed:')); assert(GenerationUI.panel().textContent.includes('1 blank field'));
    await GenerationUI.run(true, { buffer: fixture(p('{IF invalid}')), collect: () => ({ formData: {}, fullName: '' }) });
    assert(GenerationUI.panel().textContent.includes('Generation failed')); assert(!GenerationUI.panel().textContent.includes('1 blank field'));
  });
  // Studio logic uses the same Word inspector; UI upload/paste event behavior checked separately.
  const studio = document.createElement('div'); studio.innerHTML = '<div id="results"></div><div id="sandbox"><div id="sandbox-dynamic-section"><div id="sandbox-dynamic-container"></div></div></div>'; document.body.append(studio);
  await test('Studio pasted paragraphs enforce boundaries and share discovery/errors', () => {
    const good = createDocxFromText('{IF <Only> = "x"}\n<Name>\n{ELSE}\nOther\n{END IF}'); inspectTemplate(good, 'Test'); assert(document.getElementById('results').textContent.includes('template syntax checked')); assert(!document.getElementById('results').textContent.includes('Ready for Production'));
    equal([...document.querySelectorAll('#sandbox-dynamic-container input')].map(n => n.getAttribute('data-dynamic-tag')), ['Only', 'Name']);
    inspectTemplate(createDocxFromText('{IF <Only>\n= "x"}text{END IF}'), 'Bad'); assert(document.getElementById('results').textContent.includes('Template error'));
  });
  document.getElementById('test-output').textContent = JSON.stringify({ results, artifacts });
})().catch(e => { document.getElementById('test-output').textContent = JSON.stringify({ fatal: e.stack }); });
