(async () => {
  const results = [];
  const assert = (v,m='assertion failed') => { if(!v)throw Error(m); };
  const delay = () => new Promise(r => setTimeout(r, 25));
  const until = async fn => { for(let i=0;i<300;i++){if(fn())return;await delay();}throw Error('Timed out waiting for UI'); };
  const test = async (name, fn) => {document.getElementById('test-output').setAttribute('data-stage', name);try{await fn();results.push({name,passed:true});}catch(e){results.push({name,passed:false,error:e.stack});}};
  const frame = document.createElement('iframe');frame.src='/index.html';document.body.append(frame);
  await until(()=>frame.contentWindow?.eval('typeof DocxEngine') !== 'undefined' && frame.contentDocument.querySelector('#dynamic-fields-container input'));
  const app=frame.contentWindow, doc=frame.contentDocument; app.DocxEngine=app.eval('DocxEngine');
  await test('actual app controls and warning panel position',()=>{
    assert(!doc.querySelector('[name="petitioner_pronouns"]'));const buttons=doc.querySelectorAll('.action-block button');assert(buttons.length===4);
    assert(doc.getElementById('generation-status').previousElementSibling.classList.contains('action-block'));
    assert(doc.getElementById('generation-status').nextElementSibling.id==='docx-preview-container');
  });
  await test('actual Render button uses pinned docx-preview',async()=>{
    doc.getElementById('renderBtn').click();await until(()=>doc.getElementById('generation-status').textContent.includes('Preview ready.'));
    assert(doc.getElementById('docx-preview-container').querySelector('.docx-wrapper'));assert(doc.getElementById('generation-status').textContent.includes('blank fields'));
    assert(![...doc.querySelectorAll('#generation-status details')].filter(d => d.querySelector('summary').textContent.includes('blank field')).some(d=>d.open));
  });

  await test('warning counts, diagnostic reasons, expansion and collapse', () => {
    const unresolved = { location: 'Main document, paragraph 1', instruction: '{IF <Days> >= 180}', reasons: ['Invalid number in <Days>'], unevaluated: [] };
    app.GenerationUI.status('DOCX generated.', { warnings: { unresolved: [unresolved], missing: ['Date'] } });
    const panel = doc.getElementById('generation-status'); assert(panel.textContent.includes('Attorney review required — 1 condition could not be resolved'));
    assert(panel.textContent.includes('instructions and alternatives are preserved')); assert(panel.textContent.includes('Invalid number in <Days>')); assert(!panel.textContent.includes('Missing <Days>'));
    const details = panel.querySelectorAll('details'); assert(details.length === 2); assert(details[0].open); assert(!details[1].open); assert(details[1].textContent.includes('1 blank field'));
    details[0].querySelector('summary').click(); assert(!details[0].open);
    app.GenerationUI.status('New result', { warnings: { unresolved: [unresolved, unresolved], missing: ['Date', 'Facility'] } });
    assert(panel.textContent.includes('2 conditions could not be resolved')); assert(panel.querySelector('details').open); assert(panel.textContent.includes('2 blank fields'));
  });
  await test('missing-only and clear results distinguish document creation from filing', () => {
    const panel = doc.getElementById('generation-status');
    app.GenerationUI.status('DOCX generated.', { warnings: { unresolved: [], missing: ['Date'] } });
    assert(panel.textContent.includes('1 blank field remains')); assert(!panel.textContent.includes('could not be resolved')); assert(panel.textContent.includes('<Field Name> placeholders')); assert(panel.querySelectorAll('details').length === 1);
    app.GenerationUI.status('DOCX generated.', { warnings: { unresolved: [], missing: ['Date', 'Facility'] } }); assert(panel.textContent.includes('2 blank fields remain'));
    app.GenerationUI.status('DOCX generated.', { warnings: { unresolved: [], missing: [] } }); assert(panel.textContent.includes('no unresolved conditions or blank fields')); assert(panel.textContent.includes('does not establish readiness for filing')); assert(!panel.querySelector('details'));
  });
  let saved=[];app.saveAs=(blob,name)=>saved.push({blob,name});
  await test('actual Download button generates current values independently',async()=>{
    doc.getElementById('petitioner_first_name').value='Alex';doc.getElementById('petitioner_middle_name').value='';doc.getElementById('petitioner_last_name').value='Morgan';
    doc.getElementById('downloadBtn').click();await until(()=>saved.length===1);
    const z=new app.PizZip(await saved[0].blob.arrayBuffer());const text=app.DocxEngine.visible(app.DocxEngine.xml(z.file('word/document.xml').asText()));assert(text.includes('Alex Morgan'));assert(saved[0].name.includes('Alex Morgan'));
  });
  await test('actual Save and Load controls preserve dynamic spaces and status',async()=>{
    const input=doc.querySelector('[data-dynamic-tag]');input.value='  Exact entered value  ';const status=doc.getElementById('generation-status').textContent;
    doc.querySelector('.utility-actions button').click();await until(()=>saved.length===2);const file=saved[1].blob;input.value='changed';
    await app.FormStorage.load(file);assert(input.value==='  Exact entered value  ');assert(doc.getElementById('generation-status').textContent===status);
  });
  await test('actual preview of unresolved footnotes with pinned library',async()=>{
    const b=await(await fetch('/artifacts/07-footnotes-unresolved.docx')).arrayBuffer();
    const result=await app.GenerationUI.run(false,{buffer:b,collect:()=>({formData:{},fullName:''})});assert(result);
    assert(doc.getElementById('generation-status').textContent.includes('Preview ready.'));
    const text=doc.getElementById('docx-preview-container').textContent;assert(text.includes('{IF <Days> >= 180}'));assert(text.includes('INNER ELSE Alex'));assert(doc.querySelector('.docx-wrapper'));
  });
  const sf=document.createElement('iframe');sf.src='/drafting-studio.html';document.body.append(sf);
  await until(()=>sf.contentWindow?.inspectTemplate && sf.contentDocument.getElementById('petitioner_first_name'));
  const studio=sf.contentWindow,sd=sf.contentDocument;
  const fileReads = [], NativeReader = studio.FileReader;
  studio.FileReader = class extends NativeReader {
    constructor() {
      super(); fileReads.push(new Promise(resolve => this.addEventListener('loadend', resolve, { once: true })));
    }
  };
  const fileReadFinished = () => fileReads[fileReads.length - 1];

  const checkUploadEmpty = () => {
    assert(sd.getElementById('upload-template-name').textContent === 'Drag & Drop your .docx template here');
    assert(sd.getElementById('upload-template-detail').hidden);
    assert(sd.getElementById('upload-template-detail').textContent === '');
    assert(sd.getElementById('upload-template-hint').textContent === 'or click to browse from your computer');
  };
  const checkUploadLoaded = (name, fields) => {
    assert(sd.getElementById('upload-template-name').textContent === `Loaded: ${name}`);
    assert(!sd.getElementById('upload-template-detail').hidden);
    assert(sd.getElementById('upload-template-detail').textContent === `${fields} ${fields === 1 ? 'field' : 'fields'} detected · Template syntax checked`);
    assert(sd.getElementById('upload-template-hint').textContent === 'Click here or drop another Word document to replace it');
  };
  await test('upload area initial state and keyboard/click browse access', () => {
    checkUploadEmpty();
    const zone = sd.getElementById('dropzone'), input = sd.getElementById('fileInput');
    assert(zone.getAttribute('role') === 'button' && zone.tabIndex === 0);
    assert(zone.querySelector('[aria-live="polite"]'));
    const original = input.click; let clicks = 0; input.click = () => clicks++;
    try {
      zone.click();
      for (const key of ['Enter', ' ']) zone.dispatchEvent(new studio.KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }));
      assert(clicks === 3);
    } finally { input.click = original; }
  });
  await test('Studio cloned form excludes pronouns and keeps existing controls',()=>{
    assert(!sd.querySelector('[name="petitioner_pronouns"]'));assert(sd.getElementById('dropzone'));assert(sd.getElementById('fileInput'));assert(sd.getElementById('inspectSnippetBtn'));assert(sd.querySelectorAll('.action-block button').length===4);
  });
  await test('Studio paste button validates actual paragraph breaks',async()=>{
    const input=sd.getElementById('snippetInput');input.value='{IF <Only> = "yes"}\n<Name>\n{ELSE}\nOther\n{END IF}';input.dispatchEvent(new studio.Event('input'));sd.getElementById('inspectSnippetBtn').click();
    assert(sd.getElementById('results').textContent.includes('template syntax checked'));assert(sd.querySelector('[data-dynamic-tag="Only"]')); assert(sd.getElementById('active-template').textContent.includes('No active DOCX template.'));
    input.value='{IF <Only>\n= "yes"}X{END IF}';input.dispatchEvent(new studio.Event('input'));sd.getElementById('inspectSnippetBtn').click();assert(sd.getElementById('results').textContent.includes('Template error'));
  });
  const buffer=await(await fetch('/templates/petition.docx')).arrayBuffer();
  await test('Studio upload event and test download',async()=>{
    const file=new studio.File([buffer],'petition.docx',{type:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'}),data=new studio.DataTransfer();data.items.add(file);
    const upload=sd.getElementById('fileInput');upload.files=data.files;upload.dispatchEvent(new studio.Event('change'));await fileReadFinished();await until(()=>sd.getElementById('results').textContent.includes('petition.docx: template syntax'));
    checkUploadLoaded('petition.docx', studio.DocxEngine ? studio.DocxEngine.inspect(buffer).fields.length : studio.eval('DocxEngine').inspect(buffer).fields.length);
    let downloaded;studio.saveAs=(blob,name)=>downloaded={blob,name};sd.getElementById('downloadBtn').click();await until(()=>downloaded);assert(downloaded.blob.size>1000);
  });
  await test('Studio drag/drop event preserves upload workflow',async()=>{
    const data=new studio.DataTransfer();data.items.add(new studio.File([buffer],'dragged.docx'));
    sd.getElementById('dropzone').dispatchEvent(new studio.DragEvent('drop',{dataTransfer:data,bubbles:true,cancelable:true}));
    await fileReadFinished(); await until(()=>sd.getElementById('results').textContent.includes('dragged.docx: template syntax'));
    checkUploadLoaded('dragged.docx', studio.eval('DocxEngine').inspect(buffer).fields.length);
  });

  await test('Studio replacement updates filename, snippet inspection preserves DOCX identity', async () => {
    const b = studio.createDocxFromText('<Only> <Other>');
    studio.handleFile(new studio.File([b], 'replacement.docx')); await fileReadFinished(); await until(() => sd.getElementById('active-template').textContent.includes('replacement.docx'));
    const area = sd.getElementById('active-template'); assert(area.textContent.includes('2 fields detected')); assert(area.textContent.includes('Template syntax checked')); assert(area.textContent.includes('Word output not yet verified'));
    checkUploadLoaded('replacement.docx', 2);
    const input = sd.getElementById('snippetInput'); input.value='<Snippet Field>'; input.dispatchEvent(new studio.Event('input')); sd.getElementById('inspectSnippetBtn').click();
    assert(area.textContent.includes('replacement.docx')); assert(!area.textContent.includes('Pasted_Snippet.docx')); assert(area.textContent.includes('Sandbox source: pasted text snippet'));
    checkUploadLoaded('replacement.docx', 2);
    const download = await studio.runTestRender(true); assert(download);
  });
  await test('upload filename is literal text; snippet failure preserves DOCX identity', async () => {
    const name = '<img src=x onerror=alert(1)>.docx';
    const b = studio.createDocxFromText('<Only>');
    studio.handleFile(new studio.File([b], name)); await fileReadFinished();
    checkUploadLoaded(name, 1);
    assert(!sd.getElementById('dropzone').querySelector('img'));
    const input = sd.getElementById('snippetInput'); input.value = '{IF invalid}';
    input.dispatchEvent(new studio.Event('input')); sd.getElementById('inspectSnippetBtn').click();
    checkUploadLoaded(name, 1);
    assert(sd.getElementById('active-template').textContent.includes(name));
  });
  await test('Studio warnings share expansion and missing-field presentation', () => {
    studio.GenerationUI.status('DOCX generated.', { warnings: { unresolved: [{ location: 'Paragraph 1', instruction: '{IF <Days> >= 180}', reasons: [], unevaluated: [] }], missing: ['Name'] } });
    const panel = sd.getElementById('generation-status'); assert(panel.textContent.includes('1 condition could not be resolved')); assert(panel.querySelector('details').open); assert(panel.querySelectorAll('details').length === 2); assert(!panel.textContent.includes('Missing <Days>'));
  });
  await test('invalid DOCX and unsupported upload clear stale identity and sandbox', async () => {
    studio.handleFile(new studio.File(['not a Word package'], 'broken.docx'));
    await fileReadFinished(); await until(() => sd.getElementById('results').textContent.includes('Template error'));
    assert(sd.getElementById('active-template').textContent === 'No active DOCX template.'); assert(sd.getElementById('sandbox').style.display === 'none'); checkUploadEmpty();
    studio.handleFile(new studio.File([studio.createDocxFromText('<Only>')], 'valid-again.docx')); await fileReadFinished(); await until(() => sd.getElementById('active-template').textContent.includes('valid-again.docx'));
    studio.handleFile(new studio.File(['text'], 'invalid.txt')); assert(sd.getElementById('active-template').textContent === 'No active DOCX template.'); assert(!sd.getElementById('results').textContent.includes('valid-again.docx')); checkUploadEmpty();
  });

  await test('file read failure clears active template identification', async () => {
    const b = studio.createDocxFromText('<Only>'); studio.handleFile(new studio.File([b], 'before-read-failure.docx'));
    await fileReadFinished(); await until(() => sd.getElementById('active-template').textContent.includes('before-read-failure.docx'));
    const OriginalReader = studio.FileReader;
    studio.FileReader = class { readAsArrayBuffer() { this.onerror(); } };
    try { studio.handleFile(new studio.File([b], 'unreadable.docx')); checkUploadEmpty(); assert(sd.getElementById('active-template').textContent === 'No active DOCX template.'); assert(sd.getElementById('sandbox').style.display === 'none'); }
    finally { studio.FileReader = OriginalReader; }
  });

  await test('Preview and Download each expand fresh unresolved diagnostics', async () => {
    const buffer = studio.createDocxFromText('{IF <Days> >= 180}Long <Name>{ELSE}Short <Name>{END IF}');
    const options = { buffer, collect: () => ({ formData: {}, fullName: '' }) };
    assert(await studio.GenerationUI.run(true, options));
    const panel = sd.getElementById('generation-status'); assert(panel.querySelector('details').open);
    panel.querySelector('summary').click(); assert(!panel.querySelector('details').open);
    const render = studio.docx.renderAsync;
    studio.docx.renderAsync = async (_, target) => { target.textContent = 'Focused status preview'; };
    try { assert(await studio.GenerationUI.run(false, options)); assert(panel.querySelector('details').open); assert(panel.textContent.includes('Preview ready.')); }
    finally { studio.docx.renderAsync = render; }
  });
  await test('compilation and preview failures remain distinct', async () => {
    const options = { buffer: studio.createDocxFromText('<Missing>'), collect: () => ({ formData: {}, fullName: '' }) };
    const render = studio.docx.renderAsync;
    studio.docx.renderAsync = async () => { throw Error('test preview failure'); };
    try {
      assert(await studio.GenerationUI.run(false, options));
      const panel = sd.getElementById('generation-status'); assert(panel.textContent.includes('Word document created, but the preview could not be displayed:')); assert(panel.textContent.includes('1 blank field remains'));
      assert(await studio.GenerationUI.run(true, { ...options, buffer: studio.createDocxFromText('{IF invalid}') }) === null);
      assert(panel.textContent.includes('Generation failed')); assert(!panel.querySelector('details'));
    } finally { studio.docx.renderAsync = render; }
  });
  document.getElementById('test-output').textContent=JSON.stringify({results});
})().catch(e=>document.getElementById('test-output').textContent=JSON.stringify({fatal:e.stack}));
