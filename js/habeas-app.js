/** Form collection shared with the Studio; original values remain unchanged. */
const TEMPLATE_PATH = 'templates/petition.docx';
const PRIMARY_FIELDS = ['Jurisdiction', 'Detention Facility', 'Facility', 'Type of Habeas Petition', 'Petition Type', 'Type of Habeas Petition Full', 'Petitioner Name', 'Petitioner First Name', 'Petitioner Middle Name', 'Petitioner Last Name', 'Petitioner Suffix', 'Plaintiff Name', 'Defendant Name', 'Respondent Name', 'Case Number', 'Docket Number'];
function getFormData() {
  const value = id => document.getElementById(id)?.value ?? '';
  const first = value('petitioner_first_name'), middle = value('petitioner_middle_name'), last = value('petitioner_last_name'), suffix = value('petitioner_suffix');
  // Only the derived full name trims parts and normalizes its suffix separator.
  const base = [first, middle, last].map(s => s.trim()).filter(Boolean).join(' ');
  const cleanSuffix = suffix.replace(/^[,\s]+/, '').trim();
  const fullName = [base, cleanSuffix].filter(Boolean).join(base && cleanSuffix ? ', ' : '');
  const petition = document.getElementById('petition_type');
  const formData = Object.create(null);
  Object.assign(formData, {
    Jurisdiction: value('jurisdiction'), 'Detention Facility': value('detention_facility'), Facility: value('detention_facility'),
    'Type of Habeas Petition': value('petition_type'), 'Petition Type': value('petition_type'),
    'Type of Habeas Petition Full': petition?.selectedOptions[0]?.textContent ?? '',
    'Petitioner Name': fullName, 'Plaintiff Name': fullName,
    'Petitioner First Name': first, 'Petitioner Middle Name': middle, 'Petitioner Last Name': last, 'Petitioner Suffix': suffix,
    'Defendant Name': value('defendant'), 'Respondent Name': value('defendant'), 'Case Number': value('case_number'), 'Docket Number': value('case_number')
  });
  for (const input of document.querySelectorAll('input[data-dynamic-tag]')) {
    const tag = input.getAttribute('data-dynamic-tag');
    const existing = Object.keys(formData).find(k => TemplateLogic.key(k) === TemplateLogic.key(tag));
    if (existing && formData[existing] !== input.value) throw Error(`Conflicting values for <${tag}>.`);
    formData[tag] = input.value;
  }
  TemplateLogic.values(formData);
  return { formData, fullName };
}
function renderDynamicFields(fields, container, section, sample = false) {
  const primary = new Set(PRIMARY_FIELDS.map(TemplateLogic.key));
  container.replaceChildren();
  for (const tag of fields.filter(f => !primary.has(TemplateLogic.key(f)))) {
    const wrapper = document.createElement('div'), label = document.createElement('label'), input = document.createElement('input');
    wrapper.className = 'field'; input.type = 'text'; input.setAttribute('data-dynamic-tag', tag);
    input.id = `dynamic-field-${container.children.length}`; label.htmlFor = input.id; label.textContent = `${tag}:`;
    input.placeholder = `Enter ${tag}…`; if (sample) input.value = '';
    wrapper.append(label, input); container.append(wrapper);
  }
  section.style.display = container.children.length ? 'block' : 'none';
}
async function initializeDynamicFields() {
  try {
    const fields = DocxEngine.extractCustomFields(await DocxEngine.loadTemplate(TEMPLATE_PATH));
    renderDynamicFields(fields, document.getElementById('dynamic-fields-container'), document.getElementById('dynamic-fields-section'));
  } catch (e) { GenerationUI.status(`Template could not be loaded: ${e.message}`); }
}
async function runPleadingAction(download) {
  return GenerationUI.run(download, { templatePath: TEMPLATE_PATH, collect: getFormData });
}
document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('dynamic-fields-container')) return; // Studio binds its own controls.
  initializeDynamicFields();
  const jurisdiction = document.getElementById('jurisdiction'), facility = document.getElementById('detention_facility');
  DocxEngine.updateFacilityDropdown(facility, jurisdiction.value);
  jurisdiction.addEventListener('change', () => DocxEngine.updateFacilityDropdown(facility, jurisdiction.value));
  document.getElementById('renderBtn').addEventListener('click', () => runPleadingAction(false));
  document.getElementById('downloadBtn').addEventListener('click', () => runPleadingAction(true));
});
