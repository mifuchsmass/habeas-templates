/**
 * Habeas Petition Generator Application Logic
 */

const TEMPLATE_PATH = 'templates/petition.docx';

/**
 * Automatically loads the default template and renders discovered dynamic fields
 */
async function initializeDynamicFields() {
  const section = document.getElementById('dynamic-fields-section');
  const container = document.getElementById('dynamic-fields-container');
  if (!section || !container) return;

  try {
    const buffer = await DocxEngine.loadTemplate(TEMPLATE_PATH);
    const discovered = DocxEngine.extractCustomFields(buffer);

    if (discovered.length === 0) {
      section.style.display = 'none';
      container.innerHTML = '';
      return;
    }

    container.innerHTML = '';
    discovered.forEach(tag => {
      const fieldWrapper = document.createElement('div');
      fieldWrapper.className = 'field';

      const label = document.createElement('label');
      label.textContent = tag + ':';

      const input = document.createElement('input');
      input.type = 'text';
      input.setAttribute('data-dynamic-tag', tag);

      if (tag.toLowerCase().includes('date')) {
        input.placeholder = "e.g., May 15, 2026 or on or about May 2024...";
      } else {
        input.placeholder = `Enter ${tag}...`;
      }

      fieldWrapper.appendChild(label);
      fieldWrapper.appendChild(input);
      container.appendChild(fieldWrapper);
    });

    section.style.display = 'block';
  } catch (err) {
    console.warn("Could not auto-discover dynamic fields for default template:", err.message);
  }
}

/**
 * Collects form values into structured template data (Primary + Dynamic)
 */
function getFormData() {
  const petitionSelect = document.getElementById('petition_type');
  const petitionVal = petitionSelect ? petitionSelect.value : '';
  const petitionFullText = petitionSelect && petitionSelect.selectedIndex >= 0
    ? petitionSelect.options[petitionSelect.selectedIndex].text
    : '';

  const firstName = document.getElementById('petitioner_first_name')?.value.trim() || '';
  const middleName = document.getElementById('petitioner_middle_name')?.value.trim() || '';
  const lastName = document.getElementById('petitioner_last_name')?.value.trim() || '';
  const rawSuffix = document.getElementById('petitioner_suffix')?.value.trim() || '';

  const cleanSuffix = rawSuffix.replace(/^[,\s]+/, '').trim();
  const baseName = [firstName, middleName, lastName].filter(Boolean).join(' ');
  const fullName = cleanSuffix 
    ? (baseName ? `${baseName}, ${cleanSuffix}` : cleanSuffix)
    : baseName;

  const selectedPronoun = document.querySelector('input[name="petitioner_pronouns"]:checked')?.value || 'male';
  const pronounData = DocxEngine.getPronounData(selectedPronoun);

  const facilityVal = document.getElementById('detention_facility')?.value || '';

  const baseFormData = {
    "Jurisdiction": document.getElementById('jurisdiction')?.value || '',
    "Detention Facility": facilityVal,
    "Facility": facilityVal,
    "Type of Habeas Petition": petitionVal,
    "Petition Type": petitionVal,
    "Type of Habeas Petition Full": petitionFullText,

    "Petitioner Name": fullName,
    "Petitioner First Name": firstName,
    "Petitioner Middle Name": middleName,
    "Petitioner Last Name": lastName,
    "Petitioner Suffix": cleanSuffix,
    "Plaintiff Name": fullName,

    "Defendant Name": document.getElementById('defendant')?.value || '',
    "Respondent Name": document.getElementById('defendant')?.value || '',
    "Case Number": document.getElementById('case_number')?.value || ''
  };

  const dynamicData = {};
  document.querySelectorAll('input[data-dynamic-tag]').forEach(input => {
    const tag = input.getAttribute('data-dynamic-tag');
    dynamicData[tag] = input.value.trim();
  });

  const formData = Object.assign({}, baseFormData, pronounData, dynamicData);
  return { formData, fullName };
}

/**
 * Handles action execution for preview or download
 */
async function runPleadingAction(shouldDownload) {
  const { formData, fullName } = getFormData();

  await DocxEngine.generate({
    templatePath: TEMPLATE_PATH,
    data: formData,
    outputFilename: `Habeas Petition ${fullName || "Draft"}.docx`,
    selectElementForAliases: true,
    download: shouldDownload
  });
}

// Bind Action Buttons and Jurisdiction Change
document.addEventListener('DOMContentLoaded', () => {
  initializeDynamicFields();

  const jurisdictionSelect = document.getElementById('jurisdiction');
  const facilitySelect = document.getElementById('detention_facility');

  if (jurisdictionSelect && facilitySelect) {
    DocxEngine.updateFacilityDropdown(facilitySelect, jurisdictionSelect.value);
    jurisdictionSelect.addEventListener('change', (e) => {
      DocxEngine.updateFacilityDropdown(facilitySelect, e.target.value);
    });
  }

  const renderBtn = document.getElementById('renderBtn');
  const downloadBtn = document.getElementById('downloadBtn');

  if (renderBtn) renderBtn.addEventListener('click', () => runPleadingAction(false));
  if (downloadBtn) downloadBtn.addEventListener('click', () => runPleadingAction(true));
});