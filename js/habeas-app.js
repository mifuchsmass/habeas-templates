/**
 * Habeas Petition Generator Application Logic
 */

const TEMPLATE_PATH = 'templates/petition.docx';

// State-to-Facility Directory
const STATE_FACILITIES = {
  "Massachusetts": [
    { value: "Plymouth", text: "Plymouth (Plymouth County Correctional Facility)" },
    { value: "Burlington", text: "Burlington (Burlington ICE Office)" },
    { value: "Boston", text: "Boston (Boston Field Office)" }
  ],
  "New Hampshire": [
    { value: "Strafford", text: "Strafford (Strafford County House of Corrections)" },
    { value: "Berlin", text: "Berlin (FCI Berlin)" }
  ],
  "Rhode Island": [
    { value: "Wyatt", text: "Wyatt (Donald W. Wyatt Detention Facility)" }
  ],
  "Vermont": [
    { value: "Northwest", text: "Northwest (Northwest State Correctional Facility)" }
  ]
};

function updateFacilityDropdown(selectEl, state) {
  if (!selectEl) return;
  const options = STATE_FACILITIES[state] || [{ value: "N/A", text: "N/A (Standard / Not Applicable)" }];
  selectEl.innerHTML = options.map(opt => `<option value="${opt.value}">${opt.text}</option>`).join('');
}

/**
 * Derives comprehensive grammatical pronoun and party tags
 */
function getPronounData(choice) {
  const map = {
    male: {
      he: 'he', He: 'He', HE: 'HE',
      him: 'him', Him: 'Him', HIM: 'HIM',
      his: 'his', His: 'His', HIS: 'HIS',
      hers: 'his', Hers: 'His', HERS: 'HIS',
      himself: 'himself', Himself: 'Himself', HIMSELF: 'HIMSELF',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'is', was_were: 'was', has_have: 'has',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    },
    female: {
      he: 'she', He: 'She', HE: 'SHE',
      him: 'her', Him: 'Her', HIM: 'HER',
      his: 'her', His: 'Her', HIS: 'HER',
      hers: 'hers', Hers: 'Hers', HERS: 'HERS',
      himself: 'herself', Himself: 'Herself', HIMSELF: 'HERSELF',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'is', was_were: 'was', has_have: 'has',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    },
    nonbinary: {
      he: 'they', He: 'They', HE: 'THEY',
      him: 'them', Him: 'Them', HIM: 'THEM',
      his: 'their', His: 'Their', HIS: 'THEIR',
      hers: 'theirs', Hers: 'Theirs', HERS: 'THEIRS',
      himself: 'themselves', Himself: 'Themselves', HIMSELF: 'THEMSELVES',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'are', was_were: 'were', has_have: 'have',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    },
    petitioner: {
      he: 'Petitioner', He: 'Petitioner', HE: 'PETITIONER',
      him: 'Petitioner', Him: 'Petitioner', HIM: 'PETITIONER',
      his: "Petitioner's", His: "Petitioner's", HIS: "PETITIONER'S",
      hers: "Petitioner's", Hers: "Petitioner's", HERS: "PETITIONER'S",
      himself: 'Petitioner', Himself: 'Petitioner', HIMSELF: 'PETITIONER',
      petitioner: 'Petitioner', petitioner_lower: 'petitioner', PETITIONER: 'PETITIONER',
      petitioner_possessive: "Petitioner's", PETITIONER_POSSESSIVE: "PETITIONER'S",
      individual: 'an individual',
      is_are: 'is', was_were: 'was', has_have: 'has',
      s_suffix: 's',
      brings_bring: 'brings', seeks_seek: 'seeks', contends_contend: 'contends'
    }
  };

  const p = map[choice] || map.male;

  return {
    "he": p.he, "He": p.He, "HE": p.HE,
    "him": p.him, "Him": p.Him, "HIM": p.HIM,
    "his": p.his, "His": p.His, "HIS": p.HIS,
    "hers": p.hers, "Hers": p.Hers, "HERS": p.HERS,
    "himself": p.himself, "Himself": p.Himself, "HIMSELF": p.HIMSELF,
    "he/she": p.he, "He/She": p.He, "HE/SHE": p.HE,
    "his/her": p.his, "His/Her": p.His, "HIS/HER": p.HIS,
    "him/her": p.him, "Him/Her": p.Him, "HIM/HER": p.HIM,
    "himself/herself": p.himself, "Himself/Herself": p.Himself, "HIMSELF/HERSELF": p.HIMSELF,
    "his/hers": p.hers,

    "Petitioner": p.petitioner,
    "petitioner": p.petitioner_lower,
    "PETITIONER": p.PETITIONER,
    "Petitioner's": p.petitioner_possessive,
    "petitioner's": p.petitioner_possessive.toLowerCase(),
    "PETITIONER'S": p.PETITIONER_POSSESSIVE,
    "Petitioner/Petitioners": p.petitioner,
    "petitioner/petitioners": p.petitioner_lower,
    "PETITIONER/PETITIONERS": p.PETITIONER,
    "Petitioner's/Petitioners'": p.petitioner_possessive,
    "petitioner's/petitioners'": p.petitioner_possessive.toLowerCase(),

    "Pronoun Subject": p.He, "pronoun subject": p.he,
    "Pronoun Object": p.Him, "pronoun object": p.him,
    "Pronoun Possessive": p.His, "pronoun possessive": p.his,
    "Pronoun Reflexive": p.Himself, "pronoun reflexive": p.himself,

    "individual/individuals": p.individual,
    "an individual/individuals": p.individual,

    "is/are": p.is_are, "was/were": p.was_were, "has/have": p.has_have,
    "brings/bring": p.brings_bring, "seeks/seek": p.seeks_seek, "contends/contend": p.contends_contend,
    "s": p.s_suffix,

    "Pronouns": choice,
    "Gender": choice
  };
}

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
  const pronounData = getPronounData(selectedPronoun);

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
    updateFacilityDropdown(facilitySelect, jurisdictionSelect.value);
    jurisdictionSelect.addEventListener('change', (e) => {
      updateFacilityDropdown(facilitySelect, e.target.value);
    });
  }

  const renderBtn = document.getElementById('renderBtn');
  const downloadBtn = document.getElementById('downloadBtn');

  if (renderBtn) renderBtn.addEventListener('click', () => runPleadingAction(false));
  if (downloadBtn) downloadBtn.addEventListener('click', () => runPleadingAction(true));
});