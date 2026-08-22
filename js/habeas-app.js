/**
 * Habeas Petition Generator Application Logic
 */

/**
 * Derives comprehensive grammatical pronoun tags
 * @param {string} choice - "male", "female", "nonbinary", or "petitioner"
 * @returns {Object} Pronoun tag mappings
 */
function getPronounData(choice) {
  const map = {
    male: {
      he: 'he', He: 'He',
      him: 'him', Him: 'Him',
      his: 'his', His: 'His',
      hers: 'his', Hers: 'His',
      himself: 'himself', Himself: 'Himself',
      is_are: 'is', was_were: 'was', has_have: 'has'
    },
    female: {
      he: 'she', He: 'She',
      him: 'her', Him: 'Her',
      his: 'her', His: 'Her',
      hers: 'hers', Hers: 'Hers',
      himself: 'herself', Himself: 'Herself',
      is_are: 'is', was_were: 'was', has_have: 'has'
    },
    nonbinary: {
      he: 'they', He: 'They',
      him: 'them', Him: 'Them',
      his: 'their', His: 'Their',
      hers: 'theirs', Hers: 'Theirs',
      himself: 'themselves', Himself: 'Themselves',
      is_are: 'are', was_were: 'were', has_have: 'have'
    },
    petitioner: {
      he: 'Petitioner', He: 'Petitioner',
      him: 'Petitioner', Him: 'Petitioner',
      his: "Petitioner's", His: "Petitioner's",
      hers: "Petitioner's", Hers: "Petitioner's",
      himself: 'Petitioner', Himself: 'Petitioner',
      is_are: 'is', was_were: 'was', has_have: 'has'
    }
  };

  const p = map[choice] || map.male;

  return {
    "he": p.he, "He": p.He,
    "him": p.him, "Him": p.Him,
    "his": p.his, "His": p.His,
    "hers": p.hers, "Hers": p.Hers,
    "himself": p.himself, "Himself": p.Himself,
    "he/she": p.he, "He/She": p.He,
    "his/her": p.his, "His/Her": p.His,
    "him/her": p.him, "Him/Her": p.Him,
    "himself/herself": p.himself, "Himself/Herself": p.Himself,
    "is/are": p.is_are, "was/were": p.was_were, "has/have": p.has_have,
    "Pronoun Subject": p.He, "pronoun subject": p.he,
    "Pronoun Object": p.Him, "pronoun object": p.him,
    "Pronoun Possessive": p.His, "pronoun possessive": p.his,
    "Pronoun Reflexive": p.Himself, "pronoun reflexive": p.himself,
    "Pronouns": choice,
    "Gender": choice
  };
}

/**
 * Helper to collect current form values into structured template data
 */
function getFormData() {
  const petitionSelect = document.getElementById('petition_type');
  const petitionVal = petitionSelect ? petitionSelect.value : '';
  const petitionFullText = petitionSelect && petitionSelect.selectedIndex >= 0
    ? petitionSelect.options[petitionSelect.selectedIndex].text
    : '';

  // Extract name parts
  const firstName = document.getElementById('petitioner_first_name')?.value.trim() || '';
  const middleName = document.getElementById('petitioner_middle_name')?.value.trim() || '';
  const lastName = document.getElementById('petitioner_last_name')?.value.trim() || '';
  const suffix = document.getElementById('petitioner_suffix')?.value.trim() || '';

  const fullName = [firstName, middleName, lastName, suffix].filter(Boolean).join(' ');

  // Extract Pronouns (defaults to male)
  const selectedPronoun = document.querySelector('input[name="petitioner_pronouns"]:checked')?.value || 'male';
  const pronounData = getPronounData(selectedPronoun);

  const baseFormData = {
    "Jurisdiction": document.getElementById('jurisdiction')?.value || '',
    "Type of Habeas Petition": petitionVal,
    "Petition Type": petitionVal,
    "Type of Habeas Petition Full": petitionFullText,

    // Combined & Individual Name Variables
    "Petitioner Name": fullName,
    "Petitioner First Name": firstName,
    "Petitioner Middle Name": middleName,
    "Petitioner Last Name": lastName,
    "Petitioner Suffix": suffix,
    "Plaintiff Name": fullName,

    "Defendant Name": document.getElementById('defendant')?.value || '',
    "Respondent Name": document.getElementById('defendant')?.value || '',
    "Case Number": document.getElementById('case_number')?.value || ''
  };

  // Merge in pronoun aliases
  const formData = Object.assign({}, baseFormData, pronounData);

  return { formData, fullName };
}

/**
 * Handles action execution for preview or download
 */
async function runPleadingAction(shouldDownload) {
  const { formData, fullName } = getFormData();

  await DocxEngine.generate({
    templatePath: 'templates/petition.docx',
    data: formData,
    outputFilename: `Habeas Petition ${fullName || "Draft"}.docx`,
    selectElementForAliases: true,
    download: shouldDownload
  });
}

// Bind Button Click Events
document.addEventListener('DOMContentLoaded', () => {
  const renderBtn = document.getElementById('renderBtn');
  const downloadBtn = document.getElementById('downloadBtn');

  if (renderBtn) {
    renderBtn.addEventListener('click', () => runPleadingAction(false));
  }
  if (downloadBtn) {
    downloadBtn.addEventListener('click', () => runPleadingAction(true));
  }
});