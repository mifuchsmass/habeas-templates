function generateHabeasPetition() {
  const petitionSelect = document.getElementById('petition_type');
  const petitionVal = petitionSelect.value;
  const petitionFullText = petitionSelect.options[petitionSelect.selectedIndex].text;

  // Extract individual name parts
  const firstName = document.getElementById('petitioner_first_name').value.trim();
  const middleName = document.getElementById('petitioner_middle_name').value.trim();
  const lastName = document.getElementById('petitioner_last_name').value.trim();
  const suffix = document.getElementById('petitioner_suffix').value.trim();

  // Combine into full "Petitioner Name"
  const fullName = [firstName, middleName, lastName, suffix].filter(Boolean).join(' ');

  const formData = {
    "Jurisdiction": document.getElementById('jurisdiction').value,
    "Type of Habeas Petition": petitionVal,
    "Petition Type": petitionVal,
    "Type of Habeas Petition Full": petitionFullText,

    // Combined & Individual Name Variables
    "Petitioner Name": fullName,
    "Petitioner First Name": firstName,
    "Petitioner Middle Name": middleName,
    "Petitioner Last Name": lastName,
    "Petitioner Suffix": suffix,

    // Backward compatibility alias for any existing templates using Plaintiff Name
    "Plaintiff Name": fullName,

    "Defendant Name": document.getElementById('defendant').value,
    "Respondent Name": document.getElementById('defendant').value,
    "Case Number": document.getElementById('case_number').value
  };

  DocxEngine.generate({
    templatePath: 'templates/petition.docx',
    data: formData,
    outputFilename: `Habeas Petition ${fullName || "Draft"}.docx`,
    selectElementForAliases: true,
    showPreview: document.getElementById('show_preview')?.checked
  });
}