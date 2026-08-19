function generateHabeasPetition() {
  const petitionSelect = document.getElementById('petition_type');
  const petitionVal = petitionSelect.value;
  const petitionFullText = petitionSelect.options[petitionSelect.selectedIndex].text;

  const formData = {
    "Jurisdiction": document.getElementById('jurisdiction').value,
    "Type of Habeas Petition": petitionVal,
    "Petition Type": petitionVal,
    "Type of Habeas Petition Full": petitionFullText,
    "Plaintiff Name": document.getElementById('plaintiff').value,
    "Defendant Name": document.getElementById('defendant').value,
    "Case Number": document.getElementById('case_number').value
  };

  DocxEngine.generate({
    templatePath: 'templates/petition.docx',
    data: formData,
    outputFilename: `Habeas Petition ${formData["Plaintiff Name"] || "Draft"}.docx`,
    selectElementForAliases: true
  });
}