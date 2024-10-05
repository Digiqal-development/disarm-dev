//extracts the text from the first found tag from the beginning and the endIndex, str is the entire document
export function extractTextFromBrackets(str, endIndex) {  
    //pattern to match evenrything that ends a sentence
    const pattern = /[.!?\n\t]/;
  
    //pattern to match ]) - in case that after a sentence there is a tag
    const patternBrackets = /]\)/;
  
    //taking string from the beginning to end index
    let s = str.substring(0, endIndex - 2)
  
    //taking the latest index of these patterns
    let index = (s.lastIndexOf(pattern.exec(s))> s.lastIndexOf(patternBrackets.exec(s))) ? s.lastIndexOf(pattern.exec(s)) + 1 : s.lastIndexOf(patternBrackets.exec(s)) + 1;
  
    return str.substring(index, endIndex);
  }

//sort red table according to kill chain
export function sortRedSummariesTable(table, jsonObject){

  const useOrder = Array.from(new Set(Object.values(jsonObject.ids).map(item => item.use)));
  const firstRow = table[0];
  const tableData  = table.slice(1);

  tableData.sort((a, b) => {
    //sort according to kill chain of tactic id
    const useComparison = useOrder.indexOf(a[3]) - useOrder.indexOf(b[3]);
    if (useComparison !== 0) return useComparison;

    //sort by technique id
    const idComparison = a[1].localeCompare(b[1]);
    if (idComparison !== 0) return idComparison;

    //sort by text
    return a[2].localeCompare(b[2]);
});

  return [firstRow, ...tableData];
}

//remove tags from the start of the tag text
export function removeExtraTagFromText(text){
  return text.indexOf(' ') === -1 ? text : text.substring(text.indexOf(' ') + 1)
}

export function searchTechniquesFromJson(phase, tactic, checkbox, textbox, reverseJsonTechniques){

  var resultTechniques = []
  var searchTechniquesArray = []

  const newJson = reverseJsonTechniques.ids;
  for (const key in newJson){
    
    const value = newJson[key] 
    var found = false   

    //check if phase and tactic match the row
    if ((phase == "" || phase == value.phase) && (tactic == "" || value.use.includes(tactic))){

      //check if the title includes the search term
      if (value.title.toLowerCase().includes(textbox.toLowerCase())){
        resultTechniques.push(key)
        found = true
      }

      //if technique is not found check the description (if checked)
      if (!found && checkbox && value.description.toLowerCase().includes(textbox.toLowerCase())){
        resultTechniques.push(key)
      }
    }
  }

  const tag_ids = reverseJsonTechniques.ids

  for (const tag in resultTechniques){
    let tag_object = tag_ids[resultTechniques[tag]]
    let use = tag_object.use.replace(/\s*\[.*?\]\s*/, '');

    searchTechniquesArray.push({ id: resultTechniques[tag], phase: tag_object.phase, use: use, title: tag_object.title, color: "white" });
  }

  return searchTechniquesArray;
}

export function getSelectedRowText(field) {
  const selectedRow = field.querySelector('tr.selected');
  return selectedRow ? selectedRow.textContent.trim() : '';
}

export function createTagTextInsert(checkboxes, jsonTags){
  const checkboxesArray = Array.from(checkboxes);
  const tagText = checkboxesArray
  .filter(checkbox => checkbox.checked) 
  .map(checkbox => {
    let tagText = removeExtraTagFromText(checkbox.nextSibling.textContent.trim())

    //for sub techniques add prefix
    if (checkbox.value.includes(".")) tagText = addPrefixForSubTechniques(tagText, checkbox.value, jsonTags.ids)
  
    return `${tagText} [${checkbox.value}]`
  })
  .join(', '); 

  const formattedTagText = tagText ? `(${tagText})` : '';
  return formattedTagText;
}

export function createTagTextSearch(table, jsonTags){

  var tagText = '(';
  Array.from(table.rows).forEach(row => {
    if (row.classList.contains('selected')) {
      var textContent = removeExtraTagFromText(row.cells[2].textContent);
      //for sub techniques add prefix
      if (row.id.includes(".")) textContent = addPrefixForSubTechniques(textContent, row.id, jsonTags.ids)
      tagText += `${textContent} [${row.id}], `;
    }
  });
  
  tagText = tagText.slice(0, -2) + ')';
  return tagText.length <= 2 ? '' : tagText;
}

function addPrefixForSubTechniques(tagText, tagId, jsonTags){
  var originalTag = tagId.split('.')[0];
  return jsonTags[originalTag].title + ": " + tagText;

}





