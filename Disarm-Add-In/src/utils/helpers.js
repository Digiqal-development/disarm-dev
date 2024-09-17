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
export function sortRedSummariesTable(table){
  // Extract the first row and rest of table
  const firstRow = table[0];
  const restOfData  = table.slice(1);
  restOfData.sort((a, b) => {
    if (extractNumberFromBrackets(a[3]) !== extractNumberFromBrackets(b[3])) {
        return extractNumberFromBrackets(a[3]) - extractNumberFromBrackets(b[3]); 
    }
    return a[1].localeCompare(b[1]); 
  });

  //return concated table
  return [firstRow, ...restOfData];
}

function extractNumberFromBrackets(element){ 
  const match = element.match(/\[TA(\d+)\]/);
  return match ? parseInt(match[1], 10) : Infinity; 
};

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

export function createTagText(checkboxes){
  const checkboxesArray = Array.from(checkboxes);
  const tagText = checkboxesArray
  .filter(checkbox => checkbox.checked) 
  .map(checkbox => 
    `${removeExtraTagFromText(checkbox.nextSibling.textContent.trim())} [${checkbox.value}]`
  ) 
  .join(', '); 

  const formattedTagText = tagText ? `(${tagText})` : '';
  return formattedTagText;
}





