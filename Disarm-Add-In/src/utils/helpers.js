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

