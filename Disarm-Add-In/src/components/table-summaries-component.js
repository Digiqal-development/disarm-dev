import { searchTechniques } from '../services/api-services.js';
import * as helpers from './../utils/helpers.js';

var reverseJsonTechniques = '';

export async function insertRedTable() {

    await Word.run(async (context) => {
      
      //get search json only once
      if (reverseJsonTechniques == '') reverseJsonTechniques = await searchTechniques();
      
      const tag_ids = reverseJsonTechniques.ids
      
      // get text of the entire document
      var wholeDocument = context.document.body
      wholeDocument.load("text")
      await context.sync();
      var text = wholeDocument.text
      
      //match cases using regular expressions
      const smallBracketRegex = /\(([^()]*)\)/g;
      const middleBracketRegex = /\[(.*?)\]/g;
  
      const foundSmallBrackets = text.matchAll(smallBracketRegex);
  
      let foundMiddleBrackets = []
      let table = [["Technique title", "ID", "Text", "Use"]]
  
      for (const matches of foundSmallBrackets){
  
        foundMiddleBrackets = matches[0].match(middleBracketRegex);
        const commasNumber = (matches[0].match(/,/g) || []).length;
        let selectedText = helpers.extractTextFromBrackets(text, matches.index)
  
        if (foundMiddleBrackets.length != commasNumber + 1 || foundMiddleBrackets.length == 0) continue;
  
        for (let j = 0; j < foundMiddleBrackets.length; j++){
          
          let tag = foundMiddleBrackets[j].slice(1, -1)
          let tagObject = tag_ids[tag]
  
          selectedText = selectedText.replace(/^\s*/, '');

          var tagTitle = tagObject.title;
          if (tag.includes(".")) tagTitle = helpers.addPrefixForSubTechniques(tagTitle, tag, reverseJsonTechniques.ids);
                  
          // add matched cases into table
          table.push([tagTitle, tag, selectedText.slice(0, -1) + ".", tagObject.use]) 
        }
      }

      table = helpers.sortRedSummariesTable(table, reverseJsonTechniques);
      const { processedData, headerIndexRows } = extractHeaderRows(table)
      drawTable(processedData, headerIndexRows)
    }
  )}
  
  async function drawTable(processedData, headerIndexRows){
    await Word.run(async (context) => {
  
      var body = context.document.body;

    // Insert a section break before the new page
    body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);

    // Load the new section to modify its properties
    var sections = context.document.sections;
    context.load(sections);


      const secondParagraph = context.document.body.paragraphs.getLast()
      const insertedTable = secondParagraph.insertTable(processedData.length, 3, Word.InsertLocation.after, processedData);
      
      //centered alignement for table
      insertedTable.horizontalAlignment = "Centered";

      //left alignement for third column
      processedData.forEach((_, index) => {
        insertedTable.getCell(index, 2).horizontalAlignment = "Left";
      });
   
      //set column widths
      insertedTable.getCell(0, 0).columnWidth = 150
      insertedTable.getCell(0, 1).columnWidth = 60;
      insertedTable.getCell(0, 2).columnWidth = 500;
 
      //da ide na novu str/landscap
      
      //format first row
      const firstRow = insertedTable.rows.getFirst();
      firstRow.load('font');
      context.sync();
      firstRow.set({
        font: {
            color: "#FFFFFF",
            bold: true
        },
        shadingColor: "#64649b",
        horizontalAlignment: "Centered"
      });

      //format header rows
      headerIndexRows.forEach(rowIndex => {
        const cell = insertedTable.getCell(rowIndex, 0).parentRow;
        cell.load('font');
        context.sync();
        cell.set({
          horizontalAlignment: "Centered",
          shadingColor: "#9bcbfb",
          font: {
              bold: true
          }
        });
      cell.merge();

      context.sync();

      });
    });
  };

function extractHeaderRows(table) {
  const result = [["Technique Title", "ID", "Use"]];
  const headerIndexRows = [];
  let previousUse = '';

  table.forEach((row, index) => {
  
    const [technique, id, text, use] = row;  

    if (use !== previousUse) {
        result.push([use, '', '']); 
        headerIndexRows.push(result.length - 1); 
        previousUse = use; 
    }

    result.push([technique, id, text]);
    });

    return {  processedData: result, headerIndexRows };
}
