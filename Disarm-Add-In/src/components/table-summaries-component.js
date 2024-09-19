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
                  
          // add matched cases into table
          table.push([tagObject.title, tag, selectedText.slice(0, -1) + ".", tagObject.use]) 
        }
      }
      table = helpers.sortRedSummariesTable(table);
      drawTable(table)
    }
  )}
  
  async function drawTable(table){
    await Word.run(async (context) => {
  
      const secondParagraph = context.document.body.paragraphs.getLast()
      const insertedTable = secondParagraph.insertTable(table.length, 4, Word.InsertLocation.after, table);
      insertedTable.styleBuiltIn = Word.BuiltInStyleName.listTable3_Accent1;
  })
  
  };