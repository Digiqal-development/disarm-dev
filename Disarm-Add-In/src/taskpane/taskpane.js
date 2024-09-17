/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */
import * as helpers from './../utils/helpers.js';
import * as helpersUI from '../utils/helpers-ui.js';
import { getTechniques, searchTechniques } from '../services/api-services.js';
import { changeRedTagColor, changeRedTagColorSaveBtn, getRedTagColor } from '../components/formatting-component.js';
import { PHASES} from '../constants/global-variables.js';

import { chooseTechniquesPopup, phaseChooseField, tacticChooseField, techniqueChooseCheckboxField, 
  searchTechniquesPopup, phaseSearchField, tacticSearchField, checkbox, textBox, 
  listSearchTechniquesPopup, alertSearchBox, table
} from '../constants/ui-elements.js';

var jsonTechniques = '';
var reverseJsonTechniques = '';
var searchTechniquesArray = []

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {


    //red tag
    document.getElementById("insert-red-tag").onclick = () => tryCatch(insertRedTag);
    document.getElementById("search-red-tag").onclick = () => tryCatch(searchRedTag);

    //blue tag
    document.getElementById("insert-blue-tag").onclick = () => tryCatch(test1);
    document.getElementById("search-blue-tag").onclick = () => tryCatch(test2);

    //summaries
    document.getElementById("insert-red-table").onclick = () => tryCatch(insertRedTable);

    //formatting
    document.getElementById("format-red-tag-color").onclick = () => tryCatch(changeRedTagColor);

    //save
    document.getElementById("save-btn").onclick = () => tryCatch(saveButton);
    document.getElementById("save-btn1").onclick = () => tryCatch(displaySearchTechniques);
    document.getElementById("save-btn2").onclick = () => tryCatch(changeRedTagColorSaveBtn);
    document.getElementById("save-btn3").onclick = () => tryCatch(saveButtonSearchTechniques);

    //close
    document.querySelectorAll('.close').forEach(element => element.onclick = () => tryCatch(closeButton));

  }
});

$(document).on('click','tr',function(e) { handleTableColorOnClick.call(this) }); 
$(document).on('click', 'th', function(e) {
  e.stopPropagation(); 
  sortTechniquesTable.call(this);
});


async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

async function closeButton(){
  await Word.run(async (context) => {

    helpersUI.closeAllFields();
    helpersUI.unblurBackground();
    searchTechniquesArray = [];
  })
}


//insert red tag 

async function insertRedTag() {
  await Word.run(async (context) => {

    helpersUI.blurBackground();
    showPhaseOptions(true);
    jsonTechniques = await getTechniques();
    helpersUI.clearInsertRedTagFields();
  });
}


function showPhaseOptions(insertTag){
  const uiField = insertTag ? phaseChooseField : phaseSearchField;
  uiField.innerHTML = ''

  PHASES.forEach((phase, index) => {
    const row = document.createElement('tr');
  
    row.addEventListener('click', () => {
      showTacticOptions(phase.toLowerCase(), insertTag);
    });
  
    row.innerHTML = `<td>${phase}</td>`;
    row.classList.add('list', 'option1');
    row.classList.add(index % 2 === 0 ? 'even-row' : 'odd-row');
  
    uiField.appendChild(row);
  });
}

function showTacticOptions(phase, insertTag){
  
  const uiField = insertTag ? tacticChooseField : tacticSearchField;
  techniqueChooseCheckboxField.innerHTML = ''
  uiField.innerHTML = '';

  let tacticsArray = jsonTechniques[phase]
  
  Object.keys(tacticsArray).forEach((tactic, index) => {
    const row = document.createElement('tr');

    if (insertTag) {
      row.addEventListener('click', function() {helpersUI.showTechniqueOptions(tacticsArray[tactic])});
    } 
    row.innerHTML = `<td>${tactic}</td>`;
    row.classList.add('list', 'option2');
    row.classList.add(index % 2 === 0 ? 'even-row' : 'odd-row');
    uiField.appendChild(row);

  })
}


async function saveButton(){
  await Word.run(async (context) => {

    const checkboxes = document.querySelectorAll('input[type="checkbox"]');
    const tagText = helpers.createTagText(checkboxes);
    helpersUI.clearInsertRedTagFields();
    phaseChooseField.innerHTML = '';
    chooseTechniquesPopup.style.display = 'none';

    //get sentence
    var sentenceInitialRange = context.document.getSelection().getTextRanges(['\n', '.', '?'], true);
    context.load(sentenceInitialRange);
    var sentenceRange =  await context.sync().then(() => { return sentenceInitialRange.items[0]})
    var sentenceWithoutDot = (sentenceRange.text).slice(0, -1); 
     
    //insert tag text to end of sentence
    sentenceRange.insertText(sentenceWithoutDot + ' ' + tagText + '.', Word.InsertLocation.replace);

    //search tag inside range to highlight
    var rangeToHighlight = sentenceRange.search(tagText);
    rangeToHighlight.load('items');
    await context.sync();
   
    //change highlight color of the tag
    rangeToHighlight.items[0].font.set({highlightColor: getRedTagColor()});

    helpersUI.unblurBackground();
    await context.sync(); 
  })
}


//search red tag

async function searchRedTag(){
  
  searchTechniquesPopup.style.display = 'block';
  helpersUI.blurBackground()
  showPhaseOptions()
  jsonTechniques = await getTechniques();
  tacticSearchField.innerHTML = '';
}

async function displaySearchTechniques(){

  if (textBox.value == "") {
    alertSearchBox.innerHTML = 'Please fill out this field!';
  }
  else {
    searchTechniquesPopup.style.display = 'none';
    if (reverseJsonTechniques == "") reverseJsonTechniques = await searchTechniques()
    searchTechniquesArray = helpers.searchTechniquesFromJson(helpers.getSelectedRowText(phaseSearchField), helpers.getSelectedRowText(tacticSearchField), checkbox.checked, textBox.value, reverseJsonTechniques)
    helpersUI.drawFoundTechniquesTable(searchTechniquesArray)

    listSearchTechniquesPopup.style.display = 'block';
    textBox.value = '';
    checkbox.checked = ''
    alertSearchBox.innerHTML = ''
  }
}


function sortTechniquesTable(){

  var column = $(this).text().trim();

  if (column == "Phase") searchTechniquesArray.sort((a,b) => a.phase.localeCompare(b.phase))
  else if (column == "Tactic") searchTechniquesArray.sort((a,b) => a.use.localeCompare(b.use))
  else if (column == "Technique") searchTechniquesArray.sort((a,b) => a.title.localeCompare(b.title))
  helpersUI.drawFoundTechniquesTable(searchTechniquesArray)
 
}

function handleTableColorOnClick(){
  if (this.classList.toString().includes("resultsTable")) handleColorOnClickResultsTable(this);
  else if (this.classList.toString().includes("list")) helpersUI.handleColorOnClickPhaseTactic(this);
  else return;
}

function handleColorOnClickResultsTable(this1){
  if (!this1.classList.toString().includes("selected")){
    this1.classList.add("selected")
    console.log(searchTechniquesArray)
    for (var i = 0; i < searchTechniquesArray.length; i++) {
      if (searchTechniquesArray[i].id === this1.id) searchTechniquesArray[i].color = "blue";
   }
  }
  else { 
    this1.classList.remove("selected")
    for (var i = 0; i < searchTechniquesArray.length; i++) {
      if (searchTechniquesArray[i].id === this1.id) searchTechniquesArray[i].color = "white";
   }
  }
}

async function saveButtonSearchTechniques(){
  await Word.run(async (context) => {

    let tagText = '('

    //iterate through the table
    for (let i = 0; i < table.rows.length; i++) {
      const row = table.rows[i];

      //if row is colored add its content to the text
      if (row.classList.toString().includes('selected')) {

        tagText += helpers.removeExtraTagFromText(row.cells[2].textContent) + " [" + row.id + "], ";    
           
      }
    }
    
    tagText = tagText.slice(0, -2); 
    tagText += ')'
    if (tagText.length < 3) tagText = '';

    listSearchTechniquesPopup.style.display = 'none';
    phaseSearchField.innerHTML = '';
    tacticSearchField.innerHTML = '';
    helpersUI.unblurBackground();



    //get sentence
    var sentenceInitialRange = context.document.getSelection().getTextRanges(['\n', '.', '?'], true);
    context.load(sentenceInitialRange);
    var sentenceRange =  await context.sync().then(() => { return sentenceInitialRange.items[0]})

    var sentenceWithoutDot = (sentenceRange.text).slice(0, -1); 
     
    //insert tag text to end of sentence
    sentenceRange.insertText(sentenceWithoutDot + ' ' + tagText + '.', Word.InsertLocation.replace);

    //search tag inside range to highlight
    var rangeToHighlight = sentenceRange.search(tagText);
    rangeToHighlight.load('items');
    await context.sync();
   
    //change highlight color of the tag
    rangeToHighlight.items[0].font.set({highlightColor: getRedTagColor()});

    tagText = ''

    searchTechniquesArray = []

    await context.sync(); 
  })
}

//summaries

async function insertRedTable() {

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
        table.push([tagObject.title, tag, selectedText, tagObject.use]) 
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



//test
async function test1(){
  await Word.run(async (context) => {
    const originalXml =
      "<Locations><Location>Juan</Location><Location>Hong</Location><Location>Sally</Location></Locations>";
    const customXmlPart = context.document.customXmlParts.add(originalXml);
    customXmlPart.load("id");
    const xmlBlob = customXmlPart.getXml();

    await context.sync();

    const readableXml = addLineBreaksToXML(xmlBlob.value);
    console.log("Added custom XML part:");
    console.log(readableXml);

    // Store the XML part's ID in a setting so the ID is available to other functions.
    const settings = context.document.settings;
    settings.add("ContosoReviewXmlPartId", customXmlPart.id);

    await context.sync();
  });
}

function addLineBreaksToXML( xmlBlob) {
  const replaceValue = new RegExp(">");
  return xmlBlob.replace(/></g, "> <");
}

async function test2(){
  // Queries a custom XML part for elements matching the search terms.
  await Word.run(async (context) => {
    const settings = context.document.settings;
    const xmlPartIDSetting = settings.getItemOrNullObject("ContosoReviewXmlPartId").load("value");

    await context.sync();

    if (xmlPartIDSetting.value) {
      const customXmlPart = context.document.customXmlParts.getItem(xmlPartIDSetting.value);
      const xpathToQueryFor = "/Locations/Location";
      const clientResult = customXmlPart.query(xpathToQueryFor, {
        contoso: "http://schemas.contoso.com/review/1.0"
      });

      
      var range = context.document.getSelection().paragraphs.getFirst();
      context.load(range)
      var sentenceRange =  await context.sync().then(() => { return range})
      var text = ""


      console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
      for (let i = 0; i < clientResult.value.length; i++) {
        text += clientResult.value[i] + '\n';
        
      }
      sentenceRange.insertText(text, Word.InsertLocation.start);

    } else {
      console.warn("Didn't find custom XML part to query");
    }
  });
}