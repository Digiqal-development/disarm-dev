/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */
import * as helpers from './../utils/helpers.js';
import * as helpersUI from '../utils/helpers-ui.js';
import { getTechniques, searchTechniques } from '../services/api-services.js';
import { changeRedTagColor, changeRedTagColorSaveBtn, getRedTagColor } from '../components/formatting-component.js';
import { showPhaseOptions, saveButton } from '../components/red-tag-component.js';

import { 
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
    document.getElementById("save-btn").onclick = () => tryCatch(finishInsertRedTag);
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
    await showPhaseOptions(true);
    helpersUI.clearInsertRedTagFields();
  });
}

async function finishInsertRedTag(){
  await saveButton();
}

//search red tag

async function searchRedTag(){
  
  searchTechniquesPopup.style.display = 'block';
  tacticSearchField.innerHTML = '';
  helpersUI.blurBackground()
  await showPhaseOptions()
  jsonTechniques = await getTechniques();
  
}

async function displaySearchTechniques(){

  if (textBox.value == "") {
    alertSearchBox.innerHTML = 'Please fill out this field!';
  }
  else {
    searchTechniquesPopup.style.display = 'none';
    
    if (reverseJsonTechniques == "") reverseJsonTechniques = await searchTechniques()

    searchTechniquesArray = helpers.searchTechniquesFromJson(
  helpers.getSelectedRowText(phaseSearchField), 
  helpers.getSelectedRowText(tacticSearchField), 
  checkbox.checked, 
  textBox.value, 
  reverseJsonTechniques)

    helpersUI.drawFoundTechniquesTable(searchTechniquesArray)

    listSearchTechniquesPopup.style.display = 'block';
    textBox.value = '';
    checkbox.checked = ''
    alertSearchBox.innerHTML = ''
  }
}


function sortTechniquesTable(){

  var column = $(this).text().trim();
  switch (column) {
    case "Phase":
      searchTechniquesArray.sort((a, b) => a.phase.localeCompare(b.phase));
      break;
    case "Tactic":
      searchTechniquesArray.sort((a, b) => a.use.localeCompare(b.use));
      break;
    case "Technique":
      searchTechniquesArray.sort((a, b) => a.title.localeCompare(b.title));
      break;
    default:
      break;
    }
  helpersUI.drawFoundTechniquesTable(searchTechniquesArray)
}

function handleTableColorOnClick(){
  const classListString = this.classList.toString();
  if (classListString.includes("resultsTable")) {
    handleColorOnClickResultsTable(this);
  } else if (classListString.includes("list")) {
    helpersUI.handleColorOnClickPhaseTactic(this);
  }
}

function handleColorOnClickResultsTable(this1){

  const isSelected = this1.classList.toString().includes("selected");
  if (isSelected) {
    this1.classList.remove("selected");
    updateSearchTechniquesColor("white", this1);
    } 
  else {
    this1.classList.add("selected");
    updateSearchTechniquesColor("blue", this1);
  }
}

function updateSearchTechniquesColor(color, this_) {
  searchTechniquesArray.forEach(technique => {
    if (technique.id === this_.id) {
      technique.color = color;
    }
  });
}

async function saveButtonSearchTechniques(){
  await Word.run(async (context) => {

    const tagText = helpers.createTagTextSearch(table)

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

