/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */
import * as helpers from './../utils/helpers.js';
import * as helpersUI from './../utils/helpersUI.js';
import { getTechniques, searchTechniques, getClauses } from './../services/apiServices.js';
import { changeRedTagColor, changeRedTagColorSaveBtn, getRedTagColor } from './../components/formattingComponent.js';

var jsonTechniques = '';
var reverseJsonTechniques = '';

var searchTechniquesArray = []

const SEARCH_TABLE_SELECT_COLOR = "rgb(85, 172, 227)"

//insert red tag
const chooseTechniquesPopup = document.getElementById('popup-insert-red-tag');
const phaseChooseField = document.getElementById('option1');
const tacticChooseField = document.getElementById('option2');
const techniqueChooseCheckboxField = document.getElementById('option3');

//search red tag
const searchTechniquesPopup = document.getElementById('popup-search-techniques');
const phaseSearchField = document.getElementById('option1-search');
const tacticSearchField = document.getElementById('option2-search');

const checkbox = document.getElementById('search-description');
const textBox = document.getElementById('search-bar');

const listSearchTechniquesPopup = document.getElementById('popup-search-techniques-list');
const alertSearchBox = document.getElementById("alert")

const table = document.getElementById('search-results-table');


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

    tacticChooseField.innerHTML = '';
    techniqueChooseCheckboxField.style.display = 'none';
    chooseTechniquesPopup.style.display = 'block';
  });
}


function showPhaseOptions(insertTag){
  var uiField = phaseChooseField;
  if (!insertTag) uiField = phaseSearchField;
  uiField.innerHTML = '';
  var phases = ["Plan", "Prepare", "Execute", "Assess"]
  for(let i = 0; i < phases.length; i++){
    let row = document.createElement('tr');
    row.addEventListener('click', function() {
      showTacticOptions(phases[i].toLowerCase(), insertTag);
    });
    let cols = '<td>' + phases[i] + '</td>';
    row.innerHTML = cols;
    row.className += " list option1";
    if (i % 2 === 0) {
      row.classList += ' even-row';
  } else {
      row.classList += ' odd-row';
  }
    uiField.appendChild(row);
}
}

function showTacticOptions(phase, insertTag){
  techniqueChooseCheckboxField.innerHTML = ''
  var uiField = tacticChooseField;
  if (!insertTag) uiField = tacticSearchField;
  uiField.innerHTML = '';
 let tacticsArray = jsonTechniques[phase]
 uiField.innerHTML = '';
  Object.keys(tacticsArray).forEach((tactic, index) => {
    let row = document.createElement('tr');
    if (insertTag) row.addEventListener('click', function() {showTechniqueOptions(tacticsArray[tactic])});
    let cols = '<td>' + tactic + '</td>';
    row.innerHTML = cols;
    row.className += " list option2";
    if (index % 2 === 0) {
      row.classList += ' even-row';
  } else {
      row.classList+= ' odd-row';
  }
    uiField.appendChild(row);

  })
}

function showTechniqueOptions(techniquesArray){
  techniqueChooseCheckboxField.innerHTML = '';
  techniquesArray.forEach(technique => {
    techniqueChooseCheckboxField.innerHTML += '<label><input type=\"checkbox\" value=\"' + Object.keys(technique)[0] +"\">[" + Object.keys(technique)[0] + "] " + Object.values(technique)[0] + "</label><br>";
  })
  techniqueChooseCheckboxField.style.display = 'block';
}

async function saveButton(){
  await Word.run(async (context) => {

    // get values from checkboxes
    const checkboxes = document.querySelectorAll('input[type="checkbox"]');

    //create string that represents the tag 
    let tagText = '('

    checkboxes.forEach(checkbox => {
      if (checkbox.checked) {
        tagText += helpers.removeExtraTagFromText(checkbox.nextSibling.textContent.trim()) + " [" + checkbox.value + "], "  
        }
    });

    tagText = tagText.slice(0, -2); 
    tagText += ')'
    if (tagText.length < 3) tagText = '';

    //clear all selectctions
    chooseTechniquesPopup.style.display = 'none';
    phaseChooseField.innerHTML = '';
    tacticChooseField.innerHTML = '';
    techniqueChooseCheckboxField.style.display = 'none';

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



