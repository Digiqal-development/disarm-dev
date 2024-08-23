/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */

import { getTechniques, searchTechniques, getClauses } from './../services/apiServices.js';
import { extractTextFromBrackets, sortRedSummariesTable, removeExtraTagFromText } from './../utils/helpers.js';
//import { changeRedTagColor } from './../components/formattingComponent.js';

const fs = require('fs');
'use strict';

var jsonTechniques = '';
var reverseJsonTechniques = '';
var clausesGlobal = 'aaa';

var redTagColorGlobal = "#FFFE00"; 
var searchTechniquesArray = []

const SEARCH_TABLE_SELECT_COLOR = "rgb(85, 172, 227)"

const popup = document.getElementById('popup-insert-red-tag');
const option1 = document.getElementById('option1');
const option2 = document.getElementById('option2');
const option3 = document.getElementById('option3');
const saveBtn = document.getElementById('save-btn');


const popup1 = document.getElementById('popup-search-techniques');
const option1search = document.getElementById('option1-search');
const option2search = document.getElementById('option2-search');


const checkbox = document.getElementById('search-description');
const textBox = document.getElementById('search-bar');

const popup2 = document.getElementById('popup-search-techniques-list');
const alertSearchBox = document.getElementById("alert")

const table = document.getElementById('search-results-table');
const popupRed = document.getElementById('popup-red-tag-formatting');
  const colorPicker = document.getElementById('colorPicker');


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

$(document).on('click','tr',function(e) { changeTableColorOnSelect.call(this) }); 
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
    const popup1 = document.getElementById('popup-insert-red-tag');
    popup1.style.display = 'none';

    const popup2 = document.getElementById('popup-red-tag-formatting');
    popup2.style.display = 'none';

    const popup3 = document.getElementById('popup-search-techniques');
    popup3.style.display = 'none';

    const popup4 = document.getElementById('popup-search-techniques-list');
    popup4.style.display = 'none';

    const option1 = document.getElementById('option1-search');
    const option2 = document.getElementById('option2-search');

    option1.value = 'All';
    option2.value = '';

    const checkbox = document.getElementById('search-description');
    const textBox = document.getElementById('search-bar');
    const alertSearchBox = document.getElementById("alert")
    

    textBox.value = '';
    checkbox.checked = ''
    alertSearchBox.innerHTML = ''

    unblurBackground();

    searchTechniquesArray = []
  })
}




//red tag 

async function insertRedTag() {
  await Word.run(async (context) => {

    blurBackground();

    option1.innerHTML = drawPhaseOptions();

    option2.value = '';
    //option2.innerHTML = '';
    option3.style.display = 'none';

    popup.style.display = 'block';

    //getting data from the server only once
    if (jsonTechniques == '') jsonTechniques = await getTechniques();

    var selectedOption = option1.id;
    var chosenOption = jsonTechniques.plan;

    //listening for a change in the first dropdown
    option1.addEventListener('click', function() {
      selectedOption = option1.id;
      switch (selectedOption) {
        case 'plan':
            chosenOption = jsonTechniques.plan;
            break;
        case 'prepare':
            chosenOption = jsonTechniques.prepare;
            break;
        case 'execute':
            chosenOption = jsonTechniques.execute;
            break;
        case 'assess':
            chosenOption = jsonTechniques.assess;
            break;
        default:
            option2.innerHTML = '';
            break;
    }  
console.log("OK")
      // showing the change in the second dropdown
      option2.innerHTML = ''
      var jsonArray = Object.keys(chosenOption)

//////////////option2.innerHTML += '<option value=\"' + item +"\">" + item + "</option>"
      option2.innerHTML = "<table><thead></thead><tbody>"
      jsonArray.forEach(item => { 
        option2.innerHTML += "<tr><td>" + item  + "</td></tr>" });

      option2.innerHTML += "</tbody>"
        
      option2.style.display = 'block';
      option2.value = '';
    });

    option2.addEventListener('change', function() {
      // showing the techniques in a list of checkboxes

      selectedOption = option2.value;
      option3.style.display = 'block';

      let value = option2.value;
      let array = chosenOption[value]
      option3.innerHTML = ''
      array.forEach(item => { option3.innerHTML += '<label><input type=\"checkbox\" value=\"' + Object.keys(item)[0] +"\">[" + Object.keys(item)[0] + "] " + Object.values(item)[0] + "</label><br>"})
 
    });
  });
}

async function saveButton(){
  await Word.run(async (context) => {

    // get values from checkboxes
    
    const checkboxes = document.querySelectorAll('input[type="checkbox"]');
    //create string that represents the tag 
    let tagText = '('

    checkboxes.forEach(checkbox => {
      if (checkbox.checked) {
        tagText += removeExtraTagFromText(checkbox.nextSibling.textContent.trim()) + " [" + checkbox.value + "], "  
        }
    });

    tagText = tagText.slice(0, -2); 
    tagText += ')'
    if (tagText.length < 3) tagText = '';

    //clear all selectctions
    popup.style.display = 'none';
    option1.value = '';
    option2.value = '';
    option3.style.display = 'none';

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
    rangeToHighlight.items[0].font.set({highlightColor: redTagColorGlobal});
   
    tagText = ''

    unblurBackground();

    await context.sync(); 
  })
}

async function searchRedTag(){
  
  popup1.style.display = 'block';

  blurBackground()

  option2search.innerHTML = '';

  if (jsonTechniques == '') jsonTechniques = await getTechniques();

  var selectedOption = option1.value;
  var chosenOption = jsonTechniques.plan;

  //listening for a change in the first dropdown
  option1search.addEventListener('change', function() {
    
    switch (selectedOption) {
      case 'option1':
          chosenOption = jsonTechniques.plan;
          break;
      case 'option2':
          chosenOption = jsonTechniques.prepare;
          break;
      case 'option3':
          chosenOption = jsonTechniques.execute;
          break;
      case 'option4':
          chosenOption = jsonTechniques.assess;
          break;
      default:
          option2search.innerHTML = '';
          break;
  }  

      // showing the change in the second dropdown
      option2search.innerHTML = ''
      var jsonArray = Object.keys(chosenOption)
      jsonArray.forEach(item => { option2search.innerHTML += '<option value=\"' + item +"\">" + item + "</option>"});

      if (selectedOption === 'All') option2search.innerHTML = '';
        
      option2search.style.display = 'block';
      option2search.value = '';
    });

}

async function displaySearchTechniques(){

  if (textBox.value == "") {
    alertSearchBox.innerHTML = 'Please fill out this field!';
  }
  else {
    popup1.style.display = 'none';
    await searchTechniquesFromJson(option1search.value, option2search.value, checkbox.checked, textBox.value)
    popup2.style.display = 'block';
    textBox.value = '';
    checkbox.checked = ''
    alertSearchBox.innerHTML = ''
  }
}

async function searchTechniquesFromJson(phase, tactic, checkbox, textbox){

  var resultTechniques = []
  if (reverseJsonTechniques == "") reverseJsonTechniques = await searchTechniques()
  
  const newJson = reverseJsonTechniques.ids;
  for (const key in newJson){
    
    const value = newJson[key] 
    var found = false   

    //check if phase and tactic match the row
    if ((phase == "All" || phase == value.phase) && (tactic == "" || value.use.includes(tactic))){

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

  if (reverseJsonTechniques == '') await searchTechniques();
  const tag_ids = reverseJsonTechniques.ids

  for (const tag in resultTechniques){
    let tag_object = tag_ids[resultTechniques[tag]]
    let use = tag_object.use.replace(/\s*\[.*?\]\s*/, '');

    searchTechniquesArray.push({ id: resultTechniques[tag], phase: tag_object.phase, use: use, title: tag_object.title, color: "white" });
  }

  drawFoundTechniquesTable()

}

function drawFoundTechniquesTable(){

  const table = document.getElementById('search-results-table');
 
  table.innerHTML = "<thead><tr><th>Phase</th><th>Tactic</th><th>Technique</th></tr></thead>"

  for (const row in searchTechniquesArray){

    var color = searchTechniquesArray[row].color
    if (color == "blue") color = SEARCH_TABLE_SELECT_COLOR;
    
    table.innerHTML += "<tbody><tr class=\"item\" id=\"" + searchTechniquesArray[row].id 
    + "\"" + "style=\"background: " + color + "\"" +
    "><td>" + 
    searchTechniquesArray[row].phase + "</td><td>" 
     + searchTechniquesArray[row].use + "</td><td>" + "["+searchTechniquesArray[row].id+"] "+searchTechniquesArray[row].title + "</td></tr></tbody>"

  }
}

function sortTechniquesTable(){
  var column = $(this).text().trim();

  if (column == "Phase") searchTechniquesArray.sort((a,b) => a.phase.localeCompare(b.phase))
  else if (column == "Tactic") searchTechniquesArray.sort((a,b) => a.use.localeCompare(b.use))
  else if (column == "Technique") searchTechniquesArray.sort((a,b) => a.title.localeCompare(b.title))
  drawFoundTechniquesTable()
 
}

function changeTableColorOnSelect(){
 
  if ((this.style.background == "" || this.style.background =="white") && this.classList.toString() == "item"){
    $(this).css('background', '#55ace3');
    for (var i = 0; i < searchTechniquesArray.length; i++) {
      if (searchTechniquesArray[i].id === this.id) searchTechniquesArray[i].color = "blue";
   }
  }
  else { 
    $(this).css('background', 'white');
    for (var i = 0; i < searchTechniquesArray.length; i++) {
      if (searchTechniquesArray[i].id === this.id) searchTechniquesArray[i].color = "white";
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
      if (row.style.background == SEARCH_TABLE_SELECT_COLOR) {

        tagText += removeExtraTagFromText(row.cells[2].textContent) + " [" + row.id + "], ";    
           
      }
    }
    
    tagText = tagText.slice(0, -2); 
    tagText += ')'
    if (tagText.length < 3) tagText = '';

    popup2.style.display = 'none';
    option1search.value = 'All';
    option2search.value = '';
    unblurBackground();



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
    rangeToHighlight.items[0].font.set({highlightColor: redTagColorGlobal});

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
      let selectedText = extractTextFromBrackets(text, matches.index)

      if (foundMiddleBrackets.length != commasNumber + 1 || foundMiddleBrackets.length == 0) continue;

      for (let j = 0; j < foundMiddleBrackets.length; j++){
        
        let tag = foundMiddleBrackets[j].slice(1, -1)
        let tagObject = tag_ids[tag]

        selectedText = selectedText.replace(/^\s*/, '');
        
        // add matched cases into table
        table.push([tagObject.title, tag, selectedText, tagObject.use]) 
      }
    }
    table = sortRedSummariesTable(table);
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


//formatting

function changeRedTagColor(){
  colorPicker.value = redTagColorGlobal;
  popupRed.style.display = 'block';
  blurBackground()
}

async function changeRedTagColorSaveBtn(){
  var selectedColor = colorPicker.value;
  redTagColorGlobal = selectedColor;
  popupRed.style.display = 'none';
  unblurBackground()
}

function blurBackground(){
  document.body.classList.add('blur-background');
}

function unblurBackground(){
  document.body.classList.remove('blur-background');
}

function drawPhaseOptions(){
  return "<table id=\"option1\"><thead></thead><tbody><tr id=\"plan\"><td>Plan</td></tr><tr id=\"prepare\"><td>Prepare</td></tr><tr id=\"execute\"><td>Execute</td></tr><tr id=\"assess\"><td>assess</td></tr></tbody></table>"
}