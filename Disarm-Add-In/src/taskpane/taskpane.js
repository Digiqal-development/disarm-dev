/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */


const fs = require('fs');
'use strict';

var jsonTechniques = '';
var reverseJsonTechniques = '';

var redTagColorGlobal = "#FFFF00";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {

    //red tag
    document.getElementById("insert-red-tag").onclick = () => tryCatch(insertRedTag);
    document.getElementById("search-red-tag").onclick = () => tryCatch(searchRedTag);

    //summaries
    document.getElementById("insert-red-table").onclick = () => tryCatch(insertRedTable);

    //formatting
    document.getElementById("format-red-tag-color").onclick = () => tryCatch(changeRedTagColor);

    //save
    document.getElementById("save-btn").onclick = () => tryCatch(saveButton);
    document.getElementById("save-btn1").onclick = () => tryCatch(displaySearchTechniques);
    document.getElementById("save-btn2").onclick = () => tryCatch(changeRedTagColorSaveBtn);

    //close
    document.querySelectorAll('.close').forEach(element => element.onclick = () => tryCatch(closeButton));

  }
});

$(document).on('click','tr',function(e) { changeTableColorOnSelect.call(this) }); 

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

  })
}


//red tag 

async function insertRedTag() {
  await Word.run(async (context) => {

    //initialization of used components
    const popup = document.getElementById('popup-insert-red-tag');
    const option1 = document.getElementById('option1');
    const option2 = document.getElementById('option2');
    const option3 = document.getElementById('option3');
    const saveBtn = document.getElementById('save-btn');

    option1.value = '';
    option2.value = '';
    option2.innerHTML = '';
    option3.style.display = 'none';

    popup.style.display = 'block';

    //getting data from the server only once
    if (jsonTechniques == '') await getTechniques();

    planOption = jsonTechniques.plan
    chosenOption = planOption

    //listening for a change in the first dropdown
    option1.addEventListener('change', function() {
      selectedOption = option1.value;

      if (selectedOption === 'option1') {
        planOption = jsonTechniques.plan
        chosenOption = planOption
        jsonArray = Object.keys(planOption)
          
      } else if (selectedOption === 'option2') {
        prepareOption = jsonTechniques.prepare
        chosenOption = prepareOption
        jsonArray = Object.keys(prepareOption)

      } else if (selectedOption === 'option3'){
        executeOption = jsonTechniques.execute
        chosenOption = executeOption
        jsonArray = Object.keys(executeOption)

      } else if (selectedOption === 'option4'){
        assessOption = jsonTechniques.assess
        chosenOption = assessOption
        jsonArray = Object.keys(assessOption)

      } else { 
          option2.innerHTML = '';
      }

      // showing the change in the second dropdown
      option2.innerHTML = ''

      jsonArray.forEach(item => { option2.innerHTML += '<option value=\"' + item +"\">" + item + "</option>"});
        
      option2.style.display = 'block';
      option2.value = '';
    });

    option2.addEventListener('change', function() {
      // showing the techniques in a list of checkboxes

      selectedOption.option2 = option2.value;
      option3.style.display = 'block';

      let value = option2.value;
      let array = chosenOption[value]
      option3.innerHTML = ''
      array.forEach(item => { option3.innerHTML += '<label><input type=\"checkbox\" value=\"' + Object.keys(item)[0] +"\">" + Object.values(item)[0] + "</label><br>"})
 
    });
  });
}

async function saveButton(){
  await Word.run(async (context) => {

    // get values from checkboxes
    const checkboxes = document.querySelectorAll('input[type="checkbox"]');
    const popup = document.getElementById('popup-insert-red-tag');

    let text = '('

    checkboxes.forEach(checkbox => {
      if (checkbox.checked) {
          text += checkbox.nextSibling.textContent.trim() + " [" + checkbox.value + "], "
        }
    });

    
    text = text.slice(0, -2); 

    text += ')'

    popup.style.display = 'none';
    option1.value = '';
    option2.value = '';
    option3.style.display = 'none';

    var doc = context.document.getSelection();
    context.load(doc)
    var selectedLength = await context.sync().then(() => {return doc.text.length })

    // if nothing is selected get the entire sentence
    if (selectedLength < 1){
      var doc1 = context.document.getSelection().getTextRanges(['\n', '.', '?'], false);
      context.load(doc1);
      doc =  await context.sync().then(() => { return doc1.items[0]})
    }

    //highlght the text
    doc.font.set({highlightColor: redTagColorGlobal});
    
    //insert text with tags
    doc.insertText(text, Word.InsertLocation.end);
    text = ''

    await context.sync(); 
  })
}

async function getTechniques(){
  $.ajax({
    dataType: "json",
    url: "https://api.jsonbin.io/v3/b/651a869054105e766fbc92de",
    data: {
      'X-Master-Key': '111111$2a$10$aGFnKNoSPkQ7mmpPofdEqe9aZSNnEVBqKCXEkuwyl5OWz8kDfPbw',
    },
    success: function (result, status, xhr) {
      jsonTechniques = result.record
    },
    error: function (xhr, status, error) {
      console.log(error) 
  }})
}

async function searchRedTag(){
  const popup = document.getElementById('popup-search-techniques');
  popup.style.display = 'block';

  const option1 = document.getElementById('option1-search');
  const option2 = document.getElementById('option2-search');

  option2.innerHTML = '';

  if (jsonTechniques == '') await getTechniques();

    planOption = jsonTechniques.plan
    chosenOption = planOption

    //listening for a change in the first dropdown
    option1.addEventListener('change', function() {
      selectedOption = option1.value;

      if (selectedOption === 'Plan') {
        planOption = jsonTechniques.plan
        chosenOption = planOption
        jsonArray = Object.keys(planOption)
          
      } else if (selectedOption === 'Prepare') {
        prepareOption = jsonTechniques.prepare
        chosenOption = prepareOption
        jsonArray = Object.keys(prepareOption)

      } else if (selectedOption === 'Execute'){
        executeOption = jsonTechniques.execute
        chosenOption = executeOption
        jsonArray = Object.keys(executeOption)

      } else if (selectedOption === 'Assess'){
        assessOption = jsonTechniques.assess
        chosenOption = assessOption
        jsonArray = Object.keys(assessOption)

      } else { 
          option2.innerHTML = '';
      }

      // showing the change in the second dropdown
      option2.innerHTML = ''

      jsonArray.forEach(item => { option2.innerHTML += '<option value=\"' + item +"\">" + item + "</option>"});

      if (selectedOption === 'All') option2.innerHTML = '';
        
      option2.style.display = 'block';
      option2.value = '';
    });

}

async function displaySearchTechniques(){

  const option1 = document.getElementById('option1-search');
  const option2 = document.getElementById('option2-search');
  const checkbox = document.getElementById('search-description');
  const textBox = document.getElementById('search-bar');
  const popup1 = document.getElementById('popup-search-techniques');
  const popup2 = document.getElementById('popup-search-techniques-list');

  popup1.style.display = 'none';

  await searchTechniquesFromJson(option1.value, option2.value, checkbox.checked, textBox.value)

  popup2.style.display = 'block';
}

async function searchTechniquesFromJson(phase, tactic, checkbox, textbox){

  var resultTechniques = []
  if (reverseJsonTechniques == "") await searchTechniques()
  
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

  await drawFoundTechniquesTable(resultTechniques)
}

async function drawFoundTechniquesTable(results){

  const table = document.getElementById('search-results-table');
  
  if (reverseJsonTechniques == '') await searchTechniques();
  const tag_ids = reverseJsonTechniques.ids

  table.innerHTML = "<tr><th>Phase</th><th>Tactic</th><th>Technique</th></tr>"

  for (const tag in results){
    let tag_object = tag_ids[results[tag]]

    table.innerHTML += "<tr class=\"item\"><td>" + tag_object.phase + "</td><td>" 
    + tag_object.use + "</td><td>" + tag_object.title + "</td></tr>"

  }
}

function changeTableColorOnSelect(){
 
  if ((this.style.background == "" || this.style.background =="white") && this.classList.toString() == "item"){
    $(this).css('background', '#55ace3');
  }
  else { 
    $(this).css('background', 'white');
  }
}


//summaries

async function insertRedTable() {

  await Word.run(async (context) => {
    
    //get search json only once
    if (reverseJsonTechniques == '') await searchTechniques();
    

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
    let table = [["Tecnique title", "ID", "Text", "Use"]]

    for (const matches of foundSmallBrackets){
  
      foundMiddleBrackets = matches[0].match(middleBracketRegex);
      const commasNumber = (matches[0].match(/,/g) || []).length;
      let selectedText = extractTextFromBrackets(text, matches.index)

      if (foundMiddleBrackets.length != commasNumber + 1 || foundMiddleBrackets.length == 0) continue;

      for (let j = 0; j < foundMiddleBrackets.length; j++){
        
        let tag = foundMiddleBrackets[j].slice(1, -1)
        let tagObject = tag_ids[tag]
        
        // add matched cases into table
        table.push([tagObject.title, tag, selectedText, tagObject.use]) 
      }
    }
    drawTable(table)
  }
)}

async function drawTable(table){
  await Word.run(async (context) => {

    const secondParagraph = context.document.body.paragraphs.getFirst().getNext()
    const insertedTable = secondParagraph.insertTable(table.length, 4, Word.InsertLocation.after, table);
    insertedTable.styleBuiltIn = Word.BuiltInStyleName.listTable3_Accent1;
})

};

function extractTextFromBrackets(str, endIndex) {
  const pattern = /[.!?]/;
  let s = str.substring(0, endIndex - 2)
  let index = s.lastIndexOf(pattern.exec(s)) + 1

  return str.substring(index, endIndex);
}

async function searchTechniques(){
  
  $.ajax({
    dataType: "json",
    url: "https://api.jsonbin.io/v3/b/655cf2010574da7622c9d82b",
    data: {
      'X-Master-Key': '111111$2a$10$aGFnKNoSPkQ7mmpPofdEqe9aZSNnEVBqKCXEkuwyl5OWz8kDfPbw',
    },
    success: function (result, status, xhr) {
      var resp = result.record
      reverseJsonTechniques =resp
    },
    error: function (xhr, status, error) {
      console.log(error) 
  }})
}


//formatting

function changeRedTagColor(){
  const popup = document.getElementById('popup-red-tag-formatting');
  const colorPicker = document.getElementById('colorPicker');
  colorPicker.value = redTagColorGlobal;

  popup.style.display = 'block';

}

async function changeRedTagColorSaveBtn(){
  const popup = document.getElementById('popup-red-tag-formatting');
  var selectedColor = document.getElementById('colorPicker').value;
  redTagColorGlobal = selectedColor;
  popup.style.display = 'none';
}






