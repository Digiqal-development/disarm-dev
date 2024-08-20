/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */


const fs = require('fs');
'use strict';

var jsonTechniques = '';
var reverseJsonTechniques = '';
var clausesGlobal = 'aaa';

var redTagColorGlobal = "#FFFE00"; 
var searchTechniquesArray = []

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {

    //get json files with listed tecniques and tags
    getTechniques();
    searchTechniques();

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

    document.body.classList.remove('blur-background');

    searchTechniquesArray = []
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

    document.body.classList.add('blur-background');

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
      array.forEach(item => { option3.innerHTML += '<label><input type=\"checkbox\" value=\"' + Object.keys(item)[0] +"\">[" + Object.keys(item)[0] + "] " + Object.values(item)[0] + "</label><br>"})
 
    });
  });
}

async function saveButton(){
  await Word.run(async (context) => {

    // get values from checkboxes
    const checkboxes = document.querySelectorAll('input[type="checkbox"]');
    const popup = document.getElementById('popup-insert-red-tag');

    //create string that represents the tag 
    let tagText = '('

    checkboxes.forEach(checkbox => {
      if (checkbox.checked) {
        var checkboxText = checkbox.nextSibling.textContent.trim()
        tagText += checkboxText.indexOf(' ') === -1 ? checkboxText : checkboxText.substring(checkboxText.indexOf(' ') + 1) + " [" + checkbox.value + "], "
        
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

    document.body.classList.remove('blur-background');

    await context.sync(); 
  })
}

async function getClauses(sentence) {
  try {
    const result = await $.ajax({
      type: 'POST',
      dataType: 'json',
      url: 'https://localhost:7225/clauses',
      contentType: 'application/json',
      data: JSON.stringify({
        sentence: sentence.text,
        result: 'string'
      })
    });
    return result;
  } catch (error) {
    console.error('Error fetching clauses:', error);
    throw error; 
  }
}


async function getTechniques(){
  
  $.ajax({
    type: 'GET',
    headers: {
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache',
      'Expires': '0'
    },
    dataType: "json",
    url: "https://localhost:7225/techniques",
    success: function (result, status, xhr) {
      jsonTechniques = result
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

  document.body.classList.add('blur-background');

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
  const alertSearchBox = document.getElementById("alert")

  

  if (textBox.value == "") {
    alertSearchBox.innerHTML = 'Please fill out this field!';
  }
  else {
    
    popup1.style.display = 'none';
    await searchTechniquesFromJson(option1.value, option2.value, checkbox.checked, textBox.value)
   
    popup2.style.display = 'block';
    textBox.value = '';
    checkbox.checked = ''
    alertSearchBox.innerHTML = ''
  }


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
    if (color == "blue") color = "rgb(85, 172, 227)"
    
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

    const popup = document.getElementById('popup-search-techniques-list');
    const table = document.getElementById('search-results-table');
    const option1 = document.getElementById('option1-search');
    const option2 = document.getElementById('option2-search');

    let tagText = '('

    //iterate through the table
    for (let i = 0; i < table.rows.length; i++) {
      const row = table.rows[i];

      //if row is colored add its content to the text
      if (row.style.background == "rgb(85, 172, 227)") {
        
        const tableTagText = row.cells[2].textContent;
        tagText += tableTagText.indexOf(' ') === -1 ? tableTagText : tableTagText.substring(tableTagText.indexOf(' ') + 1) + " [" + row.id + "], ";    
           
      }
    }
    
    tagText = tagText.slice(0, -2); 
    tagText += ')'
    if (tagText.length < 3) tagText = '';

    popup.style.display = 'none';
    option1.value = 'All';
    option2.value = '';
    document.body.classList.remove('blur-background');



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
    let table = [["Technique title", "ID", "Text", "Use"]]

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

    const secondParagraph = context.document.body.paragraphs.getLast()
    const insertedTable = secondParagraph.insertTable(table.length, 4, Word.InsertLocation.after, table);
    insertedTable.styleBuiltIn = Word.BuiltInStyleName.listTable3_Accent1;
})

};

function extractTextFromBrackets(str, endIndex) {  
  //pattern to match evenrything that ends a sentence
  const pattern = /[.!?\n\t]/;

  //pattern to match ]) - in case that after a sentence there is a tag
  const patternBrackets = /]\)/;

  //taking string from the beginning to end index
  let s = str.substring(0, endIndex - 2)

  console.log("TEKST:", s)
  console.log(patternBrackets.exec(s))
  console.log(pattern.exec(s))


  

  //taking the latest index of these patterns
  let index = (s.lastIndexOf(pattern.exec(s))> s.lastIndexOf(patternBrackets.exec(s))) ? s.lastIndexOf(pattern.exec(s)) + 1 : s.lastIndexOf(patternBrackets.exec(s)) + 1;
   
  console.log(str.substring(index, endIndex))
  return str.substring(index, endIndex);
}

async function searchTechniques(){
  
  $.ajax({
    type: 'GET',
    headers: {
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache',
      'Expires': '0'
  },
    dataType: "json",
    url: "https://localhost:7225/tags",
    success: function (result, status, xhr) {
      var resp = result
      reverseJsonTechniques = resp
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
  document.body.classList.add('blur-background');

}

async function changeRedTagColorSaveBtn(){
  const popup = document.getElementById('popup-red-tag-formatting');
  var selectedColor = document.getElementById('colorPicker').value;
  redTagColorGlobal = selectedColor;
  popup.style.display = 'none';
  document.body.classList.remove('blur-background');
}


//test
async function test12(){
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

async function test1(){
  await Word.run(async (context) => {
  var doc1 = context.document.getSelection().getTextRanges(['\n', '.', '?'], false);
      
  context.load(doc1);
  doc =  await context.sync().then(() => { return doc1.items[0]})

  //doc.text = "nesto drugo"
  
  //send sentence to backend for clauses extraction
  //await getClauses(doc).then(()=>  console.log(clausesGlobal));

  var clauses = await getClauses(doc);
  var string = clauses.result


  let jsonString = string.replace(/'/g, '"'); // Removes leading and trailing single quotes


// Step 2: Use JSON.parse() to convert the JSON string to an array
  let array = JSON.parse(jsonString);


  
  var rangeToHighlight = doc.search(array[0]);

  rangeToHighlight.load('items');
  await context.sync();
  console.log(rangeToHighlight.items[0].text);
  rangeToHighlight.items[0].font.color = 'red';

  //var clausesArray = JSON.parse(clauses.result)
  //console.log(clausesArray[0])
  await context.sync();
  })
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

      await context.sync();

      console.log(`Queried custom XML part for ${xpathToQueryFor} and found ${clientResult.value.length} matches:`);
      for (let i = 0; i < clientResult.value.length; i++) {
        console.log(clientResult.value[i]);
      }
    } else {
      console.warn("Didn't find custom XML part to query");
    }
  });
}






