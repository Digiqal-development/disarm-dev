/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global     document, Office, Word */
const fs = require('fs');
'use strict';

const pluginDialog = document.getElementById('plugin-dialog');
const optionSelect = document.getElementById('option-select');
const applyBtn = document.getElementById('apply-btn');
const redTags = []

var redTagColorGlobal = "#FFFF00";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert-red-tag").onclick = () => tryCatch(insertRedTag);
    document.getElementById("search-red-tag").onclick = () => tryCatch(searchRedTag);
    document.getElementById("insert-red-table").onclick = () => tryCatch(insertRedTable);
    document.getElementById("format-red-tag-color").onclick = () => tryCatch(changeRedTagColor);
    document.getElementById("save-btn").onclick = () => tryCatch(saveButton);
    document.getElementById("save-btn1").onclick = () => tryCatch(displaySearchTechniques);
    document.getElementById("save-btn2").onclick = () => tryCatch(changeRedTagColorSaveBtn);
    document.getElementById("close-btn").onclick = () => tryCatch(closeButton);
    document.getElementById("close-btn1").onclick = () => tryCatch(closeButton);
    document.getElementById("close-btn2").onclick = () => tryCatch(closeButton);
  }
});

async function tryCatch(callback) {
  try {
      await callback();
  } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
  }
}

async function insertRedTag() {
  await Word.run(async (context) => {

    const pluginContainer = document.getElementById('plugin-container');
    const option1 = document.getElementById('option1');
    const option2 = document.getElementById('option2');
    const option3 = document.getElementById('option3');
    const saveBtn = document.getElementById('save-btn');


    option1.value = '';
    option2.value = '';
    option2.innerHTML = '';
    option3.style.display = 'none';

    pluginContainer.style.display = 'block';

    $.ajax({
      dataType: "json",
      url: "https://api.jsonbin.io/v3/b/651a869054105e766fbc92de",
      data: {
        'X-Master-Key': '111111$2a$10$aGFnKNoSPkQ7mmpPofdEqe9aZSNnEVBqKCXEkuwyl5OWz8kDfPbw',
      },
      success: function (result, status, xhr) {
    
          dataObject = result.record
          dataObject1 = dataObject.plan
          planOption = dataObject.plan
          chosenOption = planOption
          
          
      },
      error: function (xhr, status, error) {
              console.log(error) 
      }})

    option1.addEventListener('change', function() {
      selectedOption = option1.value;

      if (selectedOption === 'option1') {
        planOption = dataObject.plan
        chosenOption = planOption
        jsonArray = Object.keys(planOption)
          
      } else if (selectedOption === 'option2') {
        prepareOption = dataObject.prepare
        chosenOption = prepareOption
        jsonArray = Object.keys(prepareOption)

      } else if (selectedOption === 'option3'){
        executeOption = dataObject.execute
        chosenOption = executeOption
        jsonArray = Object.keys(executeOption)


      } else if (selectedOption === 'option4'){
        assessOption = dataObject.assess
        chosenOption = assessOption
        jsonArray = Object.keys(assessOption)
      } else { 
          option2.innerHTML = '';
      }

      // Show the second dropdown
      option2.innerHTML = ''
      jsonArray.forEach(item => { option2.innerHTML += '<option value=\"' + item +"\">" + item + "</option>"});
        
      option2.style.display = 'block';
      option2.value = '';
  });

  option2.addEventListener('change', function() {
      // Show the third subcomponent
      selectedOption.option2 = option2.value;
      option3.style.display = 'block';

      let value = option2.value;
      let array = chosenOption[value]
      option3.innerHTML = ''
      array.forEach(item => { option3.innerHTML += '<label><input type=\"checkbox\" value=\"' + Object.keys(item)[0] +"\">" + Object.values(item)[0] + "</label><br>"})
 

  });

  });
}

async function insertRedTable() {

  await Word.run(async (context) => {

    const doc = context.document
    
    $.ajax({
      dataType: "json",
      url: "https://api.jsonbin.io/v3/b/6527deaf12a5d376598ac60e",
      headers: {'X-Master-Key': '$2a$10$aGFnKNoSPkQ7mmpPofdEqe9aZSNnEVBqKCXEkuwyl5OWz8kDfPbw.' },
      data: {
        'X-Master-Key': '$2a$10$aGFnKNoSPkQ7mmpPofdEqe9aZSNnEVBqKCXEkuwyl5OWz8kDfPbw.',
      },
      success: async function (result, status, xhr) {
    
          dataObject = result.record

      const tag_ids = dataObject.ids
      
    

    var doc1 = context.document.body
    doc1.load("text")
    await context.sync();

    var text = doc1.text

    const smallBracketRegex = /\(([^()]*)\)/g;
    //const foundSmallBrackets1 = text.match(smallBracketRegex);

   const foundSmallBrackets = text.matchAll(/\(([^()]*)\)/g);


    const middleBracketRegex = /\[(.*?)\]/g;
    let foundMiddleBrackets = []

    let table = [["Tecnique title", "ID", "Text", "Use"]]

    

    for (const matches of foundSmallBrackets){
  
      foundMiddleBrackets = matches[0].match(middleBracketRegex);
      const commasNumber = (matches[0].match(/,/g) || []).length;
      let selectedText = extractText(text, matches.index)

      if (foundMiddleBrackets.length != commasNumber + 1 || foundMiddleBrackets.length == 0) continue;

      for (let j = 0; j < foundMiddleBrackets.length; j++){
        
        let tag = foundMiddleBrackets[j].slice(1, -1)
        let tagObject = tag_ids[tag]

        
        table.push([tagObject.title, tag, selectedText, tagObject.use])
        
      }


    }

    drawTable(table)


          
  },
  error: function (xhr, status, error) {
          console.log(error) 
  }})


 
  }


    
  
)}

async function drawTable(table){
  await Word.run(async (context) => {

    const secondParagraph = context.document.body.paragraphs.getFirst().getNext()
    const insertedTable = secondParagraph.insertTable(table.length, 4, Word.InsertLocation.after, table);
    insertedTable.styleBuiltIn = Word.BuiltInStyleName.listTable3_Accent1;
})

};

async function saveButton(){
  await Word.run(async (context) => {

  const checkboxes = document.querySelectorAll('input[type="checkbox"]');
  const pluginContainer = document.getElementById('plugin-container');

      let text = '('

      checkboxes.forEach(checkbox => {
        if (checkbox.checked) {
            text += checkbox.nextSibling.textContent.trim() + " [" + checkbox.value + "], "
        }

      
    });

    
    text = text.slice(0, -2); 

    text += ')'

    pluginContainer.style.display = 'none';

      // Clear selections and reset visibility
    option1.value = '';
    option2.value = '';
    option3.style.display = 'none';
 

    
 

  var doc = context.document.getSelection();
  context.load(doc)
  var selectedLenth = await context.sync().then(() => {return doc.text.length })


  if (selectedLenth < 1){
   var doc1 = context.document.getSelection().getTextRanges(['\n', '.', '?'], false);
   context.load(doc1);
    //var sentence = await context.sync().then(() => {return doc1.items[0].text;})
  
   doc =  await context.sync().then(() => { return doc1.items[0]})
   

  
  }

  doc.font.set({highlightColor: redTagColorGlobal});
    

   doc.insertText(text, Word.InsertLocation.end);
    text = ''



    await context.sync(); 
  })

}

async function closeButton(){
  await Word.run(async (context) => {
    const pluginContainer = document.getElementById('plugin-container');
    pluginContainer.style.display = 'none';

    const pluginContainer1 = document.getElementById('plugin-container-red-tag');
    pluginContainer1.style.display = 'none';

    const pluginContainer2 = document.getElementById('plugin-container-search-techniques');
    pluginContainer2.style.display = 'none';


    
  })
}

function extractText(str, endIndex) {
  const pattern = /[.!?]/;
  let s = str.substring(0, endIndex - 2)
  let index = s.lastIndexOf(pattern.exec(s)) + 1

  return str.substring(index, endIndex);
}

function changeRedTagColor(){
  const pluginContainer = document.getElementById('plugin-container-red-tag');
  const colorPicker = document.getElementById('colorPicker');
  colorPicker.value = redTagColorGlobal;

  pluginContainer.style.display = 'block';

}

async function changeRedTagColorSaveBtn(){
  const pluginContainer = document.getElementById('plugin-container-red-tag');
  var selectedColor = document.getElementById('colorPicker').value;
  redTagColorGlobal = selectedColor;
  pluginContainer.style.display = 'none';
}

async function searchRedTag(){
  const pluginContainer = document.getElementById('plugin-container-search-techniques');
  pluginContainer.style.display = 'block';
}

async function displaySearchTechniques(){
  const pluginContainer = document.getElementById('plugin-container-search-techniques');
  pluginContainer.style.display = 'none';

  const pluginContainer1 = document.getElementById('plugin-container-search-techniques-list');
  pluginContainer1.style.display = 'block';
}

