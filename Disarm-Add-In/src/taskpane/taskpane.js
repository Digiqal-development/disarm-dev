/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global     document, Office, Word */

const pluginDialog = document.getElementById('plugin-dialog');
const optionSelect = document.getElementById('option-select');
const applyBtn = document.getElementById('apply-btn');
const redTags = []

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert-red-tag").onclick = () => tryCatch(insertRedTag);
    document.getElementById("insert-red-table").onclick = () => tryCatch(insertRedTable);

    //document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
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
    

    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.font.set({highlightColor: "Yellow"});

    const selectedChoices = {
      option1: null,
      option2: null,
      option3: [],
    };



    const pluginContainer = document.getElementById('plugin-container');
    const option1 = document.getElementById('option1');
    const option2 = document.getElementById('option2');
    const option3 = document.getElementById('option3');
    const option4 = document.getElementById('option4');
    const saveBtn = document.getElementById('save-btn');

  // Show the div as a popup
    pluginContainer.style.display = 'block';

  // Listen for changes in the first dropdown
    option1.addEventListener('change', function() {
      // Get the selected option
      selectedOption = option1.value;

      // Simulate loading options based on the selection
      // Here, we'll populate the second dropdown with different values based on the first pick
      if (selectedOption === 'option1') {
          option2.innerHTML = '<option value="suboption11">Plan Strategy</option><option value="suboption12">Plan Objectives</option><option value="suboption13">Target Audience Analysis</option>';
      } else if (selectedOption === 'option2') {
          option2.innerHTML = '<option value="suboption21">Develop Narratives</option><option value="suboption22">Develop Content</option><option value="suboption23">Establish Social Assets</option><option value="suboption24">Establish Legitimacy</option><option value="suboption25">Microtarget</option><option value="suboption26">Select Channels And Affordances</option>';
      } else if (selectedOption === 'option3'){
        option2.innerHTML = '<option value="suboption31">Condunt Pump Priming</option><option value="suboption32">Deliver Content</option><option value="suboption33">Maximize Exposure</option><option value="suboption34">Drive Online Harms</option><option value="suboption35">Drive Offline Activity</option><option value="suboption36">Persist in the Information Enviroment</option>';
      } else if (selectedOption === 'option4'){
        option2.innerHTML = '<option value="suboption41">Asses Effectiveness</option>';
      } else {
          option2.innerHTML = '';
      }

      // Show the second dropdown
      option2.style.display = 'block';
  });

  // Listen for changes in the second dropdown
  option2.addEventListener('change', function() {
      // Show the third subcomponent
      selectedOption.option2 = option2.value;
      option3.style.display = 'block';
  });

  const checkboxes = option3.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(checkbox => {
      checkbox.addEventListener('change', function() {
          // Update the selected choices in the third subcomponent
          selectedChoices.option3 = Array.from(checkboxes)
              .filter(checkbox => checkbox.checked)
              .map(checkbox => checkbox.value);
      });
  });

  // Listen for the Save button click
  saveBtn.addEventListener('click', function() {
      // Hide the div

      redTags.push(selectedChoices);
      

      pluginContainer.style.display = 'none';

      // Clear selections and reset visibility
      option1.value = '';
      option2.value = '';
      option3.style.display = 'none';
  });

 
    await context.sync();
  });
}

async function insertRedTable() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to get a reference to the paragraph
      //        that will precede the table.
      const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

      // TODO2: Queue commands to create a table and populate it with data.

      const tableData = [
        ["Name", "ID", "Birth City"],
        ["Bob", "434", "Chicago"],
        ["Sue", "719", "Havana"],


    ];
      

    
      secondParagraph.insertTable(3, 3, Word.InsertLocation.after, tableData);

      await context.sync();
  });
}


