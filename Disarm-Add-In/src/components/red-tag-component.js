import { getTechniques, searchTechniques } from '../services/api-services.js';
import { PHASES } from '../constants/global-variables.js';
import * as helpers from './../utils/helpers.js';
import * as helpersUI from '../utils/helpers-ui.js';
import { getRedTagColor } from '../components/formatting-component.js';

import { chooseTechniquesPopup, phaseChooseField, tacticChooseField, techniqueChooseCheckboxField, 
    searchTechniquesPopup, phaseSearchField, tacticSearchField, checkbox, textBox, 
    listSearchTechniquesPopup, alertSearchBox, table
  } from '../constants/ui-elements.js';

var jsonTechniques = '';
var reverseJsonTechniques = '';
var searchTechniquesArray = []

$(document).on('click','tr',function(e) { handleTableColorOnClick.call(this) }); 
$(document).on('click', 'th', function(e) {
  e.stopPropagation(); 
  sortTechniquesTable.call(this);
});

export async function insertRedTag() {
  await Word.run(async (context) => {
    helpersUI.blurBackground();
    await showPhaseOptions(true);
    helpersUI.clearInsertRedTagFields();
  });
}


export async function showPhaseOptions(insertTag){
    jsonTechniques = await getTechniques();
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
  
export async function saveButtonInsertTechniques(){
    await Word.run(async (context) => {
  
      if (reverseJsonTechniques == "") reverseJsonTechniques = await searchTechniques()
      const checkboxes = document.querySelectorAll('input[type="checkbox"]');
      const tagText = helpers.createTagTextInsert(checkboxes, reverseJsonTechniques);

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

export async function searchRedTag(){
  
  searchTechniquesPopup.style.display = 'block';
  tacticSearchField.innerHTML = '';
  helpersUI.blurBackground()
  await showPhaseOptions()  
}

export async function displaySearchTechniques(){
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

export async function saveButtonSearchTechniques(context){
  await Word.run(async (context) => {

    if (reverseJsonTechniques == "") reverseJsonTechniques = await searchTechniques()
    const tagText = helpers.createTagTextSearch(table, reverseJsonTechniques)

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

export function clearSearchTechniquesArray(){
  searchTechniquesArray = [];
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