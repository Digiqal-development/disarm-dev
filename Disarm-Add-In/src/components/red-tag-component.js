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
  
export async function saveButton(){
    await Word.run(async (context) => {
  
      const checkboxes = document.querySelectorAll('input[type="checkbox"]');
      const tagText = helpers.createTagTextInsert(checkboxes);

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