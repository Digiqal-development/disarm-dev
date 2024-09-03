const chooseTechniquesPopup = document.getElementById('popup-insert-red-tag');
const popupRed = document.getElementById('popup-red-tag-formatting');
const searchTechniquesPopup = document.getElementById('popup-search-techniques');
const listSearchTechniquesPopup = document.getElementById('popup-search-techniques-list');
const phaseChooseField = document.getElementById('option1');
const tacticChooseField = document.getElementById('option2');
const phaseSearchField = document.getElementById('option1-search');
const tacticSearchField = document.getElementById('option2-search');
const checkbox = document.getElementById('search-description');
const textBox = document.getElementById('search-bar');
const alertSearchBox = document.getElementById("alert")

export function blurBackground(){
  document.body.classList.add('blur-background');
}
  
export function unblurBackground(){
  document.body.classList.remove('blur-background');
}

export function closeAllFields() {
    chooseTechniquesPopup.style.display = 'none';
    popupRed.style.display = 'none';
    searchTechniquesPopup.style.display = 'none';
    listSearchTechniquesPopup.style.display = 'none';
    phaseSearchField.innerHTML = '';
    tacticSearchField.innerHTML = '';
    phaseChooseField.innerHTML = '';
    tacticChooseField.innerHTML = '';
    textBox.value = '';
    checkbox.checked = ''
    alertSearchBox.innerHTML = ''
}

export function handleColorOnClickPhaseTactic(this_){
    if (this_.classList.toString().includes("selected")) this_.classList.remove("selected")
    else {
        const selector = this_.classList.toString().includes("option1") ? '.option1.list.selected' : '.option2.list.selected';
        document.querySelectorAll(selector).forEach(row => row.classList.remove('selected'));
        this_.classList.add("selected");
    }
}


export function drawFoundTechniquesTable(searchTechniquesArray){

    const table = document.getElementById('search-results-table');
   
    table.innerHTML = "<thead><tr><th>Phase</th><th>Tactic</th><th>Technique</th></tr></thead>"
  
    for (const row in searchTechniquesArray){
  
      var color = searchTechniquesArray[row].color
      if (color == "blue") color = "rgb(85, 172, 227)";
      
      table.innerHTML += "<tbody><tr class=\"resultsTable\" id=\"" 
      + searchTechniquesArray[row].id + "\"" + "><td>" 
      + searchTechniquesArray[row].phase + "</td><td>" 
      + searchTechniquesArray[row].use + "</td><td>" + "["
      + searchTechniquesArray[row].id + "] "
      + searchTechniquesArray[row].title + "</td></tr></tbody>"
    }
}