import * as UI from "../constants/ui-elements.js";

export function blurBackground() {
  const elements = document.body.querySelectorAll("*:not(.popup)");

  elements.forEach((element) => {
    element.classList.add("blur-background");
  });
}

export function unblurBackground() {
  const elements = document.body.querySelectorAll("*:not(.popup)");
  elements.forEach((element) => {
    element.classList.remove("blur-background");
  });
}

export function closeAllFields() {
  UI.chooseTechniquesPopup.style.display = "none";
  UI.popupRed.style.display = "none";
  UI.searchTechniquesPopup.style.display = "none";
  UI.listSearchTechniquesPopup.style.display = "none";
  UI.phaseSearchField.innerHTML = "";
  UI.tacticSearchField.innerHTML = "";
  UI.phaseChooseField.innerHTML = "";
  UI.tacticChooseField.innerHTML = "";
  UI.textBox.value = "";
  UI.checkbox.checked = "";
  UI.alertSearchBox.innerHTML = "";
}

export function handleColorOnClickPhaseTactic(this_) {
  if (this_.classList.toString().includes("selected")) this_.classList.remove("selected");
  else {
    const selector = this_.classList.toString().includes("option1")
      ? ".option1.list.selected"
      : ".option2.list.selected";
    document.querySelectorAll(selector).forEach((row) => row.classList.remove("selected"));
    this_.classList.add("selected");
  }
}

export function drawFoundTechniquesTable1(searchTechniquesArray) {
  UI.table.innerHTML = "<thead><tr><th>Phase</th><th>Tactic</th><th>Technique</th></tr></thead>";

  for (const row in searchTechniquesArray) {
    var color = searchTechniquesArray[row].color;
    if (color == "blue") color = "rgb(85, 172, 227)";

    UI.table.innerHTML +=
      '<tbody><tr class="resultsTable" id="' +
      searchTechniquesArray[row].id +
      '"' +
      "><td>" +
      searchTechniquesArray[row].phase +
      "</td><td>" +
      searchTechniquesArray[row].use +
      "</td><td>" +
      "[" +
      searchTechniquesArray[row].id +
      "] " +
      searchTechniquesArray[row].title +
      "</td></tr></tbody>";
  }
}

export function drawFoundTechniquesTable(searchTechniquesArray) {
  UI.table.innerHTML = "<thead><tr><th>Phase</th><th>Tactic</th><th>Technique</th></tr></thead><tbody>";

  searchTechniquesArray.forEach(({ id, phase, use, title, color }, index) => {
    const row = document.createElement("tr");

    row.classList.add("resultsTable", index % 2 === 0 ? "odd-row" : "even-row");
    row.id = id;
    row.innerHTML = `
      <td>${phase}</td>
      <td>${use}</td>
      <td>[${id}] ${title}</td>
    `;

    UI.table.querySelector("tbody").appendChild(row);
  });

  // Close the tbody tag
  UI.table.innerHTML += "</tbody>";
}

export function clearInsertRedTagFields() {
  UI.tacticChooseField.innerHTML = "";
  UI.techniqueChooseCheckboxField.style.display = "none";
  UI.chooseTechniquesPopup.style.display = "block";
}

export function showTechniqueOptions(techniquesArray, tagsArray) {
  UI.techniqueChooseCheckboxField.innerHTML = "";
  let htmlContent = "";

  const idDescMap = Object.fromEntries(Object.entries(tagsArray.ids).map(([id, task]) => [id, task.description]));

  techniquesArray.forEach((technique) => {
    const key = Object.keys(technique)[0];
    const value = Object.values(technique)[0];
    htmlContent += `
    <div class="tooltip">
      <label>
        <input type="checkbox" value="${key}">[${key}] ${value}
      </label>
      <br>
      <span class="tooltiptext">${idDescMap[key]}</span>
    </div>
  `;
  });

  UI.techniqueChooseCheckboxField.innerHTML = htmlContent;
  UI.techniqueChooseCheckboxField.style.display = "block";
}
