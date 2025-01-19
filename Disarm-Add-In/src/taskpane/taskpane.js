/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */

import * as helpersUI from '../utils/helpers-ui.js';
import { changeRedTagColor, changeRedTagColorSaveBtn } from '../components/formatting-component.js';
import { insertRedTag, searchRedTag, saveButtonInsertTechniques, displaySearchTechniques, 
  saveButtonSearchTechniques, clearSearchTechniquesArray } from '../components/red-tag-component.js';
import { insertRedTable } from '../components/table-summaries-component.js';
import { carouselPage1, carouselPage2, carouselPage3, 
  carouselSkip
} from '../constants/ui-elements.js'


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {

    document.getElementById("first-run-experience").style.display = "none";
    document.getElementById("main").style.display = "block";
    
    if (!localStorage.getItem("showedFRE")) {
       showFirstRunExperience();
     }
    
    //carousel
    document.getElementById("welcome-page-button").onclick = () => tryCatch(welcomePage);

    document.getElementById("prev-btn1").onclick = () => tryCatch(prevPage);
    document.getElementById("prev-btn2").onclick = () => tryCatch(prevPage);

    document.getElementById("next-btn1").onclick = () => tryCatch(nextPage);
    document.getElementById("next-btn2").onclick = () => tryCatch(nextPage);

    document.getElementById("skip1").onclick = () => tryCatch(skipToFinish);
    document.getElementById("skip2").onclick = () => tryCatch(skipToFinish);
    document.getElementById("get-started-button").onclick = () => tryCatch(skipToFinish);
      
    //red tag
    document.getElementById("insert-red-tag").onclick = () => tryCatch(insertRedTag);
    document.getElementById("search-red-tag").onclick = () => tryCatch(searchRedTag);

    //summaries
    document.getElementById("insert-red-table").onclick = () => tryCatch(insertRedTable);

    //formatting
    document.getElementById("format-red-tag-color").onclick = () => tryCatch(changeRedTagColor);

    //save
    document.getElementById("save-btn").onclick = () => tryCatch(saveButtonInsertTechniques);
    document.getElementById("save-btn1").onclick = () => tryCatch(displaySearchTechniques);
    document.getElementById("save-btn2").onclick = () => tryCatch(changeRedTagColorSaveBtn);
    document.getElementById("save-btn3").onclick = () => tryCatch(saveButtonSearchTechniques);

    //close
    document.querySelectorAll('.close').forEach(element => element.onclick = () => tryCatch(closeButton));

  }
});

async function showFirstRunExperience() {
  document.getElementById("first-run-experience").style.display = "block";
  document.getElementById("main").style.display = "none";
  localStorage.setItem("showedFRE", true);
}  

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

async function closeButton(){
  helpersUI.closeAllFields();
  helpersUI.unblurBackground();
  clearSearchTechniquesArray();
}

async function welcomePage() {
  document.getElementById("welcome-page").style.display = "none";
  document.getElementById("carousel").style.display = "block";
  carouselPage1.style.display = "block";
  document.getElementById("dot1").classList.add("active");
}

async function prevPage() {
  if (carouselPage2.style.display == "block"){
    carouselPage2.style.display = "none";
    document.getElementById("dot2").classList.remove("active");
    carouselPage1.style.display = "block";
    document.getElementById("dot1").classList.add("active");
  }
  else if (carouselPage3.style.display == "block"){
    carouselPage3.style.display = "none";
    document.getElementById("dot3").classList.remove("active");
    carouselPage2.style.display = "block";
    document.getElementById("dot2").classList.add("active");
  } 
}

async function nextPage() {
  if (carouselPage1.style.display == "block"){
    carouselPage1.style.display = "none";
    document.getElementById("dot1").classList.remove("active");
    carouselPage2.style.display = "block";
    document.getElementById("dot2").classList.add("active");
  }
  else if (carouselPage2.style.display == "block"){
    carouselPage2.style.display = "none";
    document.getElementById("dot2").classList.remove("active");
    carouselPage3.style.display = "block";
    document.getElementById("dot3").classList.add("active");
  } 
}

async function skipToFinish(){
  document.getElementById("first-run-experience").style.display = "none";
  document.getElementById("main").style.display = "block";
}






