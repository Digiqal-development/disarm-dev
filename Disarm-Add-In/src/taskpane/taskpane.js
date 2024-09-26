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
    document.getElementById("save-btn").onclick = () => tryCatch(saveButtonInsertTechniques);
    document.getElementById("save-btn1").onclick = () => tryCatch(displaySearchTechniques);
    document.getElementById("save-btn2").onclick = () => tryCatch(changeRedTagColorSaveBtn);
    document.getElementById("save-btn3").onclick = () => tryCatch(saveButtonSearchTechniques);

    //close
    document.querySelectorAll('.close').forEach(element => element.onclick = () => tryCatch(closeButton));

  }
});


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





