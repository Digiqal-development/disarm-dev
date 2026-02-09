/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global     document, Office, Word */

import * as helpersUI from "../utils/helpers-ui.js";
import { changeRedTagColor, changeRedTagColorSaveBtn } from "../components/formatting-component.js";
import {
  insertRedTag,
  searchRedTag,
  saveButtonInsertTechniques,
  displaySearchTechniques,
  saveButtonSearchTechniques,
  clearSearchTechniquesArray,
} from "../components/red-tag-component.js";
import { insertRedTable } from "../components/table-summaries-component.js";
import { carouselPage1, carouselPage2, carouselPage3, carouselSkip } from "../constants/ui-elements.js";
import { searchTechniques } from "../services/api-services.js";
import { getRedTagColor } from "../components/formatting-component.js";

import {
    insertObjectTag,
    saveObjectTag,
    addObjectManually,
} from "../components/object-tag-component.js";

import {
    insertObjectSummaryTables,
} from "../components/object-summary-table-component.js";

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Word) {
    return;
  }
  const suggestions = await searchTechniques();
  //console.log(suggestions.ids);
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      const cursorPoint = selection.getRange(Word.RangeLocation.start);

      const currentParagraph = cursorPoint.paragraphs.getFirst();

      const paragraphStartPoint = currentParagraph.getRange(Word.RangeLocation.start);

      const rangeFromParaStartToCursor = paragraphStartPoint.expandTo(cursorPoint);

      // 5. Load the text of this paragraph-to-cursor range.
      rangeFromParaStartToCursor.load("text");

      await context.sync();
      const textBefore = rangeFromParaStartToCursor.text;
      let currentWord = "";

      if (textBefore && textBefore.length > 0) {
        let endIndex = textBefore.length; // The character just before the cursor is at textBefore.length - 1
        let startIndex = endIndex - 1;

        while (startIndex >= 0 && !isWordDelimiter(textBefore[startIndex])) {
          startIndex--;
        }

        currentWord = textBefore.substring(startIndex + 1, endIndex);
      }

      if (currentWord.startsWith("(")) {
        await showSuggestions(currentWord, suggestions);
      }
      console.log("Current word before cursor:", currentWord);
    }).catch(function (error) {
      console.error(error);
    });
  });

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

  // object tag
  document.getElementById("insert-object-tag").onclick = () => tryCatch(insertObjectTag);
  document.getElementById("save-object-tag").onclick = () => tryCatch(saveObjectTag);
    document.getElementById("add-object-button").onclick = () => {
        document.getElementById("popup-add-object").style.display = "block";
    };
    document.getElementById("confirm-add-object").onclick = () => tryCatch(addObjectManually);

    //summaries
  document.getElementById("insert-red-table").onclick = () => tryCatch(insertRedTable);

  //object summary table
  document.getElementById("insert-object-summary-table").onclick = () => tryCatch(insertObjectSummaryTables);

  //formatting
  document.getElementById("format-red-tag-color").onclick = () => tryCatch(changeRedTagColor);

  //save
  document.getElementById("save-btn").onclick = () => tryCatch(saveButtonInsertTechniques);
  document.getElementById("save-btn1").onclick = () => tryCatch(displaySearchTechniques);
  document.getElementById("save-btn2").onclick = () => tryCatch(changeRedTagColorSaveBtn);
  document.getElementById("save-btn3").onclick = () => tryCatch(saveButtonSearchTechniques);

  //close
  document.querySelectorAll(".close").forEach((element) => (element.onclick = () => tryCatch(closeButton)));
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

async function closeButton() {
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
  if (carouselPage2.style.display == "block") {
    carouselPage2.style.display = "none";
    document.getElementById("dot2").classList.remove("active");
    carouselPage1.style.display = "block";
    document.getElementById("dot1").classList.add("active");
  } else if (carouselPage3.style.display == "block") {
    carouselPage3.style.display = "none";
    document.getElementById("dot3").classList.remove("active");
    carouselPage2.style.display = "block";
    document.getElementById("dot2").classList.add("active");
  }
}

async function nextPage() {
  if (carouselPage1.style.display == "block") {
    carouselPage1.style.display = "none";
    document.getElementById("dot1").classList.remove("active");
    carouselPage2.style.display = "block";
    document.getElementById("dot2").classList.add("active");
  } else if (carouselPage2.style.display == "block") {
    carouselPage2.style.display = "none";
    document.getElementById("dot2").classList.remove("active");
    carouselPage3.style.display = "block";
    document.getElementById("dot3").classList.add("active");
  }
}

async function skipToFinish() {
  document.getElementById("first-run-experience").style.display = "none";
  document.getElementById("main").style.display = "block";
}

let currentWord = ""; // Keep track of the word being typed

// Function to show autocomplete suggestions in the task pane
const showSuggestions = async (query, suggestions) => {
  const suggestionsList = document.getElementById("autocomplete-suggestions");
  suggestionsList.innerHTML = "";
  console.log("USO U SUGG");
  if (query.length > 0) {
    const mappedSuggestions = Object.entries(suggestions.ids).map(([id, value]) => ({ id, ...value }));

    var parsedSuggestions = mappedSuggestions.map((item) => ({
      id: item.id,
      autoCompleteSuggestionText: `${item.title} [${item.id}]`,
      insertText: `(${item.title} [${item.id}])`,
    }));

    console.log(parsedSuggestions);

    const filteredSuggestions = parsedSuggestions.filter((item) =>
      item.autoCompleteSuggestionText.toLowerCase().startsWith(query.substring(1).toLowerCase())
    );

    filteredSuggestions.forEach((suggestion) => {
      const listItem = document.createElement("li");
      listItem.textContent = suggestion.autoCompleteSuggestionText;

      listItem.addEventListener("click", async function () {
        let partialSuggestion = suggestion.insertText.replace(query, "");

        await insertTextIntoWord(partialSuggestion, query);
        suggestionsList.innerHTML = ""; // Clear suggestions after selection
      });

      suggestionsList.appendChild(listItem);
    });
  }
};

// Function to insert selected autocomplete word into the document
const insertTextIntoWord = async (text, query) => {
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    var f = selection.getRange("End");

    const cursorRange = selection.getRange("Start");
    const searchOptions = {
      matchPrefix: true,
    };

    const searchResults = cursorRange.search("(", searchOptions);

    context.load(searchResults, "items/text");

    selection.insertText(text, Word.InsertLocation.replace);

    var rangeToHighlight = selection.search(text);
    rangeToHighlight.load("items/text");

    const rangeBefore = selection.getRange("Start").getTextRanges([".", " ", ","], false);
    var rangeToHighlight2 = rangeBefore.load("text");
    await context.sync();

    rangeToHighlight.items[0].font.set({ highlightColor: getRedTagColor() });
    rangeToHighlight2.items[0].font.set({ highlightColor: getRedTagColor() });
  }).catch(function (error) {
    console.error(error);
  });
};

function isWordDelimiter(char) {
    const delimiterRegex = /[\s.,;!?\")\\\]{}:<>-]/;
    return delimiterRegex.test(char);
}