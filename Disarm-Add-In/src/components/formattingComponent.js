import * as helpersUI from './../utils/helpersUI.js';

const popupRed = document.getElementById('popup-red-tag-formatting');
const colorPicker = document.getElementById('colorPicker');
let redTagColor = "#FFFE00"; 

export function changeRedTagColor(){
  colorPicker.value = redTagColor;
  popupRed.style.display = 'block';
  helpersUI.blurBackground()
}

export function changeRedTagColorSaveBtn(){
  var selectedColor = colorPicker.value;
  redTagColor = selectedColor;
  popupRed.style.display = 'none';
  helpersUI.unblurBackground()
}

export function getRedTagColor(){
  return redTagColor;
}