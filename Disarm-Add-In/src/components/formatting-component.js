import * as helpersUI from '../utils/helpers-ui.js';
import {colorPicker, popupRed} from '../constants/ui-elements.js';
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