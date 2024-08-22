export function changeRedTagColor(){
    const popup = document.getElementById('popup-red-tag-formatting');
    const colorPicker = document.getElementById('colorPicker');
    colorPicker.value = redTagColorGlobal;
  
    popup.style.display = 'block';
    document.body.classList.add('blur-background');
  
  }