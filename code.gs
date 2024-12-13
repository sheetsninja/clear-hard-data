/*
@OnlyCurrentDoc
*/

function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu("Script Menu").addItem("Reset Values",'clearHardData').addToUi();
}

function clearHardData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let selection = sheet.getActiveRange();
  let startRow = selection.getRow();
  let startCol = selection.getColumn();
  let selectedFormulas = selection.getFormulas();

  for (i=0;i<selectedFormulas.length;i++){
    let thisRow = selectedFormulas[i];
    for (j=0;j<thisRow.length;j++){
        if (thisRow[j] == '' || thisRow[j] == null) {
            sheet.getRange(startRow+i,startCol+j).clearContent();
          }
    }
  }
}
