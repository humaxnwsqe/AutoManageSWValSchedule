function myFunction() {
  console.log('Hello, world!')
}


function myColorFunctionB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getSheetByName("Form Responses 1").getRange(2,6,ss.getLastRow());
  var cellRange = range.getValues();

  for(i = 0; i<cellRange.length; i++){
     if(cellRange[i][0] == "Open")
     {
       ss.getSheetByName("Form Responses 1").getRange(i+2,6).setBackground("red");
       ss.getSheetByName("Form Responses 1").getRange(i+2,6).setFontColor('white');
     }
  }
}



function myColorFunctionA() {
  var s = SpreadsheetApp.getActiveSheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getSheetByName("Form Responses 1").getRange(2,6,ss.getLastRow());
  var cellRange = range.getValues();
  Logger.log(cellRange);
  Logger.log(cellRange.length);
  Logger.log(cellRange.valueOf());

  for(i = 1; i<cellRange.length; i++){
     if(cellRange[i] == "Open")
     {
       Logger.log("change color here");
     } else {
       Logger.log("don't change color");
     } 
  }
}



/*
function countColoredCells(countRange) {
  var activeRg = countRange;
  var activeSht = SpreadsheetApp.getActiveSheet();
  var activeformula = activeRg.getFormula();
  var countRangeAddress = activeformula.match(/\((.*)\,/).pop().trim();
  var backGrounds = activeSht.getRange(countRangeAddress).getBackgrounds();
  var colorRefAddress = activeformula.match(/\,(.*)\)/).pop().trim();
  var BackGround = activeSht.getRange(colorRefAddress).getBackground();
  var countCells = 0;
  for (var i = 0; i < backGrounds.length; i++)
    for (var k = 0; k < backGrounds[i].length; k++)
      if ( backGrounds[i][k] == BackGround )
        countCells = countCells + 1;
  return countCells;
}
*/

/*
function countbackgrounds() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  var range_input = book.getRange("B3:B4");
  var range_output = book.getRange("B6");
  var cell_colors = range_input.getBackgroundColors();
  var color = "#58FA58";
  var count = 0;

  for( var i in cell_colors ){
  Logger.log(cell_colors[i][0])
    if( cell_colors[i][0] == color ){ ++count }
    }
  range_output.setValue(count);
}
*/

/*
function countColoredCells(countRange,colorRef) {
  var activeRange = SpreadsheetApp.getActiveRange();
  var activeSheet = activeRange.getSheet();
  var formula = activeRange.getFormula();
  
  var rangeA1Notation = formula.match(/\((.*)\,/).pop();
  var range = activeSheet.getRange(rangeA1Notation);
  var bg = range.getBackgrounds();
  var values = range.getValues();
  
  var colorCellA1Notation = formula.match(/\,(.*)\)/).pop();
  var colorCell = activeSheet.getRange(colorCellA1Notation);
  var color = colorCell.getBackground();
  
  var count = 0;
  
  for(var i=0;i<bg.length;i++)
    for(var j=0;j<bg[0].length;j++)
      if( bg[i][j] == color )
        count=count+1;
  return count;
}
*/