/**
 * 
 * @function name: getCurrentWeekInfo()
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * NW SW_Validation_Schedule.xlsx의 활성화 된 sheet 'NW_SQE_YYYY_Schedule'의 'A5' 위치에
 * 현재 주차를 W20 형식으로 보여주도록 매크로가 반영 되어 있음. 
 * 이 매크로 계산 값을 읽어와 리턴하기 위한 함수 
 */

 function getCurrentWeekInfo() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();   
    //var sheet = SpreadsheetApp.getActiveSheet();   
    var onecell = sheet.getRange('A5');   
    //console.log('Current week is ' + onecell.getValue()); 
    return onecell.getValue();
 }

  /**
 * @function name: countOnBGColorYellow()
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * 노란색으로 채운 셀 개수를 카운트 하기 위한 함수
 * sheet 내 해당 함수가 작성된 셀들이 여러 개 있음. 입력 인자인 'countRange' 역시 해당 셀에 작성된 함수 내에 포함되어 있음.
 */

function countOnBGColorYellow(countRange) {
    var activeRange = SpreadsheetApp.getActiveRange();
    var activeSheet = activeRange.getSheet();
    var formula = activeRange.getFormula();
    
    var rangeA1Notation = formula.match(/\(\"(.*)\"\)/).pop();
    var range = activeSheet.getRange(rangeA1Notation);
    var bg = range.getBackgrounds();
    //var values = range.getValues();
    
    //var colorCellA1Notation = formula.match(/\,(.*)\)/).pop();
    //var colorCell = activeSheet.getRange(colorCellA1Notation);
    var color = "#ffff00";
    
    var count = 0;
    
    for(var i=0;i<bg.length;i++)
      for(var j=0;j<bg[0].length;j++)
        if( bg[i][j] == color )
          count=count+1;
  
    return count;
}

/**
 * @function name: countOnBGColorOrange_2019()
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * 오렌지색으로 채운 셀 개수를 카운트 하기 위한 함수
 * sheet 내 해당 함수가 작성된 셀들이 여러 개 있음. 입력 인자인 'countRange' 역시 해당 셀에 작성된 함수 내에 포함되어 있음.
 */

 function countOnBGColorOrange_2019(countRange) {
    var activeRange = SpreadsheetApp.getActiveRange();
    var activeSheet = activeRange.getSheet();
    var formula = activeRange.getFormula();
    
    var rangeA1Notation = formula.match(/\(\"(.*)\"\)/).pop();
    var range = activeSheet.getRange(rangeA1Notation);
    var bg = range.getBackgrounds();
    //var values = range.getValues();
    
    //var colorCellA1Notation = formula.match(/\,(.*)\)/).pop();
    //var colorCell = activeSheet.getRange(colorCellA1Notation);
    var color = "#ff9900";
    
    var count = 0;
    
    for(var i=0;i<bg.length;i++)
      for(var j=0;j<bg[0].length;j++)
        if( bg[i][j] == color )
          count=count+1;
  
    return count;
}


  /**
 * @function name: fillCurrentWeekRangeBG()
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * 현재 주차를 초록색으로 칠하기 위한 함수
 */

function fillCurrentWeekRangeBG() {
    SpreadsheetApp.flush();
  
    //Read weeks data from the range of "H1:DI1"
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var basicrange = ss.getActiveSheet().getRange("H1:DI1");
    var values = basicrange.getValues();
    var ColNums = basicrange.getNumColumns();
    var RowNums = basicrange.getNumRows();
    var startColPos = basicrange.getColumn()
  
  
    //Log the data that was read.
    //Logger.log(JSON.stringify(data));
    //console.log(JSON.stringify(ColNums));
    //console.log(JSON.stringify(RowNums));
    //console.log('Start col position info is ' + startColPos);
  
    //console.log(getCurrentWeekInfo());
  
    for (i = 0; i < ColNums; i++) {
        if (values[0][i] !== "" && values[0][i] == getCurrentWeekInfo()){
            //console.log(values[0][i]);
            //console.log('target col is ' + i);
            //console.log('current cell info is ' + ss.getActiveSheet().getRange(RowNums, startColPos + i,RowNums, 2).setBackground("#09f726d"))
            //console.log('current cell info is ' + ss.getActiveSheet().getRange(RowNums, startColPos + i,RowNums, 2).setBackground("#09f76d"))
            ss.getActiveSheet().getRange(RowNums, startColPos + i,RowNums, 2).setBackground("#09f76d");
            //console.log('current cell\'s value is ' + ss.getActiveSheet().getRange(RowNums, startColPos + i).getValue())
            if (i - 2 >= 0) {
              //console.log('current cell info is ' + ss.getActiveSheet().getRange(RowNums, startColPos + i - 2,RowNums, 2).setBackground("#ffff00"))
              ss.getActiveSheet().getRange(RowNums, startColPos + i - 2,RowNums, 2).setBackground("#ffff00");
            }
              
            return;
        }
        
        //if (i == 10){
        //  return;
        //}
    }
}
