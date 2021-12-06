///////////////////////
/**
 * @file name: Main.gs
 * @description
 * NW SW_Validation_Schedule.xlsx 파일 활용에 필요한 업무 처리를 자동화 한 파일
 */
///////////////////////

/**
 * @function name: onOpen()
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * onOpen은 미리 약속된 함수로 파일을 새로 열때 자동으로 실행하고 싶은 코드들을 담으면 된다.
 * 현재 해당 함수에는 
 * (a) 누적된 이벤트들을 일괄 처리 하는 코드와
 * (b) 커스텀 메뉴를 생성하는 코드가 포함되어 있다.
 * 커스텀 메뉴 경우 현재 주차 위치를 그려주는 함수를 수동으로 실행하기 위한 목적이 있다.
 * (function 'fillCurrentWeekRangeBG')
 */

function onOpen(e) {
  //누적된 이벤트들을 일괄 처리
  SpreadsheetApp.flush();

  //커스텀 메뉴를 생성
  var menu = SpreadsheetApp.getUi().createMenu('Custom Menu');
  menu.addItem("Fill Current Week Position", "fillCurrentWeekRangeBG");
  menu.addToUi();

  //위와 같이 변수 하나를 만들고 필요한 메뉴를 하나씩 붙일 수 있지만 아래 같은 방식으로도 가능하다.
  // Or DocumentApp or FormApp.
  //ui.createMenu('Custom Menu')
  //    .addItem('First item', 'menuItem1')
  //    .addSeparator()
  //    .addSubMenu(ui.createMenu('Sub-menu')
  //        .addItem('Second item', 'menuItem2'))
  //    .addToUi();
}

/**
 * @function name: onEdit()
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * onEdit은 문서 내 편집이 발생 했을 때 자동 실행이 필요한 것들을 실해하도록 지원한다.
 * 
 */

function onEdit(e){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //Logger.log(sheet.getName())

  var range = e.range;

  Logger.log(range.getColumn());  
  Logger.log(range.getNumColumns());
  //range.setNote('Modified range info: ' + range.getColumn() + '/' + range.getNumColumns());

  //일정 변경 sheet, range를 입력 인자로 받아 해당 위치의 HQ, VN 검증 모델 수를 카운트 하기 위한 함수 호출
  updateMacro(sheet, range);

  
}
//*/

/**
 * @function name: updateMacro()
 * @input variable(s)
 * sheet: Active Sheet의 sheet 객체 
 * range: 수정이 발생한 영역(range)에 대한 객체 
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * 최초 onEdit 함수에서 구현 되어 있던 내용을 별도로 분리한 케이스 임.
 */

function updateMacro(sheet, range){
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var targetRangeHQ = sheet.getRange(5, range.getColumn(), 1, range.getNumColumns());
  var targetRangeVN = sheet.getRange(6, range.getColumn(), 1, range.getNumColumns());

  //targetRange.setNote('Modified targetRange info: ' + targetRange.getColumn() + '/' + targetRange.getNumColumns());

  //range.setNote('Test: ' + range.getRow() + '/' + range.getColumn() + '/' + range.getNumRows());
  //range.setNote('Test: ' + targetRange.getRow() + '/' + targetRange.getColumn() + '/' + targetRange.getNumRows());

  //Logger.log(targetRange)
  
  for(i = 0; i < targetRangeHQ.getNumColumns(); i++){    
     updateFuncInRange(targetRangeHQ.getRow(), targetRangeHQ.getColumn() + i);    
  }

  for(i = 0; i < targetRangeVN.getNumColumns(); i++){
     updateFuncInRange(targetRangeVN.getRow(), targetRangeVN.getColumn() + i);
  }
}

/** 
 * @function name: updateFuncInRange()
 * @author: Hyucksu Lee
 * @update date: 2021.11.25
 * @description
 * 해당 위치에 있는 구글 앱스 스크립트 함수를 강제로 갱신하게 끔 처리하는 함수 
 */

function updateFuncInRange(row, col){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var orig = sheet.getRange(row,col).getFormula(); 
  //var temp = orig.replace("=", "?");
  sheet.getRange(row,col).setFormula(""); 
  SpreadsheetApp.flush();
  sheet.getRange(row,col).setFormula(orig); 
  //SpreadsheetApp.flush();
}

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

