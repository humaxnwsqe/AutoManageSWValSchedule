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
  //menu.addItem("Ready to make Schedule State Update", "createEditTrigger");
  //menu.addItem("Update Count Colored Cells Number", "updateCountMacro");
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
 * (update, 2021.12.15) onEdit은 simple trigger 이며 실행 시간이 30 sec limit이 있어 trigger를 통해 여러 함수들이
 * 실행할 경우 30 sec가 넘는 경우도 종종 발생하기도 한다. 이런 문제를 해결하기 위해 onEdit 대신 
 * programmatic trigger 방식으로 해결 했다. 그래서 onEdit는 주석 처리 함
 * 
 */

/* function onEdit(e){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //Logger.log(sheet.getName())

  var range = e.range;

  console.log("\'onEdit()\' The start position of changed range: row is " + range.getRow())
  console.log("\'onEdit()\' The start position of changed range: column is " + range.getColumn());  
  console.log("\'onEdit()\' The number of changed row is " + range.getNumRows())
  console.log("\'onEdit()\' The number of changed column is " + range.getNumColumns());

  //일정 변경 sheet, range를 입력 인자로 받아 해당 위치의 HQ, VN 검증 모델 수를 카운트 하기 위한 함수 호출
  //if 조건은 편집이 발생한 range 시작 Row가 7 이상 그리고 Column은 8 이상 일 경우만 updage macro를 수행 하기 위한 조건
  if (range.getRow() >= UPDATE_VAL_RANGE_ROW_START 
    && range.getColumn() >= UPDATE_VAL_RANGE_COL_START){ 
    //runScript();
    //SpreadsheetApp.flush();
  }
  
} */


/**
 * @function name: runScript
 * @author: Hyucksu Lee 
 * @update date: 2021.12.14
 * @description
 * 이 runScript는 trigger 처리 관련 테스트를 위한 함수로 original runScript와 이름은 동일하지만
 * 목적은 다름. 테스트 목적 외에는 전체 주석 처리 함 
 */
/* function runScript(eventInfo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log(JSON.stringify(eventInfo));

  var range = eventInfo.range;
    
  console.log("\'runScript()\' The start position of changed range: row is " + range.getRow())
  console.log("\'runScript()\' The start position of changed range: column is " + range.getColumn());  
  console.log("\'runScript()\' The number of changed row is " + range.getNumRows())
  console.log("\'runScript()\' The number of changed column is " + range.getNumColumns());


  console.log("run script function has runned.!");
  console.log(`Current Active Sheet name is ${sheet.getName()}`);
} */

/**
 * @function name: createEditTrigger
 * @author: Hyucksu Lee
 * @update date: 2021.12.15
 * @description
 * onOpen이나 onEdit 함수와 같은 simple trigger가 아닌 programmatic trigger이며
 * 코드 상으로 onEdit 성격의 trigger를 생성해 준다.
 */

function createEditTrigger(){
  //var triggers = ScriptApp.getProjectTriggers()
  console.log(`ScriptApp.getProjectTriggers().length (before createEditTrigger) is ${ScriptApp.getProjectTriggers().length} !`);


  if(checkIfTriggerExists(ScriptApp.EventType.ON_EDIT, "runScript") == false){
    ScriptApp.newTrigger("runScript")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  }else {
    console.log(`Programmed edit trigger is already existed!`);
  }

  console.log(`ScriptApp.getProjectTriggers().length (after createEditTrigger) is ${ScriptApp.getProjectTriggers().length} !`);
  
  return;
}

/**
 * 
 * @function name: checkIfTriggerExists 
 * @author: Hyucksu Lee
 * @update date: 2021.12.15
 * @params
 * eventType: ON_EDIT
 * handlerFunction: runScript 함수 이름, string type
 * @description
 * createEditTrigger에서 triggering 되는 runScript event가 여러 번 생성되지 않도록 체크하기 위한 함수
 */
function checkIfTriggerExists(eventType, handlerFunction) {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;

  console.log(`checkIfTriggerExists() / eventType is ${eventType} !`);
  console.log(`checkIfTriggerExists() / handlerFunction is ${handlerFunction} !`);

if(triggers.length > 0){
  for(var i=0; i<triggers.length; i++){
    console.log(`triggers.getEventType() is ${triggers[i].getEventType()}`);
    console.log(`triggers.getHandlerFunction() is ${triggers[i].getHandlerFunction()}`);
    if(triggers[i].getEventType() === eventType &&
      triggers[i].getHandlerFunction() === handlerFunction){
        console.log(`Inside of if condition in trigger.getHandlerFunction is ${triggers[i].getHandlerFunction()}`)
        triggerExists = true;
      }
  }
}
/* 
  triggers.forEach(function (trigger) {
    if(trigger.getEventType() === eventType &&
      trigger.getHandlerFunction() === handlerFunction){
        console.log(`trigger.getHandlerFunction is ${trigger.getHandlerFunction()}`)
        triggerExists = true;
      }
      
  }); */

  console.log(`triggerExists state is ${triggerExists} !`);

  return triggerExists;
}

/* function varInit(){
  var remProp = PropertiesService.getScriptProperties();
  remProp.getProperty('remRun');
  remProp.setProperty('remRun', 'stop');
  h = 0;
  countHQ = 0;
  countVN = 0;
} */

/**
 * @function name: runScript()
 * @input variable(s)
 * evnetInfo: simple trigger 함수에서 e와 같은 역할 임 
 * @author: Hyucksu Lee
 * @update date: 2021.12.15
 * @return: null
 * @description
 * 최초 onEdit 함수에서 구현 되어 있던 내용을 별도로 분리한 케이스 임.
 * 이후 HQ, VN 셀 업데이트에 따라 검증 수 카운트를 구분해 업데이트 하도록 코드 수정
 */

function runScript(eventInfo) {
  //해당 함수 시작 시간 체크
  startTime = (new Date()).getTime() / 1000;
  console.log("start time is " + startTime);

  //현재 active 상태의 sheet 객체를 가져오기 위함
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  console.log(`Current Active Sheet name is ${sheet.getName()}`);

  var range = eventInfo.range;

  //변경 이벤트 발생 영역의 시작점을 체크. 변경 영역 시작점이 유효 범위가
  //아닐 경우 runScript trigger 삭제 처리하고 runScript 종료 처리
  if (range.getRow() < UPDATE_VAL_RANGE_ROW_START 
    || range.getColumn() < UPDATE_VAL_RANGE_COL_START){
      console.log('변경 영역 시작점이 유효하지 않음. ON_EDIT 트리거 삭제 처리.');
      deleteTriggers('runScript'); 
      return;
  }

  console.log("\'runScript()\' The start position of changed range: row is " + range.getRow())
  console.log("\'runScript()\' The start position of changed range: column is " + range.getColumn());  
  console.log("\'runScript()\' The number of changed row is " + range.getNumRows())
  console.log("\'runScript()\' The number of changed column is " + range.getNumColumns());

  var h = 0;
  var countHQ = 0;
  var countVN = 0;

  //일정 편집이 발생한 위치에서 HQ가 업데이트 된 것인지, VN이 업데이트 된 것인지 등을 먼저 확인해 개수 정보르 갖고 있기 위함
  var editedcellnums = parseInt(range.getNumRows() * range.getNumColumns());

  console.log(`\`runScript()\` Edited number of cells is(are) ${editedcellnums} and current h is ${h} !!`);

  var bg = range.getBackgrounds();


  console.log("\'runScript()\' The number of initial HQ background cell is " + countHQ);
  console.log("\'runScript()\' The number of initial VN background cell is " + countVN);

  if (h == 0) {
    for(var i=0;i<bg.length;i++){
      for(var j=0;j<bg[0].length;j++){
        
        if( bg[i][j] == HQcolor ){
          //console.log("\'runScript()\' Current cell background info is " + bg[i][j])
          //console.log("\'runScript()\' The number of counted HQ background cell is " + countHQ)
          countHQ=countHQ+1;
        }else if(bg[i][j] == VNcolor){
          //console.log("\'runScript()\' Current cell background info is " + bg[i][j])
          //console.log("\'runScript()\' The number of counted VN background cell is " + countVN)
          countVN=countVN+1;
        }
      }
    }
  }  

  console.log("\'runScript()\' The number of counted HQ background cell is " + countHQ);
  console.log("\'runScript()\' The number of counted VN background cell is " + countVN);

  //일정 편집이 발생하면 현황을 업데이트 하기 위한 위치 정보를 변수로 만들고 경우에 따라 업데이트가 필요한
  //부분만 업데이트 하는 등 변경 이벤트 발생을 최소화 함 
  var targetRangeHQ = sheet.getRange(5, range.getColumn(), 1, range.getNumColumns());
  var targetRangeVN = sheet.getRange(6, range.getColumn(), 1, range.getNumColumns());

  while (h < editedcellnums){
    console.log(`\'runScript()\' while loop entered and h: ${h} \, edited cell nums : ${editedcellnums}`);

    if (countHQ > 0 && countVN == 0) {
      console.log(`(Case 1) Only HQ cells updated case!`);
      for(i = 0; i < targetRangeHQ.getNumColumns(); i++){ 
  
        if (i == 0) {
          console.log("\'runScript()\' The start position of targetRange(HQ) info: row is " + targetRangeHQ.getRow())
          console.log("\'runScript()\' The start position of targetRange(HQ) info: column is " + targetRangeHQ.getColumn())
        }else if(i > 0){
          console.log("\'runScript()\' The cur_update position of targetRange(HQ) info: row is " + targetRangeHQ.getRow())
          console.log("\'runScript()\' The cur_update position of targetRange(HQ) info: column is " + (targetRangeHQ.getColumn() + i))
        }  
        
        updateFuncInRange(targetRangeHQ.getRow(), targetRangeHQ.getColumn() + i);   
        
  
        currentTime = (new Date()).getTime() / 1000;
        console.log("Run time is " + (currentTime - startTime));        
        
        h++; 
      }
    }else if (countVN > 0 && countHQ == 0) {
      console.log(`(Case 2) Only VN cells updated case!`);
      for(j = 0; j < targetRangeVN.getNumColumns(); j++){
        console.log("\'runScript()\' The start position of targetRange(VN) info: row is " + targetRangeVN.getRow())
        console.log("\'runScript()\' The start position of targetRange(VN) info: column is " + targetRangeVN.getColumn())
        console.log("\'runScript()\' j is " + j)
        updateFuncInRange(targetRangeVN.getRow(), targetRangeVN.getColumn() + j);
        
  
        currentTime = (new Date()).getTime() / 1000;
        console.log("current time is " + currentTime);
  
        console.log("지나간 시간은 ? " + (currentTime - startTime) );       
        
        h++;
      }
    }else {
      console.log(`(Case 3) Both HQ and VN cellS updated case! `);
      console.log(`or`);
      console.log(`(Case 4) CellS cleared case!`);
      for(i = 0; i < targetRangeHQ.getNumColumns(); i++){   
        console.log("\'runScript()\' The start position of targetRange(HQ) info: row is " + targetRangeHQ.getRow())
        console.log("\'runScript()\' The start position of targetRange(HQ) info: column is " + targetRangeHQ.getColumn())
        console.log("\'runScript()\' i is " + i)
        updateFuncInRange(targetRangeHQ.getRow(), targetRangeHQ.getColumn() + i);   
        h++;         
      }

      for(j = 0; j < targetRangeVN.getNumColumns(); j++){
        //console.log("\'runScript()\' The start position of targetRange(VN) info: row is " + targetRangeVN.getRow())
        //console.log("\'runScript()\' The start position of targetRange(VN) info: column is " + targetRangeVN.getColumn())
        //console.log("\'runScript()\' j is " + j)
        updateFuncInRange(targetRangeVN.getRow(), targetRangeVN.getColumn() + j);
        h++;        
      }
    }

    console.log(`##current h is ${h} !!!!!!!!!!!!!!!!`)
  }

  console.log('모든 카운트 작업 완료. 관련 자원 삭제 처리.');
  deleteTriggers('runScript');
  //SpreadsheetApp.flush();
  return;
}


/**
 * @function name: deleteTriggers
 * @author: Hyucksu Lee
 * @update date: 2021.12.15
 * @returns: null
 * @description
 * runScript 실행 완료 되면 마지막 단에 runScript trigger event를 제거하기 위함
 */
function deleteTriggers(funcName)
{
  var triggers = ScriptApp.getProjectTriggers();

  console.log(`\'deleteTriggers()\' function name is ${funcName} !!!!!!!!!!!!!!!!!!!!!!!!`)

  if (triggers.length > 0){    
    for(var i=0; i<triggers.length; i++){      
      if(triggers[i].getHandlerFunction() === funcName){
        console.log(`target trigger function should be deleted !!`);
        ScriptApp.deleteTrigger(triggers[i]); 
      }  
    }   
  }  

  createEditTrigger();

  return;
}

/** 
 * @function name: updateFuncInRange()
 * @author: Hyucksu Lee
 * @update date: 2022.01.05
 * @description
 * 해당 위치에 있는 구글 앱스 스크립트 함수를 강제로 갱신하게 끔 처리하는 함수 
 * cleanOrig 변수 추가함. 여러 개 셀이 업데이트 되다 보면 ? 문자가 여러개 붙어 있는 상태로 스크립트 실행 종료 되는 경우도 있었음.
 * 이를 방지하기 위해 updateFunctionRange 함수가 실행되면 무조건 ?가 있으면 제거하도록 함
 */

function updateFuncInRange(row, col){
  //console.log("Row info is " + row);
  //console.log("Column info is " + col);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var orig = sheet.getRange(row,col).getFormula();
  //
  var cleanOrig = removeQuestionMarkinString(orig); 
  //var temp = orig.replace("=", "?");
  var temp = cleanOrig.replace("=", "?");
  //sheet.getRange(row,col).setFormula(""); 
  sheet.getRange(row,col).setFormula(temp); 
  SpreadsheetApp.flush();
  sheet.getRange(row,col).setFormula(cleanOrig); 
  //SpreadsheetApp.flush();
}

function updateCountMacro(){
  SpreadsheetApp.flush();
}

/**
 * 
 * @function name: removeQuestionMarkinString()
 * @author: Hyucksu Lee
 * @update date: 2022.01.05
 * @param : inputStr
 * @description
 * 해당 컬럼의 셀 색을 카운트하는 함수 포함 문자열에 문자 '?'가 포함된 상태일 경우 문자 '?'를 제거하고 
 * 의도한 문자열 상태로 되돌리기 위한 문자열 처리 함수
 * @returns: outputStr
 */

function removeQuestionMarkinString(inputStr){
  //var inputStr = '=?????countOnBGColorOrange_2019("h7:h200")';
  var tempStr = ""
  var outputStr = ""
  var count = 0;
  var searchChar = '?';
  var addChar = '=';
  var pos = inputStr.indexOf(searchChar);
  
  console.log(inputStr);
  console.log(pos);

  while(pos !== -1){
    count++;
    pos = inputStr.indexOf(searchChar, pos + 1);
    console.log(`next \'?' position ${pos}!`);
  }
  
  //console.log(count);

  //console.log(inputStr.length);

  if(count > 0){
    console.log(`? 문자가 하나 이상 이상 남아 있던 경우`);
    //console.log(inputStr.substring(count+1, inputStr.length));
    tempStr = inputStr.substring(count+1, inputStr.length);
    outputStr = addChar.concat(tempStr);
  }else{
    console.log(`? 문자가 없는 경우`);
    outputStr = inputStr;
  }
    
  console.log(outputStr);  

  return outputStr;
}