function onOpen() {
  //set up custom menu
  var ui = SpreadsheetApp.getUi();  
  ui.createMenu('🦠 😷 콤프레샤 현장진단')
    .addItem('첫째 데이터추출','extractWsys')
    .addItem('둘째 차트그리기','drawChart')
//    .addItem('세째 범위선정','extractWsys')
    .addItem('셋째 변곡점계산','doInflection')
    .addItem('넷째 플마고저','doPMUP')
    .addItem('다섯째 로딩률계산','doLoading')
    .addItem('여섯째 로딩률그래프','drawLoading')
    .addItem('일곱째 간이산출','makeReport')
//    .addItem('여덜째 초기화','drawLoading')
    .addToUi();
};

function makeReport() {
//  https://docs.google.com/spreadsheets/d/1eQuLp2Yglk3LfAOlE-MqP1aUHkWetwFuy2tQwHIu44g/edit?usp=sharing
  var templateSheet = SpreadsheetApp.openById("1eQuLp2Yglk3LfAOlE-MqP1aUHkWetwFuy2tQwHIu44g")
  
//  Logger.log(templateSheet.getSheetName())
  var currentApp = SpreadsheetApp.getActiveSpreadsheet();
  templateSheet.getSheetByName("template").copyTo(currentApp)
//  currentApp.getName()
//  Browser.msgBox(currentApp.getActiveSheet().getName())
//  var reportSheet = currentApp.getSheetByName("template의 사본")
  var reportSheet = currentApp.getSheetByName("Copy of template")
//  reportSheet.setName("template")
  
  var dataSheet = currentApp.getSheetByName("계산결과")

  var dim2data = dataSheet.getRange(1, 8,4,4).getValues()
  
  
  var range = reportSheet.getRange(5, 6)
  range.setValue(dim2data[1][3])
  range = reportSheet.getRange(5, 8).setValue(dim2data[1][2])
  range.setValue(dim2data[1][2])
  range = reportSheet.getRange(5, 9).setValue(dim2data[3][2])
  range.setValue(dim2data[3][2])
  
  dataSheet.activate().showSheet()
  
}

function drawLoading() {
  var html = HtmlService.createHtmlOutputFromFile("drawLoadingD3Html").setHeight(700).setWidth(700);
  SpreadsheetApp.getUi().showModalDialog(html, "로딩률 그래프");
}

function getSheetDataOfLoading() {
  var chkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("계산결과");
//  var resultSheet = null;
  if (chkSheet === null) {
    Browser.msgBox("계산결과 시트가 없습니다");
    
    return null;
  }
  
  
  var dim2data = chkSheet.getRange(1, 13,5,4).getValues()
  
  
  return dim2data;
  
}

function doLoading() {
  
  var chkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("계산결과");
//  var resultSheet = null;
  if (chkSheet != null) {
    Browser.msgBox("계산결과 시트 이름이 중복됩니다");
    resultSheet = chkSheet;
    return;
  }
  
  var resultSheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = resultSheet.insertSheet("계산결과");  
  
  var edSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDChart");
  var idxRange = edSheet.getRange(2,11,1,2)  // K 열
  var idxVal = idxRange.getValues()
  
  
  var cell = activeSheet.getRange(1,1,1,6).merge()
  cell.setValue("합계")
  cell.setHorizontalAlignment("center")
  var cell = activeSheet.getRange(2,2)
  cell.setValue("플고")
  var cell = activeSheet.getRange(2,3)
  cell.setValue("플저")
  var cell = activeSheet.getRange(2,4)
  cell.setValue("마고")
  var resultRange = activeSheet.getRange(2,5)
  resultRange.setValue("마저")
  var resultRange = activeSheet.getRange(2,6)
  resultRange.setValue("총계")
  var resultRange = activeSheet.getRange(3,1)
  resultRange.setValue("합계")
  var resultRange = activeSheet.getRange(4,1)
  resultRange.setValue("비율")
  
  var targetRange = activeSheet.getRange(3, 2)   // 플고
  targetRange.setFormula(`=( SUM( EDChart!G${idxVal[0][0]}:G${idxVal[0][1]}))`);
  var targetRange = activeSheet.getRange(3, 3)   // 플저
  targetRange.setFormula(`=( SUM( EDChart!H${idxVal[0][0]}:H${idxVal[0][1]}))`);
  var targetRange = activeSheet.getRange(3, 4)   // 마고
  targetRange.setFormula(`=( SUM( EDChart!I${idxVal[0][0]}:I${idxVal[0][1]}))`);
  var targetRange = activeSheet.getRange(3, 5)   // 마저
  targetRange.setFormula(`=( SUM( EDChart!J${idxVal[0][0]}:J${idxVal[0][1]}))`);
  var targetRange = activeSheet.getRange(3, 6)   // 총계
  targetRange.setFormula(`=( SUM(B3:E3))`);
  
  var targetRange = activeSheet.getRange(4, 2)   // 비율
  targetRange.setFormula(`=(( B3 / F3) * 100 )`);
  var targetRange = activeSheet.getRange(4, 3)   // 비율
  targetRange.setFormula(`=(( C3 / F3) * 100 )`);
  var targetRange = activeSheet.getRange(4, 4)   // 비율
  targetRange.setFormula(`=(( D3 / F3) * 100 )`);
  var targetRange = activeSheet.getRange(4, 5)   // 비율
  targetRange.setFormula(`=(( E3 / F3) * 100 )`);
  var targetRange = activeSheet.getRange(4, 6)   // 총계
  targetRange.setFormula(`=( SUM(B4:E4))`);
  
  
  var cell = activeSheet.getRange(1,8)
  cell.setValue("평균전력")
  var cell = activeSheet.getRange(1,9)
  cell.setValue("최소전력")
  var cell = activeSheet.getRange(1,10)
  cell.setValue("로딩전력")
  var cell = activeSheet.getRange(1,11)
  cell.setValue("로딩률")
  var cell = activeSheet.getRange(3,10)
  cell.setValue("언로딩전력")
  var cell = activeSheet.getRange(3,11)
  cell.setValue("언로딩률")
  
  var targetRange = activeSheet.getRange(2, 8)   
  targetRange.setFormula(`=( Average( EDChart!B${idxVal[0][0]}:B${idxVal[0][1]}))`);
  var targetRange = activeSheet.getRange(2, 9)   
  targetRange.setFormula(`=( Min( EDChart!B${idxVal[0][0]}:B${idxVal[0][1]}))`);
  var targetRange = activeSheet.getRange(2, 10)   
  targetRange.setFormula(`=( SUM(B3:C3))`);
  var targetRange = activeSheet.getRange(2, 11)   
  targetRange.setFormula(`=( SUM(B4:C4))`);
  var targetRange = activeSheet.getRange(4, 10)   
  targetRange.setFormula(`=( SUM(D3:E3))`);
  var targetRange = activeSheet.getRange(4, 11)   
  targetRange.setFormula(`=( SUM(D4:E4))`);
  
  
  var vPHr = activeSheet.getRange(4, 2).getValue()
  var vPLr = activeSheet.getRange(4, 3).getValue()
  var vMHr = activeSheet.getRange(4, 4).getValue()
  var vMLr = activeSheet.getRange(4, 5).getValue()
  
  var offsetX = 10
  var offsetY = 70
  var totalWidth = 500
  var totalHieght = 300
  
  var vXwi = totalWidth * (vPHr / 100)
  var vYhi = totalHieght * 0.7
  var vX = offsetX
  var vY = offsetY + totalHieght
//  With shtX.Shapes.AddLine(vX, vY, vXwi, vYhi)
  activeSheet.getRange(1, 13).setValue("first")
  activeSheet.getRange(2, 13).setValue(vX)
  activeSheet.getRange(3, 13).setValue(vY)
  activeSheet.getRange(4, 13).setValue(vXwi)
  activeSheet.getRange(5, 13).setValue(vYhi)
  
  vX = vXwi
  vY = vYhi
  vXwi = vXwi + totalWidth * (vPLr / 100)
  vYhi = vYhi - totalHieght * 0.05
//  With shtX.Shapes.AddLine(vX, vY, vXwi, vYhi).Line
  activeSheet.getRange(1, 14).setValue("second")
  activeSheet.getRange(2, 14).setValue(vX)
  activeSheet.getRange(3, 14).setValue(vY)
  activeSheet.getRange(4, 14).setValue(vXwi)
  activeSheet.getRange(5, 14).setValue(vYhi)
  
  vX = vXwi
  vY = vYhi
  vXwi = vXwi + totalWidth * (vMHr / 100)
  vYhi = vYhi + totalHieght * 0.4
//  With shtX.Shapes.AddLine(vX, vY, vXwi, vYhi).Line
  activeSheet.getRange(1, 15).setValue("third")
  activeSheet.getRange(2, 15).setValue(vX)
  activeSheet.getRange(3, 15).setValue(vY)
  activeSheet.getRange(4, 15).setValue(vXwi)
  activeSheet.getRange(5, 15).setValue(vYhi)
  
  vX = vXwi
  vY = vYhi
  vXwi = vXwi + totalWidth * (vMLr / 100)
  vYhi = vYhi + totalHieght * 0.1
//  With shtX.Shapes.AddLine(vX, vY, vXwi, vYhi).Line
  activeSheet.getRange(1, 16).setValue("forth")
  activeSheet.getRange(2, 16).setValue(vX)
  activeSheet.getRange(3, 16).setValue(vY)
  activeSheet.getRange(4, 16).setValue(vXwi)
  activeSheet.getRange(5, 16).setValue(vYhi)  
  
  activeSheet.getRange(2, 12).setValue("x1")
  activeSheet.getRange(3, 12).setValue("y1")
  activeSheet.getRange(4, 12).setValue("x2")
  activeSheet.getRange(5, 12).setValue("y2")
  
}

function doPMUP() {
  var edSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDChart");
  
  var targetRange = edSheet.getRange(1,4)  // 변곡점
  var chkTargetVal = targetRange.getValue()
  
  if (chkTargetVal !== "변곡점" ){
       Browser.msgBox("변곡점 계산을 먼저 하세요")
       return;
  }
  
  var idxRange = edSheet.getRange(2,11,1,2)  // K 열, L 열  시작점,종료점
  var idxVal = idxRange.getValues()
  
  var targetRange = edSheet.getRange(1,5)
  targetRange.setValue("마")  // 원래는 '플마'
  var targetRange = edSheet.getRange(1,6)  
  targetRange.setValue("고저")
  
//  이렇게하면 겁나 느림  
//  for (var i = idxVal[0][0]; i < idxVal[0][1]; i++) {
//    if (edSheet.getRange(i,4).getValue() === "변곡점"){
//      if (edSheet.getRange(i-1,5).getValue() === "마"){
//          edSheet.getRange(i,5).setValue("플")
//      }else{
//          edSheet.getRange(i,5).setValue("마")
//      }
//    } else{
//          edSheet.getRange(i,5).setValue(edSheet.getRange(i-1,5).getValue())
//    }
//  }
  
  // 기본값 '플마' '마' 채우기
  var defaultRange = edSheet.getRange(2,5,idxVal[0][0])
  defaultRange.setValue("마")
  
  var calcRange = edSheet.getRange(idxVal[0][0], 5, idxVal[0][1], 1)   // 5 열, E
  calcRange.setFormula(`=(IF ( D${idxVal[0][0]} = "변곡점",  IF (E${idxVal[0][0]-1} = "마", "플", "마")   , E${idxVal[0][0]-1} ))`);
  
  var targetRangeUp = edSheet.getRange(6, 11, 1, 1)  // K
  var targetRangeDn = edSheet.getRange(6, 12, 1, 1)  // L
  var vIndexUp =  targetRangeUp.getValue()
  var vIndexDn =  targetRangeDn.getValue()
  
  var vTarget = (vIndexUp - vIndexDn) * 0.5;
  
  var calcRange = edSheet.getRange(idxVal[0][0], 6, idxVal[0][1], 1)   // 6 열, F   
  calcRange.setFormula(`=(IF (ABS(C${idxVal[0][0]}) > ${vTarget},"고","저"))`);
  
  var targetRange = edSheet.getRange(1,7)
  targetRange.setValue("플고")
  var targetRange = edSheet.getRange(1,8)  
  targetRange.setValue("플저")
  var targetRange = edSheet.getRange(1,9)
  targetRange.setValue("마고")
  var targetRange = edSheet.getRange(1,10)  
  targetRange.setValue("마저")
  
  var calcRange = edSheet.getRange(idxVal[0][0], 7, idxVal[0][1], 1)   // 7 열, G
  calcRange.setFormula(`=(IF ( AND( E${idxVal[0][0]} = "플", F${idxVal[0][0]} = "고"), B${idxVal[0][0]},""))`);
  
  var calcRange = edSheet.getRange(idxVal[0][0], 8, idxVal[0][1], 1)   // 8 열, H
  calcRange.setFormula(`=(IF ( AND( E${idxVal[0][0]} = "플", F${idxVal[0][0]} = "저"), B${idxVal[0][0]},""))`);
  
  var calcRange = edSheet.getRange(idxVal[0][0], 9, idxVal[0][1], 1)   // 9 열, F
  calcRange.setFormula(`=(IF ( AND( E${idxVal[0][0]} = "마", F${idxVal[0][0]} = "고"), B${idxVal[0][0]},""))`);
  
  var calcRange = edSheet.getRange(idxVal[0][0], 10, idxVal[0][1], 1)   // 10 열, J
  calcRange.setFormula(`=(IF ( AND( E${idxVal[0][0]} = "마", F${idxVal[0][0]} = "저"), B${idxVal[0][0]},""))`);
  
}

function doInflection() {

  var edSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDChart");
  var chkRange = edSheet.getRange(2,11,1,2)  // K 열
  var chkVal = chkRange.getValues()
  var lastRow = edSheet.getLastRow();  // total row with Data
  
  if (chkVal[0][0] === -1 || chkVal[0][0] === undefined || chkVal[0][0] === ""){
       Browser.msgBox("시작점이 지정되지 않았습니다")
       return;
  }else if (chkVal[0][1] === -1 || chkVal[0][1] === undefined || chkVal[0][1] === ""){
       Browser.msgBox("종료점이 지정되지 않았습니다")
       return;
  }else if (chkVal[0][0] >= chkVal[0][1]){
       Browser.msgBox("시작점이 종료값보다 커야합니다")
       return;
  }
  
  var targetRange = edSheet.getRange(1,3)  // K 열
  targetRange.setValue("차이")
  var difRange = edSheet.getRange(3, 3, lastRow, 1)
  difRange.setFormula("=(B3 - B2)");
  
// MsgBox "상위2% 평균 : " & indexMax & "하위10% 평균 : " & indexMin
  
  var targetRange = edSheet.getRange(3,11)  // K 열
  targetRange.setValue("상위백분위수")
  var targetRange = edSheet.getRange(4, 11, 1, 1)
  targetRange.setFormula("=PERCENTILE.EXC(B2:B2510,0.98)");
  
  var targetRange = edSheet.getRange(5,11)  // K 열
  targetRange.setValue("상위2%평균")
  var targetRangeUp = edSheet.getRange(6, 11, 1, 1)
  targetRangeUp.setFormula(`=AVERAGEIFS(B2:B${lastRow},B2:B${lastRow},">=" & K4)`);
  
  var targetRange = edSheet.getRange(3,12)  // L 열
  targetRange.setValue("하위백분위수")
  var targetRange = edSheet.getRange(4, 12, 1, 1)
  targetRange.setFormula("=PERCENTILE.EXC(B2:B2510,0.10)");
  
  var targetRange = edSheet.getRange(5,12)  
  targetRange.setValue("하위10%평균")
  var targetRangeDn = edSheet.getRange(6, 12, 1, 1)
  targetRangeDn.setFormula(`=AVERAGEIFS(B2:B${lastRow},B2:B${lastRow},">=" & L4)`);
  

  var vIndexUp =  targetRangeUp.getValue()
  var vIndexDn =  targetRangeDn.getValue()
  
//  Browser.msgBox(vIndexUp  + " / " + vIndexDn)
  
  
  var vTarget = (vIndexUp - vIndexDn) * 0.5;
  
  var targetRange = edSheet.getRange(1,4)  
  targetRange.setValue("변곡점")
  var targetRange = edSheet.getRange(2,4,lastRow,1)  
  targetRange.setValue("")

  // 시작점에 무조건 플러스 1  해줌, 
  var difRange = edSheet.getRange(chkVal[0][0]+1, 4, chkVal[0][1]+1, 1)
  difRange.setFormula(`=(IF (AND(ABS(C${chkVal[0][0]+1}) > ${vTarget}, ABS(C${chkVal[0][0]+1-1}) < ${vTarget}),"변곡점"))`);
  
}

function drawChart() { 
  var html = HtmlService.createHtmlOutputFromFile("ChartHtml").setHeight(700).setWidth(700);
  SpreadsheetApp.getUi().showModalDialog(html, "범위선택 차트");
}

function sendPointToSheet(point1, point2) {
  // 차트는 0 번부터 시작하므로 시트 행과 맞추기 위해서 +2 해준다
  var edSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDChart");
  var targetRange = edSheet.getRange(1,11)  // K 열
  targetRange.setValue("시작점")
  var targetRange = edSheet.getRange(2,11)
  targetRange.setValue(point1+2)
  var targetRange = edSheet.getRange(1,12)  // L 열
  targetRange.setValue("종료점")
  var targetRange = edSheet.getRange(2,12)
  targetRange.setValue(point2+2)
  return point1 + " / " + point2;
}

function extractWsys() {
  var rawSheet = SpreadsheetApp.getActiveSheet();
  
//  Browser.msgBox(rawSheet.getLastColumn() + " / " + rawSheet.getLastRow());
//  Browser.msgBox(rawSheet.getMaxColumns() + " / " + rawSheet.getMaxRows());
//  return;
  
  var tf = rawSheet.createTextFinder("TIME");
  var all = tf.findAll();
  // Browser.msgBox(all[0].getRow() + " / " + all[0].getColumn());
  var timeRow = all[0].getRow()
  var timeCol = all[0].getColumn()
  
  var tf2 = rawSheet.createTextFinder("W_SYS");
  var all2 = tf2.findAll();
  // Browser.msgBox(all2[0].getRow() + " / " + all2[0].getColumn());
  var wsysRow = all2[0].getRow()
  var wsysCol = all2[0].getColumn()
  
  var lastRow = rawSheet.getLastRow();
  var lastCol = rawSheet.getLastColumn();
  
//  Browser.msgBox(lastRow + " / " + lastCol);
  
  var edSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EDChart");
//  var newss = null;
  if (edSheet != null) {
    Browser.msgBox("EDChart 이름이 중복됩니다");
//    newss = chkEDSheet;
    return;
  }
  
  edSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("EDChart");
//  rawSheet.i.insertSheet("EDChart");

  

  var timeRange = rawSheet.getRange(timeRow,timeCol,lastRow)
  var rawTimeValue = timeRange.getValues()
  
//  edSheet.getRange(1, 1, lastRow).setValues(rawTimeValue)
  var newRange = edSheet.getRange(1, 1, lastRow) // .getRange(2 , 1, lastRow, 1)  //  첫번재 셀에는 텍스트 
  newRange.setNumberFormat("HH:MM:SS")
  var newTimeValue = newRange.setValues(rawTimeValue)

  var rawSysRange = rawSheet.getRange(wsysRow,wsysCol,lastRow)
  var rawWsysValue = rawSysRange.getValues()
  
  var newRange = edSheet.getRange(1, 2, lastRow, 1)
  var newSysValue = newRange.setValues(rawWsysValue)
 
  
  var newRange = edSheet.getRange(2, 3, lastRow, 1)
  newRange.setFormula("=LEFT(B2,5)");
 
  
  var rawSheet = SpreadsheetApp.getActiveSheet();
  var lastRow = rawSheet.getLastRow();
  var rawRange = rawSheet.getRange(2,3,lastRow)
  var rawWsysValue = rawRange.getValues()
  
  var copyRange = rawSheet.getRange(2, 2, lastRow)
  copyRange.setValues(rawWsysValue)
  var rawRange = rawSheet.getRange(2,3)
  rawRange.setValue(0)
}
