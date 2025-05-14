//url :https://script.google.com/macros/s/AKfycbyo4Vt0wOtbiMEgRG_BJZaEVMCi2IJyekx59pM8NAw/dev
//gitok
// Presented by BrilliantPy
/*######################### Editable Start #########################*/
function getSheetData(){
var data = SpreadsheetApp.openById('1nKhRFwFymDotxgMIKr2yeCTs0aoA7DQW0KpscS9VSsQ');
var sheet = data.getSheetByName('product','ประเภท','ตัวเลือกสินค้า','ส่วนลด','ออเดอร์','sales','expenses','สั่งของ','สินค้า');
var datarange = sheet.getDataRange();
var value = datarange.getValues();
return value;
}


/*#########################  Editable End  #########################*/
// Init
// let ss,sheet,lastRow,lastCol,range,values;

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function addDataValue(name,department) {
  initSpreadSheet();
  sheet.appendRow([new Date(),getCustomUUID(),name,department]);
  return 'SUCCESS';
}

function readDataValue() {
  initSpreadSheet();
  let dataArr = sheet.getDataRange().getValues();
  dataArr.shift();
  dataArr.reverse();
  console.log(dataArr,typeof(dataArr));
  dataArr = JSON.stringify(dataArr);
  console.log(dataArr,typeof(dataArr));
  return dataArr;
}

function getDataValue(dataId) {
  initSpreadSheet();
  let dataArr = [];
  for(let i=2;i<=lastRow;i++) {
    let curId = sheet.getRange(i,idColNum).getValue();
    if(curId == dataId) {
      dataArr = sheet.getRange(i,1,1,lastCol).getValues();
      dataArr = dataArr[0];
      break;
    }
  }
  console.log(dataArr,typeof(dataArr));
  dataArr = JSON.stringify(dataArr);
  console.log(dataArr,typeof(dataArr));
  return dataArr;
}

function updateDataValue(dataId,name,department) {
  initSpreadSheet();
  for(let i=2;i<=lastRow;i++) {
    let curId = sheet.getRange(i,idColNum).getValue();
    if(curId == dataId) {
      sheet.getRange(i,nameColNum).setValue(name);
      sheet.getRange(i,departmentColNum).setValue(department);
      break;
    }
  }
  return "SUCCESS";
}

function removeDataValue(dataId) {
  initSpreadSheet();
  for(let i=2;i<=lastRow;i++) {
    let curId = sheet.getRange(i,idColNum).getValue();
    if(curId == dataId) {
      sheet.deleteRow(i);
      break;
    }
  }
  return 'SUCCESS';
}

function testFunc() {
  // addDataValue('PPP',"Hr");
  // readDataValue();
  // getDataValue('df40098713');
  // updateDataValue('df40098713','PPP2','Dev');
  removeDataValue('df40098713');
}

function initSpreadSheet() {
  ss = SpreadsheetApp.getActive();
  sheet = ss.getSheetByName(sheetName);
  lastRow = sheet.getLastRow();
  lastCol = sheet.getLastColumn();
  // console.log("lastRow:",lastRow,",lastCol:",lastCol);
  console.log('initSpreadSheet completed');
}

function getCustomUUID() {
  let rawId = Utilities.getUuid().split("-").join("");
  rawId = rawId.slice(0,10);
  return rawId;
}