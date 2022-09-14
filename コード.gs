const SEARCH_KYE_COL = 1
const STATUS_COL     = 8
const RECEIVE_COL    = 9

var HTTPS = "https://script.google.com/macros/s";
var ID = "AKfycbyvgkZS1Nk8LlPYT8jLZNjIX825uLYV8cymqL8bG_E";
var EXEC = "exec";
var URL = HTTPS + "/" + ID + "/" + EXEC;
var spreadsheetId = '1iozmhmbNYAbv9d8yfA3dxtHa1FOC_Kar7XFwlHx-K0E';

function doGet() {
  let template = HtmlService.createTemplateFromFile("index");
  template.links = []; // こうしておくとテンプレートの方で links という変数に値が入った状態で使える
  return template.evaluate();
}

function updateStatus(key) {
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let values = sheet.getDataRange().getValues();

  let before_status = sheet.getRange(key+1, STATUS_COL).getValue();

  if(before_status == "未"){
    sheet.getRange(key+1, STATUS_COL).setValue("入庫済")
  }else if(before_status == "入庫済"){
    sheet.getRange(key+1, STATUS_COL).setValue("出庫済")
  }else{

  }

  var today = new Date();
  sheet.getRange(key+1, RECEIVE_COL).setValue(Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));

};

function doPost(e) {

  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let values = sheet.getDataRange().getValues();

  // sheet.getRange("J1").setValue("リンゴ")
  let res = [];
  let li = [];
  for (i = 1; i < values.length; i++) {
    li = values[i];
    if (li[SEARCH_KYE_COL] == e.parameters['search']) {
      res.push(li);
    }
  };

  let template = HtmlService.createTemplateFromFile("index");
  template.links = res; // こうしておくとテンプレートの方で links という変数に値が入った状態で使える
  return template.evaluate();
}
