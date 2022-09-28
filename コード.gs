const KYE_COL = 1
const REGISTRATION_COL = 2
const MAKER_COL = 3
const ITEMNAME_COL = 4
const ITEMCOUNT_COL = 5
const TANTO_COL = 6
const BIKOU_COL = 7
const STATUS_COL = 8
const RECEIVE_COL = 9
const SHIPPED_COL = 10

var HTTPS = "https://script.google.com/macros/s";
var ID = "AKfycbyvgkZS1Nk8LlPYT8jLZNjIX825uLYV8cymqL8bG_E";
var EXEC = "exec";
var URL = HTTPS + "/" + ID + "/" + EXEC;
var spreadsheetId = '1iozmhmbNYAbv9d8yfA3dxtHa1FOC_Kar7XFwlHx-K0E';

function doGet(e) {
  const page = (e.parameter.p || "index");
  let template = HtmlService.createTemplateFromFile(page);
  template.links = []; // こうしておくとテンプレートの方で links という変数に値が入った状態で使える
  return template
    .evaluate()
    .setTitle("サンプル申請")
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

function doPost(e) {

  // 履歴
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let values = sheet.getDataRange().getValues();

  // パラメータ
  let se = e.parameters['search']
  let st = e.parameters['status']

  // 結果
  let res;

  // 日付をフォーマット
  sheet.getRange("N1").setValue(se)
  sheet.getRange("N2").setValue(st)
  tmp = sheet.getRange("N1").getValue()
  if (tmp != "") {
    var dt = new Date(e.parameters['search']);
    se = Utilities.formatDate(dt, 'JST', 'yyyy/MM/dd')
  }else{
    se = ""
  }

  // 条件に合わせてフィルター
  res = values;
  if("" != se){
    res = res.filter(function (value) {
        return value[REGISTRATION_COL - 1] == se;
    })
  }
  if("" != st){
    res = res.filter(function (value) {
        return value[STATUS_COL - 1] == st;
    })
  }

  let template = HtmlService.createTemplateFromFile("index");
  template.links = res; // こうしておくとテンプレートの方で links という変数に値が入った状態で使える
  return template.evaluate();
}


function updateStatus(key) {
  var today = new Date();
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let values = sheet.getDataRange().getValues();

  let before_status = sheet.getRange(key + 1, STATUS_COL).getValue();

  if (before_status == "未") {
    sheet.getRange(key + 1, STATUS_COL).setValue("入庫済")
    sheet.getRange(key + 1, RECEIVE_COL).setValue(Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));
  } else if (before_status == "入庫済") {
    sheet.getRange(key + 1, STATUS_COL).setValue("出庫済")
    sheet.getRange(key + 1, SHIPPED_COL).setValue(Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));
  } else {
    sheet.getRange(key + 1, STATUS_COL).setValue("未")
  }
};

function updateRequest(li) {

  var today = new Date();
  let sheet = SpreadsheetApp.getActive().getActiveSheet();

  sheet.getRange("N1").setValue(li['email'])
  let lastRow = sheet.getLastRow() + 1;

  sheet.getRange(lastRow, KYE_COL).setValue(lastRow - 1);
  sheet.getRange(lastRow, REGISTRATION_COL).setValue(Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));
  sheet.getRange(lastRow, MAKER_COL).setValue(li['maker']);
  sheet.getRange(lastRow, ITEMNAME_COL).setValue(li['itemName']);
  sheet.getRange(lastRow, ITEMCOUNT_COL).setValue(li['itemCount']);
  sheet.getRange(lastRow, TANTO_COL).setValue(li['tanto']);
  sheet.getRange(lastRow, BIKOU_COL).setValue(li['bikou']);
  sheet.getRange(lastRow, STATUS_COL).setValue("未");
};