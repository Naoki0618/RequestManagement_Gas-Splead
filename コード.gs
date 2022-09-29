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

function firestoreDate() {
  var dateArray = {
    'email': 'firebase-adminsdk-mto0u@sample-1b0c4.iam.gserviceaccount.com',
    'key': "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCghVD3X3nxPpSP\n+OztPV57+VNT5jt/hfd5aNvJaMAKFnl6LitBCgFxxgO4kXKZXMuHwZyRpuzje5RT\nGGUg2oUwUyb7EApSfHvvCkvVKMtuTX/Rjow52HFKVyxKxeKD/4sQMAKF+2p6neCe\nNq+GduLNmjNcFf7/v7SoH/L+Z9IAOW1RMpFszfx0GAaUiuQBT9SWSGVt6V5A+Fg4\nJZaR7F697Gv3f+QAlxrmmmEKC1wCfjRuzM/X8jqvHzMZoRZbxjdSG5EnMhBDM0rM\n7/TdHrcgo9ihHgSl9YhQls6FYD/zCoy7VvzOcppynXHqrfv/J27gl+PMAhD/aeK0\ndPgh2/0dAgMBAAECggEADIBVfe6BoLgu+cd5LEDLSvxv8OjNWXElhN8VvunZiu+V\nJl7SH46X7jRttcIeGrOPZlM9zlohuNW3B4Gu3pAmL01Ki+MD6sinHka/ASrcLQr8\nGWXwpdClghSn7mra6UzNl8UlbSnXcRU6mRfJM7+uijSoK1PLOD/F4hIa6pVLVZkD\nviT7rFsHtpZf20D5wONucgDsNNvIZZpxgWy5pI4oadzJbN7lUDDyI7ANq2TAgS4T\nulUb3OdnHMvdBTBDyEN8LZDn6V9R/3qWyYHpX1HHsLQ3EJesf6HfVdy3o9oCvKMn\nScj5Esj0YZORHWmrpriKphdl8p4R4RzWC2XO4C5oAQKBgQDUkCOQ91D1jufbr/ny\n1Gjmg17cLix9g+B0zw3/Oy8jtpcwxQbuASqYL1pTEb0Ugb0c+36fjKdaiyaAbo25\ngJMp5PbkN5UA/Y/p6/3AFQDtENQj1B6Eky4FiHWUksowkHUKPaqFSeeHXR0RAZ8C\nvoI+cDGy/tTVj7zRcpAXGjH1oQKBgQDBUq1n6lF8yGTPU0VBVSstMwtUy8alSJZT\n+Ab8Y/2VT0w9HsuCe3+nZhZBjw5Iz3iepwZ0aATbHWbznl8mqRRVmW1t0ym7EaV4\nv5tkEWNBgsZ9L8UQKIdcrQrhpKj9ZX+lJv1gdTGjAde3iDdyaVzlEmcxGgQUJ4PH\n0PLPEwkd/QKBgEXS81vzYczIHLG1pM13qN3P2aFKKaMxZtH4Egj9UAbTO+bxUc1s\n5KkJJQqUkR/jXlPe6UFP2smLXCJkLnn5Gl5wsAlXmMKyiEu3Eau/OoalOIpsa3nx\nPvTiVn1vmqtJSKkMiK8wD7YPiDTF643jNrV79VdvDkr45HWIxHxSRocBAoGBAKjT\nGrv01NSz69ViUsiLJ/mA6hRTIFaW3TDXGMKwT3NknJ+DlRWN5By7+hOmakMLa7qh\nAfIGJLd1JcL6Ov34Cdn28qlGDttevbKFIZ5x0MwU+GG6pc1Gl29Hbok+0pT3XlFL\ni1oA/ifsJAYS3tj7SjSBrbwjjAxNtbd5sZFEfmHBAoGAWEQy0x47R5C17yKuaD4J\nhPYetsBpoSftziSc4gAm/XVeirm9Iz35fCT3JnY5QcL1wFlZlJrQqCzucqJ1bYEk\nnlIRlGKLzr6ivWhCHXSmjW0F1lOW/46ogN3BYY1fbsKPU2d+xlT0jmtksc2qrvlE\ndVNHfS/OS9+4KDqJdsNwnzM=\n-----END PRIVATE KEY-----\n",
    'projectId': 'sample-1b0c4'
  }
  return dateArray;
}
// メイン処理部
function TEST(){
  // CloudFirestoreの認証
  var dateArray = firestoreDate();
  var firestore = FirestoreApp.getFirestore(dateArray.email, dateArray.key, dateArray.projectId);
  // CloudFirestoreからデータを読み込む
  const doc = firestore.getDocument("Sample/cVLqfmdKvVVPhiwjqcoP");
  console.log(doc);
}


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