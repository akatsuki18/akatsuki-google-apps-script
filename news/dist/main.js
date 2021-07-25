function addRow() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("news");
  
  // シートの最終行を取得
  var lastRow = sheet.getLastRow();

  // コピーする行数
  var copyRow = lastRow + 1;
  
  var newsId = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
  sheet.getRange(copyRow, 1).setValue(newsId); // news_id
}

function save() {
  var newsId = Browser.inputBox("news_id を入力してください", Browser.Buttons.OK_CANCEL);
  var sheet = SpreadsheetApp.getActive().getSheetByName("news");
  var row = findRow(sheet, Number(newsId), 1);
  var datasetId = 'goldpony';
  const tableId = 'news';

  var date = sheet.getRange(`B${row}`).getValue();
  var title = sheet.getRange(`C${row}`).getValue();
  var body = sheet.getRange(`D${row}`).getValue();
  var photoUrl = sheet.getRange(`E${row}`).getValue();

  var email = PropertiesService.getScriptProperties().getProperty('serviceAccountEmail');
  var keyString = PropertiesService.getScriptProperties().getProperty('serviceAccountKey');
  var key = keyString.replace(/\\n/g, "\n");
  var projectId = PropertiesService.getScriptProperties().getProperty('projectId');
  var firestore = FirestoreApp.getFirestore(email, key, projectId);
  // 登録データ
  var data = {
    'news_id': parseInt(newsId),
    'date': date,
    'title': title,
    'body': body,
    'photo_url': photoUrl
  };
  var news = firestore.query('news').Where('news_id', '==', parseInt(newsId)).Execute();
  if (news.length > 0) {
      var index = news[0].name.lastIndexOf('/');
      var documentId = news[0].name.substring(index + 1);
      firestore.deleteDocument("news/" + documentId);
  }
  firestore.createDocument('news', data);
}

function findRow(sheet,val,col){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(var i=1;i<dat.length;i++){
    if(dat[i][col-1] === val){
      return i+1;
    }
  }
  return 0;
}

function convCsv(range) {
  try {
    var data = range.getValues();
    var ret = "";
    if (data.length > 1) {
      var csv = "";
      for (var i = 0; i < data.length; i++) {
        for (var j = 0; j < data[i].length; j++) {
          if (data[i][j].toString().indexOf(",") != -1) {
            data[i][j] = "\"" + data[i][j] + "\"";
          }
        }
        if (i < data.length-1) {
          csv += data[i].join(",") + "\r\n";
        } else {
          csv += data[i];
        }
      }
      ret = csv;
    }
    return ret;
  }
  catch(e) {
    Logger.log(e);
  }
}

function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('データ登録', [
      {name: '行追加', functionName: 'addRow'},
      {name: '保存', functionName: 'save'},
    ]);
}
