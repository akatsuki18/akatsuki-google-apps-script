const datasetId = 'goldpony';
const bqProjectId = 'nifty-bindery-293409';
const tempTableName = 'temp_hittings';
const email = PropertiesService.getScriptProperties().getProperty('serviceAccountEmail');
const keyString = PropertiesService.getScriptProperties().getProperty('serviceAccountKey');
const key = keyString.replace(/\\n/g, "\n");
const projectId = PropertiesService.getScriptProperties().getProperty('projectId');
const firestore = FirestoreApp.getFirestore(email, key, projectId);

function convCsv(range) {
  try {
    let data = range.getValues();
    let ret = '';
    if (data.length > 1) {
      let csv = '';
      for (let i = 0; i < data.length; i++) {
        for (let j = 0; j < data[i].length; j++) {
          if (data[i][j].toString().indexOf(',') != -1) {
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

async function deleteFirestoreDocument(collectionName, idName, id) {
  let documents = firestore.query(collectionName).Where(idName, '==', parseInt(id)).Execute();
  if (documents.length > 0) {
    documents.forEach(function (document) {
      let index = document.name.lastIndexOf('/');
      let documentId = document.name.substring(index + 1);
      firestore.deleteDocument(`${collectionName}/${documentId}`);
    })
  }
}

function findRows(sheet, val, col){
  let rows = [];
  let dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(let i = 1; i < dat.length; i++) {
    if(dat[i][col-1] === val){
      rows.push(i+1);
    }
  }

  return rows;
}

async function save() {
  let gameId = Browser.inputBox('game_id を入力してください', Browser.Buttons.OK_CANCEL);

  await execute(gameId);
}
