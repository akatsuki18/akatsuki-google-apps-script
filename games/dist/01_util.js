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
  let query = firestore.query(collectionName);
  if (idName !== null) query = query.Where(idName, '==', parseInt(id));

  let documents = query.Execute();
  if (documents.length > 0) {
    documents.forEach(function (document) {
      let index = document.name.lastIndexOf('/');
      let documentId = document.name.substring(index + 1);
      firestore.deleteDocument(`${collectionName}/${documentId}`);
    })
  }
}

function findRow(sheet, val, col) {
  let row = ''
  let dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(let i = 1; i < dat.length; i++){
    if(dat[i][col-1] === val){
      row = i+1;
    }
  }
  return row;
}

function findRows(sheet, val, col) {
  let rows = [];
  let dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(let i = 1; i < dat.length; i++) {
    if(dat[i][col-1] === val){
      rows.push(i+1);
    }
  }

  return rows;
}

function getQueryResult(query) {
  const request = { query: query };
  const queryResults = BigQuery.Jobs.query(request, bqProjectId);
  const headers = queryResults.schema.fields.map(({name}) => name);
  const rows = queryResults.rows;

  if (rows === undefined) return null;

  let records = rows.map(({ f }) => f.map(({ v }) => v));
  records.unshift(headers);

  return records;
}

async function save() {
  let gameId = Browser.inputBox('game_id を入力してください', Browser.Buttons.OK_CANCEL);

  await execute(gameId);
}
