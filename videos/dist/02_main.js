function addRow() {
  let sheet = SpreadsheetApp.getActive().getSheetByName("videos");
  
  // シートの最終行を取得
  let lastRow = sheet.getLastRow();

  // コピーする行数
  let copyRow = lastRow + 1;
  
  let gameId = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
  sheet.getRange(copyRow, 1).setValue(gameId); // game_id
  let lastTurn = sheet.getRange(lastRow, 1).getValue();
  let turn = 1;
  if (lastTurn) {
    turn = Number(sheet.getRange(lastRow, 2).getValue()) + 1;
  }
  sheet.getRange(copyRow, 2).setValue(turn); // turn
}

async function execute(gameId) {
  insert(gameId);
}

async function insert(gameId) {
  let tableId = 'videos';
  let sheet = SpreadsheetApp.getActive().getSheetByName('videos');
  
  await deleteFirestoreDocument(tableId, 'game_id', gameId);

  let turn = videoUrl = '';
  let values = [];
  let firestoreDocument = {};
  let rows = findRows(sheet, Number(gameId), 1);
  rows.forEach(function (row) {
    turn = sheet.getRange(`B${Number(row)}`).getValue();
    videoUrl = sheet.getRange(`C${Number(row)}`).getValue();

    values.push(`(${gameId}, ${turn}, "${videoUrl}")`);

    firestoreDocument = {
      game_id: parseInt(gameId),
      turn: parseInt(turn),
      video_url: videoUrl,
    };
    firestore.createDocument(tableId, firestoreDocument);
  });

  let query = `
    #StandardSQL \n
    DELETE FROM ${datasetId}.${tableId};
    INSERT INTO ${datasetId}.${tableId} (game_id, turn, video_url)
    VALUES ${values.join()};
  `;

  let request = { query: query };
  BigQuery.Jobs.query(request, bqProjectId);
}

function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('データ登録', [
      {name: '行追加', functionName: 'addRow'},
      {name: '保存', functionName: 'save'},
    ]);
}
