function addRow() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("games");
  
  // シートの最終行を取得
  var lastRow = sheet.getLastRow();

  // コピーする行数
  var copyRow = lastRow + 1;
  
  var gameId = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
  sheet.getRange(copyRow, 1).setValue(gameId); // game_id
  sheet.getRange(lastRow, 2).copyTo(sheet.getRange(copyRow, 2)); // opponent_team_id
  sheet.getRange(lastRow, 3).copyTo(sheet.getRange(copyRow, 3)); // opponent_team_name
  sheet.getRange(lastRow, 4).copyTo(sheet.getRange(copyRow, 4)); // stadium_id
  sheet.getRange(lastRow, 5).copyTo(sheet.getRange(copyRow, 5)); // stadium_name
}

async function execute(gameId) {
  await insertGames(gameId);
}

async function insertGames(gameId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName("games");
  var row = findRow(sheet, Number(gameId), 1);
  const tableId = 'games';

  await deleteFirestoreDocument(tableId, 'game_id', gameId);

  var opponent_team_id = sheet.getRange(`B${row}`).getValue();
  var opponent_team_name = sheet.getRange(`C${row}`).getValue();
  var stadium_id = sheet.getRange(`D${row}`).getValue();
  var stadium_name = sheet.getRange(`E${row}`).getValue();
  var own_score = sheet.getRange(`F${row}`).getValue();
  var own_run = sheet.getRange(`G${row}`).getValue();
  var own_hit = sheet.getRange(`H${row}`).getValue();
  var own_error = sheet.getRange(`I${row}`).getValue();
  var opponent_score = sheet.getRange(`J${row}`).getValue();
  var opponent_run = sheet.getRange(`K${row}`).getValue();
  var opponent_hit = sheet.getRange(`L${row}`).getValue();
  var opponent_error = sheet.getRange(`M${row}`).getValue();
  var top = sheet.getRange(`N${row}`).getValue();
  var offense = sheet.getRange(`O${row}`).getValue();
  var game_datetime = sheet.getRange(`P${row}`).getValue();

  // win/lose/hold/save
  let query = `
    #StandardSQL \n
    SELECT
      players.player_id,
      players.last_name,
      players.first_name,
      players.player_no,
      players.photo_url,
      win,
      lose,
      hold,
      save
    FROM
      ${datasetId}.pitchings LEFT JOIN
      ${datasetId}.players ON pitchings.player_id = players.player_id
    WHERE
      game_id = ${gameId}
  `;

  let win = {};
  let lose = {};
  let save = {};

  let records = getQueryResult(query);
  if (records !== null) {
    records.shift();

    records.forEach((record) => {
      if (record[5] === '1') {
        // win
        win = {
          player_id: record[0],
          last_name: record[1],
          first_name: record[2],
          player_no: record[3],
          photo_url: record[4],
        }
      } else if (record[6] === '1') {
        // lose
        lose = {
          player_id: record[0],
          last_name: record[1],
          first_name: record[2],
          player_no: record[3],
          photo_url: record[4],
        }
      } else if (record[8] === '1') {
        // save
        save = {
          player_id: record[0],
          last_name: record[1],
          first_name: record[2],
          player_no: record[3],
          photo_url: record[4],
        }
      }
    });
  }

  // 2B/3B/HR
  query = `
    #StandardSQL \n
    SELECT
      players.player_id,
      players.last_name,
      players.first_name,
      players.player_no,
      players.photo_url,
      hit2,
      hit3,
      hr
    FROM
      ${datasetId}.hittings LEFT JOIN
      ${datasetId}.players ON hittings.player_id = players.player_id
    WHERE
      game_id = ${gameId} AND
      (hit2 >= 1 OR hit3 >= 1 OR hr >= 1)
  `;

  let hit2 = [];
  let hit3 = [];
  let hr = [];

  records = getQueryResult(query);
  if (records !== null) {
    records.shift()

    records.forEach((record) => {
      if (parseInt(record[5]) >= 1) { // 2B
        hit2.push(
          {
            player_id: record[0],
            last_name: record[1],
            first_name: record[2],
            player_no: record[3],
            photo_url: record[4],
            num: parseInt(record[5]),
          }
        )
      } else if (parseInt(record[6]) >= 1) { // 3B
        hit3.push(
          {
            player_id: record[0],
            last_name: record[1],
            first_name: record[2],
            player_no: record[3],
            photo_url: record[4],
            num: parseInt(record[6]),
          }
        )
      } else if (parseInt(record[7]) >= 1) { // HR
        hr.push(
          {
            player_id: record[0],
            last_name: record[1],
            first_name: record[2],
            player_no: record[3],
            photo_url: record[4],
            num: parseInt(record[7]),
          }
        )
      }
    });
  }

  let record = {
    'game_id': parseInt(gameId),
    'opponent_team_id': parseInt(opponent_team_id),
    'opponent_team_name': opponent_team_name,
    'stadium_id': parseInt(stadium_id),
    'stadium_name': stadium_name,
    'own_team_name': 'あかつき',
    'own_score': own_score,
    'own_run': parseInt(own_run),
    'own_hit': parseInt(own_hit),
    'own_error': parseInt(own_error),
    'opponent_score': opponent_score,
    'opponent_run': parseInt(opponent_run),
    'opponent_hit': parseInt(opponent_hit),
    'opponent_error': parseInt(opponent_error),
    'top': parseInt(top),
    'offense': offense,
    'game_datetime': game_datetime,
    'win': win,
    'lose': lose,
    'save': save,
    'hit2': hit2,
    'hit3': hit3,
    'hr': hr,
  };
  firestore.createDocument(tableId, record);

  query = `
    #StandardSQL \n
    delete from ${datasetId}.${tableId} where game_id = ${gameId};
    INSERT INTO ${datasetId}.${tableId} 
    (
      game_id,
      opponent_team_id,
      opponent_team_name,
      stadium_id,
      stadium_name,
      own_score,
      own_run,
      own_hit,
      own_error,
      opponent_score,
      opponent_run,
      opponent_hit,
      opponent_error,
      top,
      offense,
      game_datetime
    )
    VALUES (
      ${gameId}, 
      ${opponent_team_id}, 
      "${opponent_team_name}", 
      ${stadium_id}, 
      "${stadium_name}", 
      "${own_score}", 
      ${own_run}, 
      ${own_hit}, 
      ${own_error}, 
      "${opponent_score}", 
      ${opponent_run}, 
      ${opponent_hit}, 
      ${opponent_error}, 
      ${top}, 
      "${offense}", 
      "${game_datetime}"
    );
  `;
  var request = {
    query: query
  };
  var queryResults = BigQuery.Jobs.query(request, bqProjectId);
  queryResults.jobReference.jobId;
}

function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('データ登録', [
      {name: '行追加', functionName: 'addRow'},
      {name: '保存', functionName: 'save'},
    ]);
}
