function addRow() {
    var sheet = SpreadsheetApp.getActive().getSheetByName('players');
    // シートの最終行を取得
    var lastRow = sheet.getLastRow();
    // コピーする行数
    var copyRow = lastRow + 1;
    var playerId = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
    sheet.getRange(copyRow, 1).setValue(playerId); // player_id
}

async function execute(playerId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('players');
  var row = findRow(sheet, playerId, 1);
  var tableId = 'players';

  await deleteFirestoreDocument(tableId, 'player_id', playerId);

  var playerNo = sheet.getRange("B" + row).isBlank() ? null : sheet.getRange("B" + row).getValue();
  var lastName = sheet.getRange("C" + row).getValue();
  var firstName = sheet.getRange("D" + row).getValue();
  var member = sheet.getRange("E" + row).getValue();
  var pitcher = sheet.getRange("F" + row).getValue();
  var catcher = sheet.getRange("G" + row).getValue();
  var first = sheet.getRange("H" + row).getValue();
  var second = sheet.getRange("I" + row).getValue();
  var third = sheet.getRange("J" + row).getValue();
  var shortstop = sheet.getRange("K" + row).getValue();
  var outfielder = sheet.getRange("L" + row).getValue();
  var photoUrl = sheet.getRange("M" + row).getValue();

  let record = {
    'player_id': parseInt(playerId),
    'player_no': parseInt(playerNo),
    'last_name': lastName,
    'first_name': firstName,
    'member': parseInt(member),
    'pitcher': parseInt(pitcher),
    'catcher': parseInt(catcher),
    'first': parseInt(first),
    'second': parseInt(second),
    'third': parseInt(third),
    'shortstop': parseInt(shortstop),
    'outfielder': parseInt(outfielder),
    'photo_url': photoUrl
  };
  firestore.createDocument(tableId, record);

  var query = `
    #StandardSQL \n
    delete from ${datasetId}.${tableId} where player_id = ${playerId};
    INSERT INTO ${datasetId}.${tableId}
    (
      player_id,
      player_no,
      last_name,
      first_name,
      member,
      pitcher,
      catcher,
      first,
      second,
      third,
      shortstop,
      outfielder,
      photo_url
    ) values (
      ${playerId},
      ${playerNo},
      "${lastName}",
      "${firstName}",
      ${member},
      ${pitcher},
      ${catcher},
      ${first},
      ${second},
      ${third},
      ${shortstop},
      ${outfielder},
      "${photoUrl}"
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
    { name: '行追加', functionName: 'addRow' },
    { name: '保存', functionName: 'save' },
  ]);
}
