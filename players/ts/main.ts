function addRow() {
  const sheet: any = SpreadsheetApp.getActive().getSheetByName('players');
  
  // シートの最終行を取得
  let lastRow: number = sheet.getLastRow();

  // コピーする行数
  let copyRow: number = lastRow + 1;
  
  let playerId: number = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
  sheet.getRange(copyRow, 1).setValue(playerId); // player_id
}

function save() {
  let playerId: string = Browser.inputBox("player_id を入力してください", Browser.Buttons.OK_CANCEL);
  const sheet: any = SpreadsheetApp.getActive().getSheetByName('players');
  const row: number = findRow(sheet, playerId, 1);
  const datasetId: string = 'goldpony';
  const tableId: string = 'players';

  let playerNo: number = sheet.getRange(`B${row}`).isBlank() ? null : sheet.getRange(`B${row}`).getValue();
  let lastName: string = sheet.getRange(`C${row}`).getValue();
  let firstName: string = sheet.getRange(`D${row}`).getValue();
  let member: number = sheet.getRange(`E${row}`).getValue();
  let pitcher: number = sheet.getRange(`F${row}`).getValue();
  let catcher: number = sheet.getRange(`G${row}`).getValue();
  let first: number = sheet.getRange(`H${row}`).getValue();
  let second: number = sheet.getRange(`I${row}`).getValue();
  let third: number = sheet.getRange(`J${row}`).getValue();
  let shortstop: number = sheet.getRange(`K${row}`).getValue();
  let outfielder: number = sheet.getRange(`L${row}`).getValue();
  let photo_url: string = sheet.getRange(`M${row}`).getValue();

  const projectId: string = 'nifty-bindery-293409';
  const query: string = `#StandardSQL \n delete from ${datasetId}.${tableId} where player_id = ${playerId};INSERT INTO ${datasetId}.${tableId} (player_id, player_no, last_name, first_name, member, pitcher, catcher, first, second, third, shortstop, outfielder, photo_url) values (${playerId}, ${playerNo}, "${lastName}", "${firstName}", ${member}, ${pitcher}, ${catcher}, ${first}, ${second}, ${third}, ${shortstop}, ${outfielder}, "${photo_url}");`;
  let request = {
    query: query
  };
  const bigqueryJobs: any = Bigquery.Jobs;
  let queryResults: any = bigqueryJobs.query(request, projectId);
  queryResults.jobReference.jobId;
}

function findRow(sheet: any, val: string, col: number){
  let dat: string = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得

  for(var i=1;i<dat.length;i++){
    if(String(dat[i][col-1]) === val){
      return i+1;
    }
  }
  return 0;
}

function convCsv(range: any) {
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
