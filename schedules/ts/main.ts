function addRow() {
  const sheet: any = SpreadsheetApp.getActive().getSheetByName('schedules');
  
  // シートの最終行を取得
  let lastRow: number = sheet.getLastRow();

  // コピーする行数
  let copyRow: number = lastRow + 1;
  
  let scheduleId: number = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
  sheet.getRange(copyRow, 1).setValue(scheduleId); // schedule_id
  sheet.getRange(lastRow, 2).copyTo(sheet.getRange(copyRow, 2)); // game_date
  sheet.getRange(lastRow, 3).copyTo(sheet.getRange(copyRow, 3)); // start_time
  sheet.getRange(lastRow, 4).copyTo(sheet.getRange(copyRow, 4)); // end_time
  sheet.getRange(lastRow, 5).copyTo(sheet.getRange(copyRow, 5)); // stadium_id
  sheet.getRange(lastRow, 6).copyTo(sheet.getRange(copyRow, 6)); // stadium_name
  sheet.getRange(lastRow, 8).copyTo(sheet.getRange(copyRow, 8)); // opponent_team_id
  sheet.getRange(lastRow, 9).copyTo(sheet.getRange(copyRow, 9)); // opponent_team_name
}

function save() {
  let scheduleId: string = Browser.inputBox("schedule_id を入力してください", Browser.Buttons.OK_CANCEL);
  const sheet: any = SpreadsheetApp.getActive().getSheetByName('schedules');
  const row: number = findRow(sheet, scheduleId, 1);
  const datasetId: string = 'goldpony';
  const tableId: string = 'schedules';
  
  let gameDate: string = sheet.getRange(`B${row}`).getValue();
  let startTime: string = sheet.getRange(`C${row}`).getValue();
  let endTime: string = sheet.getRange(`D${row}`).getValue();
  let stadiumId: number = sheet.getRange(`E${row}`).isBlank() ? null : sheet.getRange(`E${row}`).getValue();
  let stadiumName: string = sheet.getRange(`F${row}`).isBlank() ? null : sheet.getRange(`F${row}`).getValue();
  if (stadiumName !== null) stadiumName = "${stadiumName}";
  let mapUrl: string = sheet.getRange(`G${row}`).isBlank() ? null : sheet.getRange(`G${row}`).getValue();
  if (mapUrl !== null) mapUrl = "${mapUrl}";
  let opponentTeamId: number = sheet.getRange(`H${row}`).isBlank() ? null : sheet.getRange(`H${row}`).getValue();
  let opponentTeamName: string = sheet.getRange(`I${row}`).isBlank() ? null : sheet.getRange(`I${row}`).getValue();
  if (opponentTeamName !== null) opponentTeamName = "${opponentTeamName}";

  const projectId: string = 'nifty-bindery-293409';
  const query: string = `#StandardSQL \n delete from ${datasetId}.${tableId} where schedule_id = ${scheduleId};INSERT INTO ${datasetId}.${tableId} (schedule_id, game_date, start_time, end_time, stadium_id, stadium_name, map_url, opponent_team_id, opponent_team_name) values (${scheduleId}, "${gameDate}", "${startTime}", "${endTime}", ${stadiumId}, ${stadiumName}, ${mapUrl}, ${opponentTeamId}, ${opponentTeamName});`;
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
    let data = range.getValues();
    let ret = "";
    if (data.length > 1) {
      let csv = "";
      for (let i = 0; i < data.length; i++) {
        for (let j = 0; j < data[i].length; j++) {
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
