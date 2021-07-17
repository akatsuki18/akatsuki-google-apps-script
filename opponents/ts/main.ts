function addRow() {
  const sheet: any = SpreadsheetApp.getActive().getSheetByName('opponents');
  
  // シートの最終行を取得
  let lastRow: number = sheet.getLastRow();

  // コピーする行数
  let copyRow: number = lastRow + 1;
  
  let opponentId: number = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
  sheet.getRange(copyRow, 1).setValue(opponentId); // opponent_id
}

function save() {
  let opponentId: string = Browser.inputBox("opponent_id を入力してください", Browser.Buttons.OK_CANCEL);
  const sheet: any = SpreadsheetApp.getActive().getSheetByName('opponents');
  const row: number = findRow(sheet, opponentId, 1);
  const datasetId: string = 'goldpony';
  const tableId: string = 'opponents';
  
  let teamName: string = sheet.getRange(`B${row}`).getValue();
  let teamUrl: string = sheet.getRange(`C${row}`).isBlank() ? null : sheet.getRange(`C${row}`).getValue();
  if (teamUrl !== null) teamUrl = `"${teamUrl}"`;

  const projectId: string = 'nifty-bindery-293409';
  const query: string = `#StandardSQL \n delete from ${datasetId}.${tableId} where opponent_id = ${opponentId};INSERT INTO ${datasetId}.${tableId} (opponent_id, team_name, team_url) values (${opponentId}, "${teamName}", ${teamUrl});`;
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
