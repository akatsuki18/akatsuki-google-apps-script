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

function createTable() {
  // Replace this value with the project ID listed in the Google
  // Cloud Platform project.
  const projectId: string = 'nifty-bindery-293409';
  // Create a dataset in the BigQuery UI (https://bigquery.cloud.google.com)
  // and enter its ID below.
  const datasetId: string = 'goldpony';

  const tableId: string = 'schedules';
  const sheet: any = SpreadsheetApp.getActive().getSheetByName("schedules");
  
  // Create the table.
  let table: any = {
    tableReference: {
      projectId: projectId,
      datasetId: datasetId,
      tableId: tableId
    },
    schema: {
      fields: [
        {name: 'game_date', type: 'DATE'},
        {name: 'start_time', type: 'STRING'},
        {name: 'end_time', type: 'STRING'},
        {name: 'stadium_id', type: 'INTEGER'},
        {name: 'stadium_name', type: 'STRING'},
        {name: 'map_url', type: 'STRING'},
        {name: 'opponent_team_id', type: 'INTEGER'},
        {name: 'opponent_team_name', type: 'STRING'}
      ]
    }
  };

  const bigqueryTables: any = Bigquery.Tables;
  try{
    bigqueryTables.remove(projectId, datasetId, tableId); 
  } catch(e) {}
  bigqueryTables.insert(table, projectId, datasetId);

  let range: any = sheet.getDataRange();
  let csv: any = convCsv(range);
  let blob = Utilities.newBlob(csv).setContentType('application/octet-stream');
  let job = {
    configuration: {
      load: {
        destinationTable: {
          projectId: projectId,
          datasetId: datasetId,
          tableId: tableId
        },
        skipLeadingRows: 1
      }
    }
  };
  const bigqueryJobs: any = Bigquery.Jobs;
  job = bigqueryJobs.insert(job, projectId, blob);
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
      {name: '保存', functionName: 'createTable'},
    ]);
}
