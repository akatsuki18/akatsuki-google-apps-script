"use strict";
function addRow() {
    var sheet = SpreadsheetApp.getActive().getSheetByName('opponents');
    // シートの最終行を取得
    var lastRow = sheet.getLastRow();
    // コピーする行数
    var copyRow = lastRow + 1;
    var opponentId = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
    sheet.getRange(copyRow, 1).setValue(opponentId); // opponent_id
}
function save() {
    var opponentId = Browser.inputBox("opponent_id を入力してください", Browser.Buttons.OK_CANCEL);
    var sheet = SpreadsheetApp.getActive().getSheetByName('opponents');
    var row = findRow(sheet, opponentId, 1);
    var datasetId = 'goldpony';
    var tableId = 'opponents';
    var teamName = sheet.getRange("B" + row).getValue();
    var teamUrl = sheet.getRange("C" + row).isBlank() ? null : sheet.getRange("C" + row).getValue();
    if (teamUrl !== null)
        teamUrl = "\"" + teamUrl + "\"";
    var projectId = 'nifty-bindery-293409';
    var query = "#StandardSQL \n delete from " + datasetId + "." + tableId + " where opponent_id = " + opponentId + ";INSERT INTO " + datasetId + "." + tableId + " (opponent_id, team_name, team_url) values (" + opponentId + ", \"" + teamName + "\", " + teamUrl + ");";
    var request = {
        query: query
    };
    var bigqueryJobs = Bigquery.Jobs;
    var queryResults = bigqueryJobs.query(request, projectId);
    queryResults.jobReference.jobId;
}
function findRow(sheet, val, col) {
    var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
    for (var i = 1; i < dat.length; i++) {
        if (String(dat[i][col - 1]) === val) {
            return i + 1;
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
                if (i < data.length - 1) {
                    csv += data[i].join(",") + "\r\n";
                }
                else {
                    csv += data[i];
                }
            }
            ret = csv;
        }
        return ret;
    }
    catch (e) {
        Logger.log(e);
    }
}
function onOpen() {
    SpreadsheetApp
        .getActiveSpreadsheet()
        .addMenu('データ登録', [
        { name: '行追加', functionName: 'addRow' },
        { name: '保存', functionName: 'save' },
    ]);
}
