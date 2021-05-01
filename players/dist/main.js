"use strict";
function addRow() {
    var sheet = SpreadsheetApp.getActive().getSheetByName('players');
    // シートの最終行を取得
    var lastRow = sheet.getLastRow();
    // コピーする行数
    var copyRow = lastRow + 1;
    var playerId = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
    sheet.getRange(copyRow, 1).setValue(playerId); // player_id
}
function save() {
    var playerId = Browser.inputBox("player_id を入力してください", Browser.Buttons.OK_CANCEL);
    var sheet = SpreadsheetApp.getActive().getSheetByName('players');
    var row = findRow(sheet, playerId, 1);
    var datasetId = 'goldpony';
    var tableId = 'players';
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
    var photo_url = sheet.getRange("M" + row).getValue();
    var projectId = 'nifty-bindery-293409';
    var query = "#StandardSQL \n delete from " + datasetId + "." + tableId + " where player_id = " + playerId + ";INSERT INTO " + datasetId + "." + tableId + " (player_id, player_no, last_name, first_name, member, pitcher, catcher, first, second, third, shortstop, outfielder, photo_url) values (" + playerId + ", " + playerNo + ", \"" + lastName + "\", \"" + firstName + "\", " + member + ", " + pitcher + ", " + catcher + ", " + first + ", " + second + ", " + third + ", " + shortstop + ", " + outfielder + ", \"" + photo_url + "\");";
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
