"use strict";
function addRow() {
    var sheet = SpreadsheetApp.getActive().getSheetByName('schedules');
    // シートの最終行を取得
    var lastRow = sheet.getLastRow();
    // コピーする行数
    var copyRow = lastRow + 1;
    var scheduleId = Number(sheet.getRange(lastRow, 1).getValue()) + 1;
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
    var scheduleId = Browser.inputBox("schedule_id を入力してください", Browser.Buttons.OK_CANCEL);
    var sheet = SpreadsheetApp.getActive().getSheetByName('schedules');
    var row = findRow(sheet, scheduleId, 1);
    var gameDate = sheet.getRange("B" + row).getValue();
    var startTime = sheet.getRange("C" + row).getValue();
    var endTime = sheet.getRange("D" + row).getValue();
    var stadiumId = sheet.getRange("E" + row).isBlank() ? null : sheet.getRange("E" + row).getValue();
    var stadiumName = sheet.getRange("F" + row).isBlank() ? null : sheet.getRange("F" + row).getValue();
    var mapUrl = sheet.getRange("G" + row).isBlank() ? null : sheet.getRange("G" + row).getValue();
    var opponentTeamId = sheet.getRange("H" + row).isBlank() ? null : sheet.getRange("H" + row).getValue();
    var opponentTeamName = sheet.getRange("I" + row).isBlank() ? null : sheet.getRange("I" + row).getValue();
    var opponentTeamUrl = sheet.getRange("J" + row).isBlank() ? null : sheet.getRange("J" + row).getValue();
    var email = PropertiesService.getScriptProperties().getProperty('serviceAccountEmail');
    var keyString = PropertiesService.getScriptProperties().getProperty('serviceAccountKey');
    var key = keyString.replace(/\\n/g, "\n");
    var projectId = PropertiesService.getScriptProperties().getProperty('projectId');
    var firestore = FirestoreApp.getFirestore(email, key, projectId);
    // 登録データ
    var data = {
        'schedule_id': parseInt(scheduleId),
        'game_date': gameDate,
        'start_time': startTime,
        'end_time': endTime,
        'stadium_id': stadiumId,
        'stadium_name': stadiumName,
        'map_url': mapUrl,
        'opponent_team_id': opponentTeamId,
        'opponent_team_name': opponentTeamName,
        'opponent_team_url': opponentTeamUrl,
    };
    var schedule = firestore.query('schedules').Where('schedule_id', '==', parseInt(scheduleId)).Execute();
    if (schedule.length > 0) {
        var index = schedule[0].name.lastIndexOf('/');
        var documentId = schedule[0].name.substring(index + 1);
        firestore.updateDocument(`schedules/${documentId}`, data);
    } else {
        firestore.createDocument('schedules', data);
    }
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
