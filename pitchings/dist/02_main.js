const PITCHING_RATE = 1.0

function addNewRow() {
  var gameBook = SpreadsheetApp.openById('18eF9bkU7HdPs-1sHjnQhY7_l1foO3DLnR2WZ5cQuzfo');
  var gameSheet = gameBook.getSheetByName('games');
  var gameSheetLastRow = gameSheet.getLastRow();
  var gameId = gameSheet.getRange(gameSheetLastRow, 1).getValue();

  var sheet = SpreadsheetApp.getActive().getSheetByName("pitchings");
  var lastRow = sheet.getLastRow();
  var copyRow = lastRow + 1;
  sheet.getRange(copyRow, 1).setValue(gameId);
  sheet.getRange(lastRow, 2).copyTo(sheet.getRange(copyRow, 2)); // player_id
  sheet.getRange(lastRow, 3).copyTo(sheet.getRange(copyRow, 3)); // player_name
  sheet.getRange(copyRow, 4).setValue(1); // turn
  sheet.getRange(copyRow, 5).setValue(0);
  sheet.getRange(copyRow, 7).setValue(0);
  sheet.getRange(copyRow, 8).setValue(0);
  sheet.getRange(copyRow, 9).setValue(0);
  sheet.getRange(copyRow, 10).setValue(0);
  sheet.getRange(copyRow, 11).setValue(0);
  sheet.getRange(copyRow, 12).setValue(0);
  sheet.getRange(copyRow, 13).setValue(0);
  sheet.getRange(copyRow, 14).setValue(0);
  sheet.getRange(copyRow, 15).setValue(0);
  sheet.getRange(copyRow, 16).setValue(0);
  sheet.getRange(copyRow, 17).setValue(0);
  sheet.getRange(copyRow, 18).setValue(0);
  sheet.getRange(copyRow, 19).setValue(0);
}

function addRow() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("pitchings");
  
  // シートの最終行を取得
  var lastRow = sheet.getLastRow();

  // コピーする行数
  var copyRow = lastRow + 1;
  
  sheet.getRange(lastRow, 1).copyTo(sheet.getRange(copyRow, 1)); // game_id
  sheet.getRange(lastRow, 2).copyTo(sheet.getRange(copyRow, 2)); // player_id
  sheet.getRange(lastRow, 3).copyTo(sheet.getRange(copyRow, 3)); // player_name
  var lastTurn = sheet.getRange(lastRow, 4).getValue();
  var turn = 1;
  if (lastTurn) turn = lastTurn + 1;
  sheet.getRange(copyRow, 4).setValue(turn); // turn
  sheet.getRange(copyRow, 5).setValue(0);
  sheet.getRange(copyRow, 7).setValue(0);
  sheet.getRange(copyRow, 8).setValue(0);
  sheet.getRange(copyRow, 9).setValue(0);
  sheet.getRange(copyRow, 10).setValue(0);
  sheet.getRange(copyRow, 11).setValue(0);
  sheet.getRange(copyRow, 12).setValue(0);
  sheet.getRange(copyRow, 13).setValue(0);
  sheet.getRange(copyRow, 14).setValue(0);
  sheet.getRange(copyRow, 15).setValue(0);
  sheet.getRange(copyRow, 16).setValue(0);
  sheet.getRange(copyRow, 17).setValue(0);
  sheet.getRange(copyRow, 18).setValue(0);
  sheet.getRange(copyRow, 19).setValue(0);
}

async function execute(gameId) {
  const year = await getLatestYear();
  const gamesCount = await calculateGamesCount(year);
  const regulation = gamesCount * PITCHING_RATE;

  await insertPitchings(gameId);
  await insertFirestorePitchings(gameId);
  await insertFirestorePitchingStats(year, regulation);
  await insertFirestorePitchingPlayerTotals(year);
  await insertFirestorePitchingYearTotals(year, gamesCount);

  // Leaders
  await insertSumPitchings(year);
  await insertFirestorePitchingLeaders(year, regulation);
}

async function insertPitchings(gameId) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('pitchings');
  var rows = findRows(sheet, Number(gameId), 1);

  var values = [];
  rows.forEach(function(row) {
    var playerId = sheet.getRange(`B${row}`).getValue();
    var playerName = sheet.getRange(`C${row}`).getValue();
    var turn = sheet.getRange(`D${row}`).getValue();
    var inning = sheet.getRange(`E${row}`).getValue();
    var inning_rest = sheet.getRange(`F${row}`).getValue();
    if(!inning_rest) inning_rest = null;
    var bat = sheet.getRange(`G${row}`).getValue();
    var hit = sheet.getRange(`H${row}`).getValue();
    var hr = sheet.getRange(`I${row}`).getValue();
    var bb = sheet.getRange(`J${row}`).getValue();
    var db = sheet.getRange(`K${row}`).getValue();
    var so = sheet.getRange(`L${row}`).getValue();
    var wp = sheet.getRange(`M${row}`).getValue();
    var run = sheet.getRange(`N${row}`).getValue();
    var erun = sheet.getRange(`O${row}`).getValue();
    var win = sheet.getRange(`P${row}`).getValue();
    var lose = sheet.getRange(`Q${row}`).getValue();
    var hold = sheet.getRange(`R${row}`).getValue();
    var save = sheet.getRange(`S${row}`).getValue();
  
    values.push(`(${gameId}, ${playerId}, "${playerName}", ${turn}, ${inning}, ${inning_rest}, ${bat}, ${hit}, ${hr}, ${bb}, ${db}, ${so}, ${wp}, ${run}, ${erun}, ${win}, ${lose}, ${hold}, ${save})`);
  });

  const query = `
    #StandardSQL \n
    DELETE FROM ${datasetId}.pitchings where game_id = ${gameId};
    INSERT INTO ${datasetId}.pitchings 
    (game_id, player_id, player_name, turn, inning, inning_rest,
      bat, hit, hr, bb, db, so, wp, run, erun, win, lose, hold, save)
    VALUES ${values.join()};
  `;
  var request = {
    query: query
  };
  var queryResults = BigQuery.Jobs.query(request, bqProjectId);
  queryResults.jobReference.jobId;
}

async function insertFirestorePitchings(gameId) {
  await deleteFirestoreDocument('pitchings', 'game_id', gameId);

  const query = `
    #StandardSQL \n
    SELECT
      game_id,
      players.player_id,
      players.player_no,
      players.last_name,
      players.first_name,
      players.photo_url,
      turn,
      inning,
      inning_rest,
      bat,
      hit,
      hr,
      bb,
      db,
      so,
      wp,
      run,
      erun,
      win,
      lose,
      hold,
      save
    FROM
      ${datasetId}.pitchings left join
      ${datasetId}.players ON pitchings.player_id = players.player_id
    WHERE
      game_id = ${gameId}
  `;

  let records = getQueryResult(query);
  records.shift();

  let data = {};
  records.forEach((record) => {
    data = {
      'game_id': parseInt(record[0]),
      'player_id': parseInt(record[1]),
      'player_no': parseInt(record[2]),
      'last_name': record[3],
      'first_name': record[4],
      'photo_url': record[5],
      'turn': parseInt(record[6]),
      'inning': record[7] === null ? null : parseInt(record[7]),
      'inning_rest': record[8] === null ? null : parseInt(record[8]),
      'bat': parseInt(record[9]),
      'hit': parseInt(record[10]),
      'hr': parseInt(record[11]),
      'bb': parseInt(record[12]),
      'db': parseInt(record[13]),
      'so': parseInt(record[14]),
      'wp': parseInt(record[15]),
      'run': parseInt(record[16]),
      'erun': parseInt(record[17]),
      'win': parseInt(record[18]),
      'lose': parseInt(record[19]),
      'hold': parseInt(record[20]),
      'save': parseInt(record[21]),
    };
    firestore.createDocument('pitchings', data);
  })
}

async function insertFirestorePitchingStats(year, regulation) {
  let query = `
    #StandardSQL \n
    with base as ( 
      select 
        player_id,
        count(games) games,
        cast(
          CASE 
            WHEN SUM(inning_rest) is NULL THEN SUM(inning)
            ELSE SUM(inning) + (SUM(inning_rest) / 3)
          END as float64
        ) as innings,
        mod(SUM(inning_rest), 3) as inning_rests,
        sum(bat) bats,
        sum(hit) hits,
        sum(hr) hrs,
        sum(bb) bbs,
        sum(db) dbs,
        sum(so) sos,
        sum(wp) wps,
        sum(run) runs,
        sum(erun) eruns,
        ifnull(sum(win), 0) wins,
        ifnull(sum(lose), 0) loses,
        ifnull(sum(hold), 0) holds,
        ifnull(sum(save), 0) saves
      from 
        goldpony.pitchings left join 
        goldpony.games on pitchings.game_id = games.game_id
      where 
        FORMAT_DATETIME("%Y", datetime(game_datetime)) = '${year}'
      group by 
        player_id
    )

    select 
      base.*,
      players.player_no,
      players.last_name,
      players.first_name,
      players.photo_url,
      case 
        when inning_rests is null then cast(innings as string)
        else concat(cast(innings as int64), '.', inning_rests)
      end innings_text,
      SAFE_DIVIDE(eruns * 9, innings) AS era,
      format('%.2f', ROUND(SAFE_DIVIDE(eruns * 9, innings), 2)) AS era_text,
      SAFE_DIVIDE((bbs + hits), (
        CASE 
          WHEN inning_rests is NULL OR inning_rests = 0 THEN innings
          ELSE innings + (inning_rests / 3)
        END
      )) AS whip,
      format('%.2f', SAFE_DIVIDE((bbs + hits), (
        CASE 
          WHEN inning_rests is NULL OR inning_rests = 0 THEN innings
          ELSE innings + (inning_rests / 3)
        END
      ))) as whip_text,
      format('%.3f', SAFE_DIVIDE(hits, (bats - bbs - dbs))) ave,
      CASE
        when hits = 0 then '.000'
        when SAFE_DIVIDE(hits, (bats - bbs - dbs)) = 1 then '1.00'
        else concat('.', cast(SAFE_DIVIDE(hits, (bats - bbs - dbs)) * 1000 as int64))
      end ave_text,
      CASE 
        WHEN inning_rests is NULL and innings >= ${regulation} THEN 1
        WHEN inning_rests >= 0 and (innings + inning_rests) >= ${regulation} THEN 1
        WHEN inning_rests is NULL and innings < ${regulation} THEN 2
        WHEN inning_rests >= 0 and (innings + inning_rests) < ${regulation} THEN 2
      END over_regulation
    from 
      base left join
      goldpony.players on base.player_id = players.player_id
  `;

  let records = getQueryResult(query);
  records.shift();

  let data = {};
  for await (let record of records) {
    let keyValues = [
      { idName: 'year', id: parseInt(year) },
      { idName: 'player_id', id: parseInt(record[0]) },
    ]
    await deleteFirestoreDocumentMultipleConditions('pitchingStats', keyValues);

    data = {
      'year': parseInt(year),
      'player_id': parseInt(record[0]),
      'games': parseInt(record[1]),
      'innings': parseInt(record[2]),
      'inning_rests': record[3] === null ? null : parseInt(record[3]),
      'ab': parseInt(record[4]),
      'hit': parseInt(record[5]),
      'hr': parseInt(record[6]),
      'bb': parseInt(record[7]),
      'db': parseInt(record[8]),
      'so': parseInt(record[9]),
      'wp': parseInt(record[10]),
      'run': parseInt(record[11]),
      'erun': parseInt(record[12]),
      'win': parseInt(record[13]),
      'lose': parseInt(record[14]),
      'hold': parseInt(record[15]),
      'save': parseInt(record[16]),
      'player_no': record[17] === null ? null : parseInt(record[17]),
      'last_name': record[18] === null ? null : record[18],
      'first_name': record[19] === null ? null : record[19],
      'photo_url': record[20] === null ? null : record[20],
      'inning_text': record[21] === null ? null : record[21],
      'era': record[22] === null ? null : parseFloat(record[22]),
      'era_text': record[23] === null ? null : record[23],
      'whip': record[24] === null ? null : parseFloat(record[24]),
      'whip_text': record[25] === null ? null : record[25],
      'ave': record[26] === null ? null : parseFloat(record[26]),
      'ave_text': record[27] === null ? null : record[27],
      'over_regulation': parseInt(record[28]),
    };
    firestore.createDocument('pitchingStats', data);
  }
}

async function insertFirestorePitchingPlayerTotals() {
  let query = `
    #StandardSQL \n
    with base as ( 
      select 
        player_id,
        count(games) games,
        cast(
          CASE 
            WHEN SUM(inning_rest) is NULL THEN SUM(inning)
            ELSE SUM(inning) + (SUM(inning_rest) / 3)
          END as float64
        ) as innings,
        mod(SUM(inning_rest), 3) as inning_rests,
        sum(bat) bats,
        sum(hit) hits,
        sum(hr) hrs,
        sum(bb) bbs,
        sum(db) dbs,
        sum(so) sos,
        sum(wp) wps,
        sum(run) runs,
        sum(erun) eruns,
        ifnull(sum(win), 0) wins,
        ifnull(sum(lose), 0) loses,
        ifnull(sum(hold), 0) holds,
        ifnull(sum(save), 0) saves
      from 
        goldpony.pitchings left join 
        goldpony.games on pitchings.game_id = games.game_id
      group by 
        player_id
    )

    select 
      base.*,
      players.player_no,
      players.last_name,
      players.first_name,
      players.photo_url,
      case 
        when inning_rests is null then cast(innings as string)
        else concat(cast(innings as int64), '.', inning_rests)
      end innings_text,
      SAFE_DIVIDE(eruns * 9, innings) AS era,
      format('%.2f', ROUND(SAFE_DIVIDE(eruns * 9, innings), 2)) AS era_text,
      SAFE_DIVIDE((bbs + hits), (
        CASE 
          WHEN inning_rests is NULL OR inning_rests = 0 THEN innings
          ELSE innings + (inning_rests / 3)
        END
      )) AS whip,
      format('%.2f', SAFE_DIVIDE((bbs + hits), (
        CASE 
          WHEN inning_rests is NULL OR inning_rests = 0 THEN innings
          ELSE innings + (inning_rests / 3)
        END
      ))) as whip_text,
      format('%.3f', SAFE_DIVIDE(hits, (bats - bbs - dbs))) ave,
      CASE
        when hits = 0 then '.000'
        when SAFE_DIVIDE(hits, (bats - bbs - dbs)) = 1 then '1.00'
        else concat('.', cast(SAFE_DIVIDE(hits, (bats - bbs - dbs)) * 1000 as int64))
      end ave_text
    from 
      base left join
      goldpony.players on base.player_id = players.player_id
  `;

  let records = getQueryResult(query);
  records.shift();

  let data = {};
  for await (let record of records) {
    await deleteFirestoreDocument('pitchingPlayerTotals', 'player_id', parseInt(record[0]));

    data = {
      'player_id': parseInt(record[0]),
      'games': parseInt(record[1]),
      'innings': parseInt(record[2]),
      'inning_rests': record[3] === null ? null : parseInt(record[3]),
      'ab': parseInt(record[4]),
      'hit': parseInt(record[5]),
      'hr': parseInt(record[6]),
      'bb': parseInt(record[7]),
      'db': parseInt(record[8]),
      'so': parseInt(record[9]),
      'wp': parseInt(record[10]),
      'run': parseInt(record[11]),
      'erun': parseInt(record[12]),
      'win': parseInt(record[13]),
      'lose': parseInt(record[14]),
      'hold': parseInt(record[15]),
      'save': parseInt(record[16]),
      'player_no': record[17] === null ? null : parseInt(record[17]),
      'last_name': record[18] === null ? null : record[18],
      'first_name': record[19] === null ? null : record[19],
      'photo_url': record[20] === null ? null : record[20],
      'inning_text': record[21] === null ? null : record[21],
      'era': record[22] === null ? null : parseFloat(record[22]),
      'era_text': record[23] === null ? null : record[23],
      'whip': record[24] === null ? null : parseFloat(record[24]),
      'whip_text': record[25] === null ? null : record[25],
      'ave': record[26] === null ? null : parseFloat(record[26]),
      'ave_text': record[27] === null ? null : record[27],
    };
    firestore.createDocument('pitchingPlayerTotals', data);
  }
}

async function insertFirestorePitchingYearTotals(year, gamesCount) {
  let query = `
    #StandardSQL \n
    with base as ( 
      select
        FORMAT_DATETIME("%Y", datetime(game_datetime)) year,
        ${gamesCount} games,
        cast(
          CASE 
            WHEN SUM(inning_rest) is NULL THEN SUM(inning)
            ELSE SUM(inning) + (SUM(inning_rest) / 3)
          END as float64
        ) as innings,
        mod(SUM(inning_rest), 3) as inning_rests,
        sum(bat) bats,
        sum(hit) hits,
        sum(hr) hrs,
        sum(bb) bbs,
        sum(db) dbs,
        sum(so) sos,
        sum(wp) wps,
        sum(run) runs,
        sum(erun) eruns,
        ifnull(sum(win), 0) wins,
        ifnull(sum(lose), 0) loses,
        ifnull(sum(hold), 0) holds,
        ifnull(sum(save), 0) saves
      from 
        goldpony.pitchings left join 
        goldpony.games on pitchings.game_id = games.game_id
      where
        FORMAT_DATETIME("%Y", datetime(game_datetime)) = '${year}'
      group by 
        year
    )

    select 
      base.*,
      case 
        when inning_rests is null then cast(innings as string)
        else concat(cast(innings as int64), '.', inning_rests)
      end innings_text,
      SAFE_DIVIDE(eruns * 9, innings) AS era,
      format('%.2f', ROUND(SAFE_DIVIDE(eruns * 9, innings), 2)) AS era_text,
      SAFE_DIVIDE((bbs + hits), (
        CASE 
          WHEN inning_rests is NULL OR inning_rests = 0 THEN innings
          ELSE innings + (inning_rests / 3)
        END
      )) AS whip,
      format('%.2f', SAFE_DIVIDE((bbs + hits), (
        CASE 
          WHEN inning_rests is NULL OR inning_rests = 0 THEN innings
          ELSE innings + (inning_rests / 3)
        END
      ))) as whip_text,
      format('%.3f', SAFE_DIVIDE(hits, (bats - bbs - dbs))) ave,
      CASE
        when hits = 0 then '.000'
        when SAFE_DIVIDE(hits, (bats - bbs - dbs)) = 1 then '1.00'
        else concat('.', cast(SAFE_DIVIDE(hits, (bats - bbs - dbs)) * 1000 as int64))
      end ave_text
    from 
      base
  `;

  let records = getQueryResult(query);
  records.shift();

  let data = {};
  for await (let record of records) {
    await deleteFirestoreDocument('pitchingYearTotals', 'year', year);

    data = {
      'year': parseInt(record[0]),
      'games' : parseInt(record[1]),
      'innings': parseInt(record[2]),
      'inning_rests': record[3] === null ? null : parseInt(record[3]),
      'ab': parseInt(record[4]),
      'hit': parseInt(record[5]),
      'hr': parseInt(record[6]),
      'bb': parseInt(record[7]),
      'db': parseInt(record[8]),
      'so': parseInt(record[9]),
      'wp': parseInt(record[10]),
      'run': parseInt(record[11]),
      'erun': parseInt(record[12]),
      'win': parseInt(record[13]),
      'lose': parseInt(record[14]),
      'hold': parseInt(record[15]),
      'save': parseInt(record[16]),
      'inning_text': record[17] === null ? null : record[17],
      'era': record[18] === null ? null : parseFloat(record[18]),
      'era_text': record[19] === null ? null : record[19],
      'whip': record[20] === null ? null : parseFloat(record[20]),
      'whip_text': record[21] === null ? null : record[21],
      'ave': record[22] === null ? null : parseFloat(record[22]),
      'ave_text': record[23] === null ? null : record[23],
    };
    firestore.createDocument('pitchingYearTotals', data);
  }
}

async function getLatestYear() {
  let query = `
    #StandardSQL \n
    SELECT
      FORMAT_DATETIME('%Y', datetime(game_datetime)) year
    FROM
      ${datasetId}.games
    ORDER BY
      game_datetime DESC
    LIMIT
      1
  `;

  let records = getQueryResult(query);
  records.shift();

  return records[0][0];
}

async function insertSumPitchings(year) {
  const query = `
    #StandardSQL \n
    TRUNCATE TABLE ${datasetId}.sum_pitchings_latest_year;
    INSERT INTO ${datasetId}.sum_pitchings_latest_year
    (player_id, games, inning, inning_rest, total_inning, bat, hit, hr, bb, db, so, wp, run, erun, win, lose, hold, save, last_name, first_name, player_no, photo_url)
    SELECT
      sum_pitchings.*,
      players.last_name,
      players.first_name,
      players.player_no,
      players.photo_url,
    FROM
      (
        SELECT
          player_id,
          COUNT(player_id) games,
          SUM(inning) inning,
          SUM(inning_rest) inning_rest,
          CASE 
            WHEN SUM(inning_rest) is NULL THEN SUM(inning)
            ELSE SUM(inning) + (SUM(inning_rest) / 3)
          END as total_inning,
          SUM(bat) bat,
          SUM(hit) hit,
          SUM(hr) hr,
          SUM(bb) bb,
          SUM(db) db,
          SUM(so) so,
          SUM(wp) wp,
          SUM(run) run,
          SUM(erun) erun,
          SUM(win) win,
          SUM(lose) lose,
          SUM(hold) hold,
          SUM(save) save
        FROM
          ${datasetId}.pitchings
        WHERE
          game_id IN (
            SELECT
              game_id
            FROM
              ${datasetId}.games
            WHERE
              FORMAT_DATETIME('%Y', datetime(game_datetime)) = '${year}'
          )
        GROUP BY
          player_id
      ) sum_pitchings LEFT JOIN goldpony.players ON sum_pitchings.player_id = players.player_id
    ;
  `;

  request = {
    query: query
  };
  BigQuery.Jobs.query(request, bqProjectId);
}

async function insertFirestorePitchingLeaders(year, regulation) {
  // 既存の leaders のレコードを削除
  let leaders = firestore.query('pitchingLeaders').Execute();
  if (leaders.length > 0) {
    leaders.forEach(function (leader) {
      let index = leader.name.lastIndexOf('/');
      let documentId = leader.name.substring(index + 1);
      firestore.deleteDocument(`pitchingLeaders/${documentId}`);
    })
  }

  // era
  let query = `
    #StandardSQL \n
    SELECT
      player_id,
      last_name,
      first_name,
      player_no,
      photo_url,
      FORMAT('%.2f', ROUND(erun * 9 / total_inning, 2)) AS era_text,
      ROUND(erun * 9 / total_inning, 2) AS era
    FROM
      ${datasetId}.sum_pitchings_latest_year
    WHERE
      total_inning >= ${regulation}
    ORDER BY
      era
    LIMIT
      1
  `;

  records = getQueryResult(query);
  records.shift();
  record = records[0];
  let data = {
    'key': 'era',
    'player_id': record[0],
    'last_name': record[1],
    'first_name': record[2],
    'player_no': record[3],
    'photo_url': record[4],
    'value': record[5]
  };
  firestore.createDocument('pitchingLeaders', data);

  // so
  query = `
    #StandardSQL \n
    SELECT
      player_id,
      last_name,
      first_name,
      player_no,
      photo_url,
      so
    FROM
      ${datasetId}.sum_pitchings_latest_year
    ORDER BY
      so desc
    LIMIT
      1
  `;

  records = getQueryResult(query);
  records.shift();
  record = records[0];
  data = {
    'key': 'so',
    'player_id': record[0],
    'last_name': record[1],
    'first_name': record[2],
    'player_no': record[3],
    'photo_url': record[4],
    'value': record[5]
  };

  firestore.createDocument('pitchingLeaders', data);

  // whip
  query = `
    #StandardSQL \n
    SELECT
      player_id,
      last_name,
      first_name,
      player_no,
      photo_url,
      ROUND((bb + hit) / total_inning, 2) as whip_text,
      (bb + hit) / total_inning as whip
    FROM
      ${datasetId}.sum_pitchings_latest_year
    WHERE
      total_inning >= ${regulation}
    ORDER BY
      whip
    LIMIT
      1
  `;

  records = getQueryResult(query);
  records.shift();
  record = records[0];
  data = {
    'key': 'whip',
    'player_id': record[0],
    'last_name': record[1],
    'first_name': record[2],
    'player_no': record[3],
    'photo_url': record[4],
    'value': record[5]
  };

  firestore.createDocument('pitchingLeaders', data);
}

async function calculateGamesCount(year) {
  let query = `
    #StandardSQL \n   
    SELECT
      COUNT(1)
    FROM
      ${datasetId}.games
    WHERE
      FORMAT_DATETIME('%Y', datetime(game_datetime)) = '${year}'
  `;
  records = getQueryResult(query);
  records.shift();
  return records[0][0];
}
