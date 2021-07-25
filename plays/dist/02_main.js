const HITTING_RATE = 2.1

function addRow() {
  let sheet = SpreadsheetApp.getActive().getSheetByName("plays");
  
  // シートの最終行を取得
  let lastRow = sheet.getLastRow();

  // コピーする行数
  let copyRow = lastRow + 1;
  
  sheet.getRange(lastRow, 1).copyTo(sheet.getRange(copyRow, 1)); // game_id
  sheet.getRange(lastRow, 2).copyTo(sheet.getRange(copyRow, 2)); // player_id
  sheet.getRange(lastRow, 3).copyTo(sheet.getRange(copyRow, 3)); // player_name
  let nextSort = Number(sheet.getRange(lastRow, 4).getValue()) + 1;
  sheet.getRange(copyRow, 4).setValue(nextSort); // sort
  sheet.getRange(lastRow, 5).copyTo(sheet.getRange(copyRow, 5)); // inning
  let nextBattingOrder = Number(sheet.getRange(lastRow, 6).getValue()) + 1; 
  sheet.getRange(copyRow, 6).setValue(nextBattingOrder);  // batting_order
  sheet.getRange(copyRow, 7).setValue(1); // turn
  sheet.getRange(lastRow, 8).copyTo(sheet.getRange(copyRow, 8)); // direction
  sheet.getRange(lastRow, 9).copyTo(sheet.getRange(copyRow, 9)); // result
  sheet.getRange(copyRow, 10).setValue(0); // run
  sheet.getRange(copyRow, 11).setValue(0); // rbi
  sheet.getRange(copyRow, 12).setValue(0); // steal
}

async function execute(gameId) {
  const year = await getLatestYear();
  const gamesCount = await calculateGamesCount(year);
  const regulation = gamesCount * HITTING_RATE;

  await insertPlays(gameId);
  await aggregateHittings(gameId);
  await insertHittings(gameId);
  await insertFirestoreHittings(gameId);
  await insertFirestoreHittingStats(year, regulation);
  await insertFirestoreHittingPlayerTotals(year);
  await insertFirestoreHittingYearTotals(year, gamesCount);

  // Leaders
  await insertSumHittings(year);
  await insertFirestoreHittingLeaders(year, regulation);
}

async function insertPlays(gameId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = findRows(sheet, Number(gameId), 1);
  const tableId = 'plays';

  await deleteFirestoreDocument(tableId, 'game_id', gameId);
  await deleteFirestoreDocument('playVideos', 'game_id', gameId);

  let values = [];
  let data;
  rows.forEach(function(row) {
    let gameId = sheet.getRange(`A${Number(row)}`).getValue();
    let playerId = sheet.getRange(`B${Number(row)}`).getValue();
    let playerName = sheet.getRange(`C${Number(row)}`).getValue();
    let sort = sheet.getRange(`D${Number(row)}`).getValue();
    let inning = sheet.getRange(`E${Number(row)}`).getValue();
    let batting_order = sheet.getRange(`F${Number(row)}`).getValue();
    let turn = sheet.getRange(`G${Number(row)}`).getValue();
    let direction = sheet.getRange(`H${Number(row)}`).getValue();
    let result = sheet.getRange(`I${Number(row)}`).getValue();
    let run = sheet.getRange(`J${Number(row)}`).getValue();
    let rbi = sheet.getRange(`K${Number(row)}`).getValue();
    let steal = sheet.getRange(`L${Number(row)}`).getValue();
    let video_url = sheet.getRange(`M${Number(row)}`).getValue();

    values.push(`(${gameId}, ${playerId}, "${playerName}", ${sort}, ${inning}, ${batting_order}, ${turn}, "${direction}", "${result}", ${run}, ${rbi}, ${steal}, "${video_url}")`);
    let query = `SELECT player_id, first_name, last_name, player_no, photo_url FROM ${datasetId}.players WHERE player_id = ${playerId}`;
    data = getQueryResult(query);

    let firstName = data[1][1];
    let lastName = data[1][2];
    let playerNo = data[1][3];
    let photoUrl = data[1][4];

    let record = {
      'game_id': parseInt(gameId),
      'player_id': parseInt(playerId),
      'first_name': firstName,
      'last_name': lastName,
      'player_no': playerNo,
      'photo_url': photoUrl,
      'sort': parseInt(sort),
      'inning': inning,
      'batting_order': parseInt(batting_order),
      'turn': parseInt(turn),
      'direction': direction,
      'result': result,
      'run': parseInt(run),
      'rbi': parseInt(rbi),
      'steal': parseInt(steal),
      'video_url': video_url,
    };
    firestore.createDocument(tableId, record);

    if (record['video_url'] !== '') firestore.createDocument('playVideos', record);
  });

  const insertQuery = `
    #StandardSQL \n
    delete from ${datasetId}.${tableId} where game_id = ${gameId};
    INSERT INTO ${datasetId}.${tableId} 
    (game_id, player_id, player_name, sort, inning, batting_order, turn, direction, result, run, rbi, steal, video_url)
    values ${values.join()};
  `;
  let insertRequest = {
    query: insertQuery
  };
  BigQuery.Jobs.query(insertRequest, bqProjectId);
}

// plays から集計
async function aggregateHittings(gameId) {
  const query = `
    #StandardSQL \n
    truncate table ${datasetId}.${tempTableName};
    insert into ${datasetId}.${tempTableName}
    select
      game_id,
      player_id,
      batting_order,
      turn,
      1 as pa,
      case
        when result = '四' or result = '四球' or result = '死' or result = '死球' or result = '打妨' or result = '犠打' or result = '犠飛' then 0
        else 1 
      end as ab,
      run,
      rbi,
      case
        when result = '四' or result = '四球' then 1
        else 0
      end as bb,
      case
        when result = '死' or result = '死球' then 1
        else 0
      end as db,
      case
        when result = '犠打' then 1
        else 0
      end as sh,
      case
        when result = '犠飛' then 1
        else 0
      end as sf,
      case
        when result = '三振' then 1
        else 0
      end as so,
      steal,
      case
        when result = '安' then 1
        else 0
      end as h1,
      case
        when result = '二' then 1
        else 0
      end as h2,
      case
        when result = '三' then 1
        else 0
      end as h3,
      case
        when result = '本' then 1
        else 0
      end as hr
    from
      ${datasetId}.plays
    where
      game_id = ${gameId}
  `;

  const request = {
    query: query
  };
  BigQuery.Jobs.query(request, bqProjectId);
}

async function insertHittings(gameId) {
  const query = `
    #StandardSQL \n
    delete from ${datasetId}.hittings where game_id = ${gameId};
    insert into ${datasetId}.hittings
    select
      game_id,
      player_id,
      batting_order,
      turn,
      sum(pa) as pa,
      sum(ab) as ab,
      sum(run) as run,
      sum(h1) + sum(h2) + sum(h3) + sum(hr) as hit,
      sum(rbi) as rbi,
      sum(bb) as bb,
      sum(db) as db,
      sum(sh) as sh,
      sum(sf) as sf,
      sum(so) as so,
      sum(steal) as steal,
      sum(h2) as hit2,
      sum(h3) as hit3,
      sum(hr) as hr
    from
      ${datasetId}.${tempTableName}
    group by
      game_id,
      player_id,
      batting_order,
      turn
  `;

  const request = {
    query: query
  };
  BigQuery.Jobs.query(request, bqProjectId);
}

async function insertFirestoreHittings(gameId) {
  await deleteFirestoreDocument('hittings', 'game_id', gameId);

  const query = `
    #StandardSQL \n
    SELECT
      game_id,
      hittings.player_id,
      batting_order,
      turn,
      pa,
      ab,
      run,
      hit,
      rbi,
      bb,
      db,
      sh,
      sf,
      so,
      steal,
      hit2,
      hit3,
      hr,
      players.last_name,
      players.first_name,
      players.player_no,
      players.photo_url
    FROM
      ${datasetId}.hittings left join
      ${datasetId}.players ON hittings.player_id = players.player_id
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
      'batting_order': parseInt(record[2]),
      'turn': parseInt(record[3]),
      'pa': parseInt(record[4]),
      'ab': parseInt(record[5]),
      'run': parseInt(record[6]),
      'hit': parseInt(record[7]),
      'rbi': parseInt(record[8]),
      'bb': parseInt(record[9]),
      'db': parseInt(record[10]),
      'sh': parseInt(record[11]),
      'sf': parseInt(record[12]),
      'so': parseInt(record[13]),
      'steal': parseInt(record[14]),
      'hit2': parseInt(record[15]),
      'hit3': parseInt(record[16]),
      'hr': parseInt(record[17]),
      'last_name': record[18],
      'first_name': record[19],
      'player_no': record[20],
      'photo_url': record[21],
    };
    firestore.createDocument('hittings', data);
  })
}

async function insertFirestoreHittingStats(year, regulation) {
  const query = `
    #StandardSQL \n
    select
      FORMAT_DATETIME("%Y", datetime(games.game_datetime)) year,
      players.player_id,
      players.player_no,
      players.last_name,
      players.first_name,
      players.photo_url,
      count(1) games,
      sum(pa) pa,
      sum(ab) ab,
      sum(run) run,
      sum(hit) hit,
      sum(rbi) rbi,
      sum(bb) bb,
      sum(db) db,
      sum(sh) sh,
      sum(sf) sf,
      sum(so) so,
      sum(steal) steal,
      sum(hit2) hit2,
      sum(hit3) hit3,
      sum(hr) hr,
      format('%.3f', SAFE_DIVIDE(sum(hittings.hit), sum(hittings.ab))) ave,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(sum(hittings.hit), sum(hittings.ab))) as string), 2, 4)
      END as ave_text,
      format('%.3f', SAFE_DIVIDE(sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db), sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf))) obp,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db), sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf))) as string), 2, 4)
      END AS obp_text,
      format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) slg,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        WHEN SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)) >= 1 THEN
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) as string), 1, 5)
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) as string), 2, 4)
      END AS slg_text,
      format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) ops,
      CASE 
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        WHEN (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) >= 1 THEN
          SUBSTR(CAST(format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) as string), 1, 5)
        ELSE
          SUBSTR(CAST(format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) as string), 2, 4) 
      END AS ops_text,
      CASE
        WHEN sum(hittings.pa) >= ${regulation} THEN 1
        ELSE 2
      END as over_regulation
    from
      ${datasetId}.hittings left join
      ${datasetId}.games on hittings.game_id = games.game_id left join 
      ${datasetId}.players on hittings.player_id = players.player_id
    where
      FORMAT_DATETIME("%Y", datetime(games.game_datetime)) = '${year}'
    group by
      year, 
      players.player_id,
      players.player_no,
      players.last_name,
      players.first_name,
      players.photo_url
    order by
      players.player_id
  `;

  let records = getQueryResult(query);
  records.shift();

  let data = {};
  for await (let record of records) {
    let keyValues = [
      { idName: 'year', id: parseInt(year) },
      { idName: 'player_id', id: parseInt(record[1]) },
    ]
    await deleteFirestoreDocumentMultipleConditions('hittingStats', keyValues);

    data = {
      'year': parseInt(record[0]),
      'player_id': parseInt(record[1]),
      'player_no': record[2] === null ? null : parseInt(record[2]),
      'last_name': record[3] === null ? null : record[3],
      'first_name': record[4] === null ? null : record[4],
      'photo_url': record[5] === null ? null : record[5],
      'games' : parseInt(record[6]),
      'pa': parseInt(record[7]),
      'ab': parseInt(record[8]),
      'run': parseInt(record[9]),
      'hit': parseInt(record[10]),
      'rbi': parseInt(record[11]),
      'bb': parseInt(record[12]),
      'db': parseInt(record[13]),
      'sh': parseInt(record[14]),
      'sf': parseInt(record[15]),
      'so': parseInt(record[16]),
      'steal': parseInt(record[17]),
      'hit2': parseInt(record[18]),
      'hit3': parseInt(record[19]),
      'hr': parseInt(record[20]),
      'ave' : record[21] === null ? null : parseFloat(record[21]),
      'ave_text' : record[22],
      'obp' : record[23] === null ? null : parseFloat(record[23]),
      'obp_text' : record[24],
      'slg' : record[25] === null ? null : parseFloat(record[25]),
      'slg_text' : record[26],
      'ops' : record[27] === null ? null : parseFloat(record[27]),
      'ops_text': record[28],
      'over_regulation' : parseInt(record[29]),
    };
    firestore.createDocument('hittingStats', data);
  }
}

async function insertFirestoreHittingPlayerTotals() {
  const query = `
    #StandardSQL \n
    select
      players.player_id,
      players.player_no,
      players.last_name,
      players.first_name,
      players.photo_url,
      count(1) games,
      sum(pa) pa,
      sum(ab) ab,
      sum(run) run,
      sum(hit) hit,
      sum(rbi) rbi,
      sum(bb) bb,
      sum(db) db,
      sum(sh) sh,
      sum(sf) sf,
      sum(so) so,
      sum(steal) steal,
      sum(hit2) hit2,
      sum(hit3) hit3,
      sum(hr) hr,
      format('%.3f', SAFE_DIVIDE(sum(hittings.hit), sum(hittings.ab))) ave,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(sum(hittings.hit), sum(hittings.ab))) as string), 2, 4)
      END as ave_text,
      format('%.3f', SAFE_DIVIDE(sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db), sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf))) obp,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db), sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf))) as string), 2, 4)
      END AS obp_text,
      format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) slg,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        WHEN SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)) >= 1 THEN
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) as string), 1, 5)
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) as string), 2, 4)
      END AS slg_text,
      format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) ops,
      CASE 
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        WHEN (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) >= 1 THEN
          SUBSTR(CAST(format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) as string), 1, 5)
        ELSE
          SUBSTR(CAST(format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) as string), 2, 4) 
      END AS ops_text
    from
      ${datasetId}.hittings left join
      ${datasetId}.games on hittings.game_id = games.game_id left join 
      ${datasetId}.players on hittings.player_id = players.player_id
    group by
      players.player_id,
      players.player_no,
      players.last_name,
      players.first_name,
      players.photo_url
  `;

  let records = getQueryResult(query);
  records.shift();

  let data = {};
  for await (let record of records) {
    await deleteFirestoreDocument('hittingPlayerTotals', 'player_id', parseInt(record[0]));

    data = {
      'player_id': parseInt(record[0]),
      'player_no': record[1] === null ? null : parseInt(record[1]),
      'last_name': record[2] === null ? null : record[2],
      'first_name': record[3] === null ? null : record[3],
      'photo_url': record[4] === null ? null : record[4],
      'games' : parseInt(record[5]),
      'pa': parseInt(record[6]),
      'ab': parseInt(record[7]),
      'run': parseInt(record[8]),
      'hit': parseInt(record[9]),
      'rbi': parseInt(record[10]),
      'bb': parseInt(record[11]),
      'db': parseInt(record[12]),
      'sh': parseInt(record[13]),
      'sf': parseInt(record[14]),
      'so': parseInt(record[15]),
      'steal': parseInt(record[16]),
      'hit2': parseInt(record[17]),
      'hit3': parseInt(record[18]),
      'hr': parseInt(record[19]),
      'ave' : record[20] === null ? null : parseFloat(record[20]),
      'ave_text' : record[21],
      'obp' : record[22] === null ? null : parseFloat(record[22]),
      'obp_text' : record[23],
      'slg' : record[24] === null ? null : parseFloat(record[24]),
      'slg_text' : record[25],
      'ops' : record[26] === null ? null : parseFloat(record[26]),
      'ops_text': record[27]
    };
    firestore.createDocument('hittingPlayerTotals', data);
  }
}

async function insertFirestoreHittingYearTotals(year, gamesCount) {
  const query = `
    #StandardSQL \n
    select
      FORMAT_DATETIME("%Y", datetime(games.game_datetime)) year,
      ${gamesCount} games,
      sum(pa) pa,
      sum(ab) ab,
      sum(run) run,
      sum(hit) hit,
      sum(rbi) rbi,
      sum(bb) bb,
      sum(db) db,
      sum(sh) sh,
      sum(sf) sf,
      sum(so) so,
      sum(steal) steal,
      sum(hit2) hit2,
      sum(hit3) hit3,
      sum(hr) hr,
      format('%.3f', SAFE_DIVIDE(sum(hittings.hit), sum(hittings.ab))) ave,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(sum(hittings.hit), sum(hittings.ab))) as string), 2, 4)
      END as ave_text,
      format('%.3f', SAFE_DIVIDE(sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db), sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf))) obp,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db), sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf))) as string), 2, 4)
      END AS obp_text,
      format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) slg,
      CASE
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        WHEN SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)) >= 1 THEN
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) as string), 1, 5)
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) as string), 2, 4)
      END AS slg_text,
      format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) ops,
      CASE 
        WHEN sum(hittings.ab) = 0 THEN 
          '---'
        WHEN (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab))) >= 1 THEN
          SUBSTR(CAST(format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) as string), 1, 5)
        ELSE
          SUBSTR(CAST(format('%.3f', (SAFE_DIVIDE((sum(hittings.hit) + sum(hittings.bb) + sum(hittings.db)), (sum(hittings.ab) + sum(hittings.bb) + sum(hittings.db) + sum(hittings.sf)))) + (SAFE_DIVIDE(((sum(hittings.hit) - sum(hittings.hit2) - sum(hittings.hit3) - sum(hittings.hr)) + (sum(hittings.hit2) * 2) + (sum(hittings.hit3) * 3) + (sum(hittings.hr) * 4)), sum(hittings.ab)))) as string), 2, 4) 
      END AS ops_text
    from
      ${datasetId}.hittings left join
      ${datasetId}.games on hittings.game_id = games.game_id left join 
      ${datasetId}.players on hittings.player_id = players.player_id
    where
      FORMAT_DATETIME("%Y", datetime(games.game_datetime)) = '${year}'
    group by
      year 
  `;

  let records = getQueryResult(query);
  records.shift();

  let data = {};
  for await (let record of records) {
    await deleteFirestoreDocument('hittingYearTotals', 'year', year);

    data = {
      'year': parseInt(record[0]),
      'games' : parseInt(record[1]),
      'pa': parseInt(record[2]),
      'ab': parseInt(record[3]),
      'run': parseInt(record[4]),
      'hit': parseInt(record[5]),
      'rbi': parseInt(record[6]),
      'bb': parseInt(record[7]),
      'db': parseInt(record[8]),
      'sh': parseInt(record[9]),
      'sf': parseInt(record[10]),
      'so': parseInt(record[11]),
      'steal': parseInt(record[12]),
      'hit2': parseInt(record[13]),
      'hit3': parseInt(record[14]),
      'hr': parseInt(record[15]),
      'ave' : record[16] === null ? null : parseFloat(record[16]),
      'ave_text' : record[17],
      'obp' : record[18] === null ? null : parseFloat(record[18]),
      'obp_text' : record[19],
      'slg' : record[20] === null ? null : parseFloat(record[20]),
      'slg_text' : record[21],
      'ops' : record[22] === null ? null : parseFloat(record[22]),
      'ops_text': record[23]
    };
    firestore.createDocument('hittingYearTotals', data);
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

async function insertSumHittings(year) {
  const query = `
    #StandardSQL \n
    TRUNCATE TABLE ${datasetId}.sum_hittings_latest_year;
    INSERT INTO ${datasetId}.sum_hittings_latest_year
    (player_id, games, pa, ab, run, hit, rbi, bb, db, sh, sf, so, steal, hit2, hit3, hr, last_name, first_name, player_no, photo_url)
    SELECT
      sum_hittings.*,
      players.last_name,
      players.first_name,
      players.player_no,
      players.photo_url,
    FROM
      (
        SELECT
          player_id,
          COUNT(player_id) games,
          SUM(pa) pa,
          SUM(ab) ab,
          SUM(run) run,
          SUM(hit) hit,
          SUM(rbi) rbi,
          SUM(bb) bb,
          SUM(db) db,
          SUM(sh) sh,
          SUM(sf) sf,
          SUM(so) so,
          SUM(steal) steal,
          SUM(hit2) hit2, 
          SUM(hit3) hit3,
          SUM(hr) hr
        FROM
          ${datasetId}.hittings
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
      ) sum_hittings LEFT JOIN goldpony.players ON sum_hittings.player_id = players.player_id
    ;
  `;

  request = {
    query: query
  };
  BigQuery.Jobs.query(request, bqProjectId);
}

async function insertFirestoreHittingLeaders(year, regulation) {
  await deleteFirestoreDocument('hittingLeaders', null, null);

  // ave
  let query = `
    #StandardSQL \n
    SELECT
      player_id,
      last_name,
      first_name,
      player_no,
      photo_url,
      CASE
        WHEN ab = 0 THEN 
          '---'
        ELSE
          SUBSTR(CAST(format('%.3f', SAFE_DIVIDE(hit, ab)) as string), 2, 4)
      END as ave_text,
      hit / ab as ave
    FROM
      ${datasetId}.sum_hittings_latest_year
    WHERE
      pa >= ${regulation}
    ORDER BY
      ave desc
    LIMIT
      1
  `;

  records = getQueryResult(query);
  records.shift();
  record = records[0];
  let data = {
    'key': 'avg',
    'player_id': record[0],
    'last_name': record[1],
    'first_name': record[2],
    'player_no': record[3],
    'photo_url': record[4],
    'value': record[5]
  };
  firestore.createDocument('hittingLeaders', data);

  // rbi
  query = `
    #StandardSQL \n
    SELECT
      player_id,
      last_name,
      first_name,
      player_no,
      photo_url,
      rbi
    FROM
      ${datasetId}.sum_hittings_latest_year
    ORDER BY
      rbi desc
    LIMIT
      1
  `;

  records = getQueryResult(query);
  records.shift();
  record = records[0];
  data = {
    'key': 'rbi',
    'player_id': record[0],
    'last_name': record[1],
    'first_name': record[2],
    'player_no': record[3],
    'photo_url': record[4],
    'value': record[5]
  };

  firestore.createDocument('hittingLeaders', data);

  // run
  query = `
    #StandardSQL \n
    SELECT
      player_id,
      last_name,
      first_name,
      player_no,
      photo_url,
      run
    FROM
      ${datasetId}.sum_hittings_latest_year
    ORDER BY
      run desc
    LIMIT
      1
  `;

  records = getQueryResult(query);
  records.shift();
  record = records[0];
  data = {
    'key': 'run',
    'player_id': record[0],
    'last_name': record[1],
    'first_name': record[2],
    'player_no': record[3],
    'photo_url': record[4],
    'value': record[5]
  };

  firestore.createDocument('hittingLeaders', data);

  // steal
  query = `
    #StandardSQL \n
    SELECT
      player_id,
      last_name,
      first_name,
      player_no,
      photo_url,
      steal
    FROM
      ${datasetId}.sum_hittings_latest_year
    ORDER BY
      steal desc
    LIMIT
      1
  `;

  records = getQueryResult(query);
  records.shift();
  record = records[0];
  data = {
    'key': 'sb',
    'player_id': record[0],
    'last_name': record[1],
    'first_name': record[2],
    'player_no': record[3],
    'photo_url': record[4],
    'value': record[5]
  };

  firestore.createDocument('hittingLeaders', data);
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

function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('データ登録', [
      {name: '行追加', functionName: 'addRow'},
      {name: '保存', functionName: 'save'},
    ]);
}
