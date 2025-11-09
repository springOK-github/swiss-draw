/**
 * スイス方式トーナメントマッチングシステム
 * @fileoverview 対戦ドメイン - マッチング管理と対戦結果の記録・修正
 * @author springOK
 */

// =========================================
// マッチング管理（スイス方式）
// =========================================

/**
 * スイス方式のマッチングを行います
 * 同じ勝点のプレイヤー同士をマッチングし、再戦を回避します
 * @param {number} roundNumber - ラウンド番号
 * @returns {number} 成立したマッチング数
 */
function matchPlayersSwiss(roundNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  let lock = null;

  try {
    lock = acquireLock('スイス方式マッチング実行');

    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const historySheet = ss.getSheetByName(SHEET_HISTORY);

    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    const { indices: historyIndices, data: historyData } = getSheetStructure(historySheet, SHEET_HISTORY);
    const { indices: matchIndices } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

    // プレイヤー名のマップを作成
    const playerNameMap = new Map();
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      const playerName = row[playerIndices["プレイヤー名"]];
      if (!playerId) continue;
      playerNameMap.set(playerId, playerName);
    }

    // 過去対戦相手のマップを作成（Byeは除外）
    const opponentsMap = new Map();
    const p1Col = historyIndices["ID1"];
    const p2Col = historyIndices["ID2"];
    const resultCol = historyIndices["結果"];

    for (let i = 1; i < historyData.length; i++) {
      const row = historyData[i];
      const p1 = row[p1Col];
      const p2 = row[p2Col];
      const result = row[resultCol];

      // Byeの場合はスキップ（実際の対戦ではないため）
      if (result === "Bye" || !p1 || !p2) continue;

      if (!opponentsMap.has(p1)) opponentsMap.set(p1, new Set());
      if (!opponentsMap.has(p2)) opponentsMap.set(p2, new Set());

      opponentsMap.get(p1).add(p2);
      opponentsMap.get(p2).add(p1);
    }

    // 参加中のプレイヤーを勝点順でソート
    const activePlayers = playerData
      .slice(1)
      .filter(row => row[playerIndices["参加状況"]] === PLAYER_STATUS.ACTIVE)
      .sort((a, b) => {
        // 勝点が多い順（降順）
        const pointsDiff = (b[playerIndices["勝点"]] || 0) - (a[playerIndices["勝点"]] || 0);
        if (pointsDiff !== 0) return pointsDiff;

        // 勝点が同じ場合は、勝数が多い順
        const winsDiff = (b[playerIndices["勝数"]] || 0) - (a[playerIndices["勝数"]] || 0);
        if (winsDiff !== 0) return winsDiff;

        // それでも同じ場合は試合数が少ない順（Byeを受けたプレイヤーを後回し）
        return (a[playerIndices["試合数"]] || 0) - (b[playerIndices["試合数"]] || 0);
      });

    if (activePlayers.length < 2) {
      Logger.log(`警告: 参加中のプレイヤーは ${activePlayers.length} 人です。2人以上必要です。`);
      return 0;
    }

    Logger.log(`--- ラウンド${roundNumber} スイス方式マッチング開始 ---`);
    Logger.log(`参加プレイヤー数: ${activePlayers.length}人`);

    let matches = [];
    let availablePlayers = [...activePlayers];
    let byePlayer = null;

    // 奇数人数の場合、Byeを決定
    if (availablePlayers.length % 2 === 1) {
      // 勝点が最も低いプレイヤーにByeを与える（末尾のプレイヤー）
      byePlayer = availablePlayers.pop();
      const byePlayerId = byePlayer[playerIndices["プレイヤーID"]];
      Logger.log(`Bye: ${byePlayerId} (勝点: ${byePlayer[playerIndices["勝点"]] || 0})`);
    }

    // 勝点グループごとにマッチング（グループ内でランダム化）
    let remainingPlayers = [...availablePlayers];

    // 勝点グループごとにシャッフル
    let currentPoints = null;
    let groupStart = 0;

    for (let i = 0; i <= remainingPlayers.length; i++) {
      const points = i < remainingPlayers.length ? (remainingPlayers[i][playerIndices["勝点"]] || 0) : null;

      if (currentPoints !== null && (points !== currentPoints || i === remainingPlayers.length)) {
        // 現在のグループをシャッフル
        const groupSize = i - groupStart;
        for (let j = groupStart + groupSize - 1; j > groupStart; j--) {
          const k = groupStart + Math.floor(Math.random() * (j - groupStart + 1));
          [remainingPlayers[j], remainingPlayers[k]] = [remainingPlayers[k], remainingPlayers[j]];
        }
        groupStart = i;
      }

      currentPoints = points;
    }

    // 全プレイヤーをマッチング（勝点に関わらず未対戦相手を優先）
    while (remainingPlayers.length >= 2) {
      const p1 = remainingPlayers.shift();
      const p1Id = p1[playerIndices["プレイヤーID"]];
      const p1Opponents = opponentsMap.get(p1Id) || new Set();

      let p2Index = -1;
      for (let i = 0; i < remainingPlayers.length; i++) {
        const p2Id = remainingPlayers[i][playerIndices["プレイヤーID"]];
        if (!p1Opponents.has(p2Id)) {
          p2Index = i;
          break;
        }
      }

      if (p2Index !== -1) {
        const p2 = remainingPlayers.splice(p2Index, 1)[0];
        const p2Id = p2[playerIndices["プレイヤーID"]];
        matches.push([p1Id, p2Id]);
        Logger.log(`マッチング成立: ${p1Id} (勝点${p1[playerIndices["勝点"]] || 0}) vs ${p2Id} (勝点${p2[playerIndices["勝点"]] || 0})`);
      } else {
        // 未対戦相手が見つからない場合
        Logger.log(`警告: ${p1Id} の未対戦相手が見つかりませんでした`);
        break;
      }
    }

    if (remainingPlayers.length > 0) {
      Logger.log(`警告: ${remainingPlayers.length} 人のプレイヤーがマッチングされませんでした`);
    }

    Logger.log(`マッチング成立: ${matches.length}組`);

    // マッチング結果をシートに反映
    let tableNumber = TABLE_CONFIG.MIN_TABLE_NUMBER;
    for (const match of matches) {
      const [p1Id, p2Id] = match;
      inProgressSheet.appendRow([
        roundNumber,
        tableNumber,
        p1Id,
        playerNameMap.get(p1Id) || p1Id,
        p2Id,
        playerNameMap.get(p2Id) || p2Id,
        "" // 結果は空
      ]);
      tableNumber++;
    }

    // Byeの処理
    if (byePlayer) {
      const byePlayerId = byePlayer[playerIndices["プレイヤーID"]];
      const byePlayerName = byePlayer[playerIndices["プレイヤー名"]];

      // Byeを現在のラウンドシートに記録（結果も記録）
      inProgressSheet.appendRow([
        roundNumber,
        tableNumber, // Byeにも卓番号を割り当て
        byePlayerId,
        byePlayerName,
        "",
        "", // 相手名は空欄
        "Bye" // 結果列に記録
      ]);

      // Byeの結果を対戦履歴に即座に記録
      recordByeResult(byePlayerId, roundNumber, tableNumber);
    }

    return matches.length + (byePlayer ? 1 : 0);

  } catch (e) {
    Logger.log("matchPlayersSwiss エラー: " + e.message);
    return 0;
  } finally {
    releaseLock(lock);
  }
}

/**
 * Byeの結果を記録します
 * @param {string} playerId - プレイヤーID
 * @param {number} roundNumber - ラウンド番号
 */
function recordByeResult(playerId, roundNumber, tableNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  const historySheet = ss.getSheetByName(SHEET_HISTORY);

  try {
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    const playerName = getPlayerName(playerId);

    // プレイヤーの統計を更新
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      if (row[playerIndices["プレイヤーID"]] === playerId) {
        const rowNum = i + 1;
        const currentPoints = parseInt(row[playerIndices["勝点"]]) || 0;
        const currentWins = parseInt(row[playerIndices["勝数"]]) || 0;
        const currentTotal = parseInt(row[playerIndices["試合数"]]) || 0;

        playerSheet.getRange(rowNum, playerIndices["勝点"] + 1).setValue(currentPoints + SWISS_CONFIG.POINTS_BYE);
        playerSheet.getRange(rowNum, playerIndices["勝数"] + 1).setValue(currentWins + 1);
        playerSheet.getRange(rowNum, playerIndices["試合数"] + 1).setValue(currentTotal + 1);

        const currentTime = new Date();
        const formattedTime = Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
        playerSheet.getRange(rowNum, playerIndices["最終対戦日時"] + 1).setValue(formattedTime);

        // 対戦履歴に記録
        const newId = "T" + Utilities.formatString("%04d", historySheet.getLastRow());
        historySheet.appendRow([
          newId,
          roundNumber,
          formattedTime,
          tableNumber, // Byeの卓番号を記録
          playerId,
          playerName,
          "",
          "", // 相手名は空欄
          playerName,
          "Bye"
        ]);

        Logger.log(`Bye記録: ${playerId} がラウンド${roundNumber}、卓${tableNumber}でBye`);
        break;
      }
    }
  } catch (e) {
    Logger.log("recordByeResult エラー: " + e.message);
  }
}

// =========================================
// 対戦結果記録
// =========================================

/**
 * カスタムメニューから実行するためのラッパー関数（スイス方式対応）
 */
function promptAndRecordResult() {
  const ui = SpreadsheetApp.getUi();
  const currentRound = getCurrentRound();

  if (currentRound === 0) {
    ui.alert('エラー', 'トーナメントが開始されていません。先にラウンドを開始してください。', ui.ButtonSet.OK);
    return;
  }

  // 結果の種類を選択
  const resultTypeResponse = ui.prompt(
    '対戦結果の記録',
    `結果の種類を選択してください：\n\n` +
    `1: 勝敗（どちらかが勝利）\n` +
    `2: 引き分け（両者敗北扱い、0勝点）\n\n` +
    `数字を入力してください：`,
    ui.ButtonSet.OK_CANCEL
  );

  if (resultTypeResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const resultType = resultTypeResponse.getResponseText().trim();

  if (resultType === '1') {
    // 勝敗の記録
    recordWinLoss();
  } else if (resultType === '2') {
    // 引き分けの記録
    recordDraw();
  } else {
    ui.alert('エラー', '1 または 2 を入力してください。', ui.ButtonSet.OK);
  }
}

/**
 * 勝敗を記録します
 */
function recordWinLoss() {
  const ui = SpreadsheetApp.getUi();

  const winnerResponse = ui.prompt(
    '勝者の入力',
    `勝者のプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。`,
    ui.ButtonSet.OK_CANCEL
  );

  if (winnerResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const rawId = winnerResponse.getResponseText().trim();

  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー', 'IDは数字のみで入力してください。', ui.ButtonSet.OK);
    return;
  }

  const formattedWinnerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
  let loserId = null;
  let matchRow = -1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const p1 = row[indices["ID1"]];
    const p2 = row[indices["ID2"]];

    if (p1 === formattedWinnerId && p2) {
      loserId = p2;
      matchRow = i;
      break;
    } else if (p2 === formattedWinnerId && p1) {
      loserId = p1;
      matchRow = i;
      break;
    }
  }

  if (loserId === null) {
    ui.alert('エラー', `勝者ID (${formattedWinnerId}) は現在のラウンドに見つかりませんでした。`, ui.ButtonSet.OK);
    return;
  }

  const confirmResponse = ui.alert(
    '対戦結果の確認',
    `以下の内容で記録してよろしいですか？\n\n` +
    `勝者: ${getPlayerName(formattedWinnerId)}\n` +
    `敗者: ${getPlayerName(loserId)}`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  try {
    recordMatchResult(formattedWinnerId, loserId, matchRow, 'win');
  } catch (e) {
    ui.alert('エラー', "エラーが発生しました: " + e.toString(), ui.ButtonSet.OK);
    Logger.log("recordWinLoss エラー: " + e.toString());
  }
}

/**
 * 引き分けを記録します
 */
function recordDraw() {
  const ui = SpreadsheetApp.getUi();

  const playerResponse = ui.prompt(
    'プレイヤーの入力',
    `引き分けた対戦のプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。`,
    ui.ButtonSet.OK_CANCEL
  );

  if (playerResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const rawId = playerResponse.getResponseText().trim();

  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー', 'IDは数字のみで入力してください。', ui.ButtonSet.OK);
    return;
  }

  const formattedPlayerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
  let opponentId = null;
  let matchRow = -1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const p1 = row[indices["ID1"]];
    const p2 = row[indices["ID2"]];

    if (p1 === formattedPlayerId && p2) {
      opponentId = p2;
      matchRow = i;
      break;
    } else if (p2 === formattedPlayerId && p1) {
      opponentId = p1;
      matchRow = i;
      break;
    }
  }

  if (opponentId === null) {
    ui.alert('エラー', `プレイヤーID (${formattedPlayerId}) は現在のラウンドに見つかりませんでした。`, ui.ButtonSet.OK);
    return;
  }

  const confirmResponse = ui.alert(
    '引き分けの確認',
    `以下の対戦を引き分けとして記録してよろしいですか？\n\n` +
    `${getPlayerName(formattedPlayerId)} vs ${getPlayerName(opponentId)}`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  try {
    recordMatchResult(formattedPlayerId, opponentId, matchRow, 'draw');
  } catch (e) {
    ui.alert('エラー', "エラーが発生しました: " + e.toString(), ui.ButtonSet.OK);
    Logger.log("recordDraw エラー: " + e.toString());
  }
}

/**
 * 対戦結果を記録します（スイス方式対応）
 * @param {string} player1Id - プレイヤー1のID（勝者または引き分けの一方）
 * @param {string} player2Id - プレイヤー2のID（敗者または引き分けの一方）
 * @param {number} matchRow - 現在のラウンドシートの行番号（0-indexed）
 * @param {string} resultType - 'win'（player1が勝利）または 'draw'（引き分け）
 */
function recordMatchResult(player1Id, player2Id, matchRow, resultType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lock = null;

  try {
    lock = acquireLock('対戦結果の記録');

    const currentRound = getCurrentRound();
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    const { indices: matchIndices, data: matchData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

    const currentTime = new Date();
    const formattedTime = Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // 対戦情報を取得
    const matchRowData = matchData[matchRow];
    const tableNumber = matchRowData[matchIndices["卓番号"]];
    const roundNumber = matchRowData[matchIndices["ラウンド"]];

    let winnerId, loserId, winnerName, loserName, resultText;

    if (resultType === 'win') {
      winnerId = player1Id;
      loserId = player2Id;
      winnerName = getPlayerName(winnerId);
      loserName = getPlayerName(loserId);
      resultText = `${winnerName} 勝利`;
    } else if (resultType === 'draw') {
      winnerId = null;
      loserId = null;
      winnerName = getPlayerName(player1Id);
      loserName = getPlayerName(player2Id);
      resultText = '両負け';
    }

    // 現在のラウンドシートに結果を記録
    inProgressSheet.getRange(matchRow + 1, matchIndices["結果"] + 1).setValue(resultText);

    // 対戦履歴に記録
    const newId = "T" + Utilities.formatString("%04d", historySheet.getLastRow());
    historySheet.appendRow([
      newId,
      roundNumber,
      formattedTime,
      tableNumber,
      player1Id,
      resultType === 'win' ? winnerName : getPlayerName(player1Id),
      player2Id,
      resultType === 'win' ? loserName : getPlayerName(player2Id),
      resultType === 'win' ? winnerName : '',
      resultText
    ]);

    // プレイヤーの統計を更新
    updatePlayerStats(player1Id, resultType === 'win' ? 'win' : 'draw', formattedTime);
    updatePlayerStats(player2Id, resultType === 'win' ? 'loss' : 'draw', formattedTime);

    Logger.log(`対戦結果記録: ${player1Id} vs ${player2Id}, 結果: ${resultText}`);

  } catch (e) {
    Logger.log("recordMatchResult エラー: " + e.message);
    throw e;
  } finally {
    releaseLock(lock);
  }
}

/**
 * プレイヤーの統計情報を更新します（スイス方式対応）
 * @param {string} playerId - プレイヤーID
 * @param {string} result - 'win', 'loss', 'draw'
 * @param {string} timestamp - タイムスタンプ
 */
function updatePlayerStats(playerId, result, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[indices["プレイヤーID"]] === playerId) {
        const rowNum = i + 1;
        const currentPoints = parseInt(row[indices["勝点"]]) || 0;
        const currentWins = parseInt(row[indices["勝数"]]) || 0;
        const currentLosses = parseInt(row[indices["敗数"]]) || 0;
        const currentTotal = parseInt(row[indices["試合数"]]) || 0;

        let pointsToAdd = 0;
        let winsToAdd = 0;
        let lossesToAdd = 0;

        if (result === 'win') {
          pointsToAdd = SWISS_CONFIG.POINTS_WIN;
          winsToAdd = 1;
        } else if (result === 'loss') {
          pointsToAdd = SWISS_CONFIG.POINTS_LOSS;
          lossesToAdd = 1;
        } else if (result === 'draw') {
          // 引き分けは両者敗北扱い（0勝点）
          pointsToAdd = SWISS_CONFIG.POINTS_DRAW;
          lossesToAdd = 1;
        }

        playerSheet.getRange(rowNum, indices["勝点"] + 1).setValue(currentPoints + pointsToAdd);
        playerSheet.getRange(rowNum, indices["勝数"] + 1).setValue(currentWins + winsToAdd);
        playerSheet.getRange(rowNum, indices["敗数"] + 1).setValue(currentLosses + lossesToAdd);
        playerSheet.getRange(rowNum, indices["試合数"] + 1).setValue(currentTotal + 1);
        playerSheet.getRange(rowNum, indices["最終対戦日時"] + 1).setValue(timestamp);

        return;
      }
    }
    Logger.log(`エラー: プレイヤー ${playerId} が見つかりません。`);
  } catch (e) {
    Logger.log("updatePlayerStats エラー: " + e.message);
  }
}

// =========================================
// 対戦結果修正系
// =========================================

/**
 * 対戦結果の勝敗を修正します。
 * 対戦IDを指定して、勝者と敗者を入れ替えます。
 * 両プレイヤーの統計情報（勝数・敗数）も自動的に調整されます。
 */
function correctMatchResult() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lock = null;

  try {
    // 1. 対戦IDの入力
    const response = ui.prompt(
      '対戦結果の修正',
      '修正する対戦IDの**数字部分のみ**を入力してください (例: T0001なら「1」)。',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    const rawId = response.getResponseText().trim();
    if (!/^\d+$/.test(rawId)) {
      ui.alert('エラー', 'IDは数字のみで入力してください。', ui.ButtonSet.OK);
      return;
    }

    const matchId = "T" + Utilities.formatString("%04d", parseInt(rawId, 10));

    lock = acquireLock('対戦結果の修正');

    // 2. 対戦履歴から該当の対戦を検索
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    const { indices: historyIndices, data: historyData } = getSheetStructure(historySheet, SHEET_HISTORY);

    let matchRow = -1;
    let matchData = null;

    for (let i = 1; i < historyData.length; i++) {
      const row = historyData[i];
      if (row[historyIndices["対戦ID"]] === matchId) {
        matchRow = i + 1;
        matchData = row;
        break;
      }
    }

    if (matchRow === -1) {
      ui.alert('エラー', `対戦ID ${matchId} が見つかりません。`, ui.ButtonSet.OK);
      return;
    }

    // 3. 現在の勝者と敗者を取得
    const currentWinnerId = matchData[historyIndices["ID1"]];
    const currentWinnerName = matchData[historyIndices["プレイヤー1"]];
    const currentLoserId = matchData[historyIndices["ID2"]];
    const currentLoserName = matchData[historyIndices["プレイヤー2"]];

    // 4. 修正の確認
    const confirmResponse = ui.alert(
      '勝敗修正の確認',
      `対戦ID: ${matchId}\n\n` +
      `【現在】\n` +
      `勝者: ${currentWinnerName} (${currentWinnerId})\n` +
      `敗者: ${currentLoserName} (${currentLoserId})\n\n` +
      `【修正後】\n` +
      `勝者: ${currentLoserName} (${currentLoserId})\n` +
      `敗者: ${currentWinnerName} (${currentWinnerId})\n\n` +
      '勝敗を入れ替えますか？',
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    // 5. 対戦履歴を更新（勝者と敗者を入れ替え）
    historySheet.getRange(matchRow, historyIndices["ID1"] + 1).setValue(currentLoserId);
    historySheet.getRange(matchRow, historyIndices["プレイヤー1"] + 1).setValue(currentLoserName);
    historySheet.getRange(matchRow, historyIndices["ID2"] + 1).setValue(currentWinnerId);
    historySheet.getRange(matchRow, historyIndices["プレイヤー2"] + 1).setValue(currentWinnerName);
    historySheet.getRange(matchRow, historyIndices["勝者名"] + 1).setValue(currentLoserName);

    // 6. プレイヤーの統計を修正
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      const rowNum = i + 1;

      if (playerId === currentWinnerId) {
        // 元の勝者: 勝数-1、敗数+1
        const currentWins = parseInt(row[playerIndices["勝数"]]) || 0;
        const currentLosses = parseInt(row[playerIndices["敗数"]]) || 0;
        playerSheet.getRange(rowNum, playerIndices["勝数"] + 1).setValue(Math.max(0, currentWins - 1));
        playerSheet.getRange(rowNum, playerIndices["敗数"] + 1).setValue(currentLosses + 1);
      } else if (playerId === currentLoserId) {
        // 元の敗者: 勝数+1、敗数-1
        const currentWins = parseInt(row[playerIndices["勝数"]]) || 0;
        const currentLosses = parseInt(row[playerIndices["敗数"]]) || 0;
        playerSheet.getRange(rowNum, playerIndices["勝数"] + 1).setValue(currentWins + 1);
        playerSheet.getRange(rowNum, playerIndices["敗数"] + 1).setValue(Math.max(0, currentLosses - 1));
      }
    }

    Logger.log(`対戦結果修正完了: ${matchId}, 新勝者: ${currentLoserId}, 新敗者: ${currentWinnerId}`);

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("correctMatchResult エラー: " + e.toString());
  } finally {
    releaseLock(lock);
  }
}
