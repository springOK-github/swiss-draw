/**
 * スイス方式トーナメントマッチングシステム
 * @fileoverview プレイヤードメイン - プレイヤーの操作、順位表示、統計管理
 * @author springOK
 */

// =========================================
// プレイヤー操作（UI層）
// =========================================

/**
 * 新しいプレイヤーを登録します。（本番・運営用）
 * 実行すると、次のID（例: P009）が自動で採番され、シートに追加されます。
 */
function registerPlayer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  const ui = SpreadsheetApp.getUi();
  let lock = null;

  try {
    // トーナメントが終了しているかチェック
    const tournamentStatus = getTournamentStatus();
    if (tournamentStatus === TOURNAMENT_STATUS.FINISHED) {
      ui.alert(
        'トーナメント終了済み',
        'このトーナメントは既に終了しています。\n新しいプレイヤーは登録できません。',
        ui.ButtonSet.OK
      );
      return;
    }
    
    lock = acquireLock('プレイヤー登録');
    getSheetStructure(playerSheet, SHEET_PLAYERS);

    const response = ui.prompt(
      'プレイヤー登録',
      'プレイヤー名を入力してください：',
      ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log('プレイヤー登録がキャンセルされました。');
      return;
    }

    // 既存の最大ID番号を取得
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    let maxIdNumber = 0;

    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][indices["プレイヤーID"]];
      if (playerId && playerId.startsWith(PLAYER_ID_PREFIX)) {
        const idNumber = parseInt(playerId.substring(PLAYER_ID_PREFIX.length), 10);
        if (!isNaN(idNumber) && idNumber > maxIdNumber) {
          maxIdNumber = idNumber;
        }
      }
    }

    const newIdNumber = maxIdNumber + 1;
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);

    // プレイヤー名が空の場合はIDを使用
    let playerName = response.getResponseText().trim();
    if (!playerName) {
      playerName = newId;
    }

    // 名前確認ダイアログ
    const confirmResponse = ui.alert(
      '登録確認',
      `プレイヤー名: ${playerName}\nプレイヤーID: ${newId}\n\nこの内容で登録しますか？`,
      ui.ButtonSet.YES_NO);

    if (confirmResponse == ui.Button.NO) {
      Logger.log('プレイヤー登録が確認段階でキャンセルされました。');
      return;
    }

    const currentTime = new Date();
    const formattedTime = Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // スイス方式用: 勝点, 勝数, 敗数, 試合数, OMW%, 参加状況, 最終対戦日時
    playerSheet.appendRow([newId, playerName, 0, 0, 0, 0, 0, PLAYER_STATUS.ACTIVE, ""]);
    Logger.log(`プレイヤー ${newId} を登録しました。`);

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("registerPlayer エラー: " + e.toString());
  }
  finally {
    releaseLock(lock);
  }
}

/**
 * プレイヤーを大会からドロップアウトさせます。
 * 参加状況を「終了」に変更します。
 */
function dropoutPlayer() {
  const ui = SpreadsheetApp.getUi();

  // トーナメントが終了しているかチェック
  const tournamentStatus = getTournamentStatus();
  if (tournamentStatus === TOURNAMENT_STATUS.FINISHED) {
    ui.alert(
      'トーナメント終了済み',
      'このトーナメントは既に終了しています。\nプレイヤーのステータス変更はできません。',
      ui.ButtonSet.OK
    );
    return;
  }

  const playerId = promptPlayerId(
    'プレイヤーのドロップアウト',
    'ドロップアウトするプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。'
  );
  if (!playerId) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  let lock = null;

  try {
    lock = acquireLock('プレイヤードロップアウト');
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    let found = false;
    let playerName = playerId;
    let targetRowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[indices["プレイヤーID"]] === playerId) {
        found = true;
        playerName = row[indices["プレイヤー名"]];
        targetRowIndex = i + 1;
        break;
      }
    }

    if (!found) {
      ui.alert('エラー', `プレイヤー ${playerId} が見つかりません。`, ui.ButtonSet.OK);
      return;
    }

    const confirmResponse = ui.alert(
      'ドロップアウトの確認',
      `プレイヤー名: ${playerName}\nプレイヤーID: ${playerId}\n\nドロップアウトさせます。\n\nよろしいですか？`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    playerSheet.getRange(targetRowIndex, indices["参加状況"] + 1).setValue(PLAYER_STATUS.DROPPED);
    Logger.log(`プレイヤー ${playerId} をドロップアウトさせました。`);

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("dropoutPlayer エラー: " + e.toString());
  } finally {
    releaseLock(lock);
  }
}

// =========================================
// 順位表示
// =========================================

/**
 * プレイヤーの勝率を計算します（対戦相手の平均勝率）
 * @param {string} playerId - プレイヤーID
 * @returns {number} 勝率（0.0～1.0）
 */
function calculateOpponentWinRate(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_HISTORY);
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices: historyIndices, data: historyData } = getSheetStructure(historySheet, SHEET_HISTORY);
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    // このプレイヤーの対戦相手を収集（Byeは除外）
    const opponents = new Set();

    for (let i = 1; i < historyData.length; i++) {
      const row = historyData[i];
      const p1 = row[historyIndices["ID1"]];
      const p2 = row[historyIndices["ID2"]];
      const result = row[historyIndices["結果"]];

      // Byeの場合はスキップ
      if (result === "Bye" || !p2) {
        continue;
      }

      if (p1 === playerId) {
        opponents.add(p2);
      } else if (p2 === playerId) {
        opponents.add(p1);
      }
    }

    if (opponents.size === 0) {
      return 0.333; // 対戦がない場合のデフォルト値
    }

    // 対戦相手の勝率を計算
    let totalWinRate = 0;
    let opponentCount = 0;

    for (const opponentId of opponents) {
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        if (row[playerIndices["プレイヤーID"]] === opponentId) {
          const wins = parseInt(row[playerIndices["勝数"]]) || 0;
          const matches = parseInt(row[playerIndices["試合数"]]) || 0;

          if (matches > 0) {
            const winRate = wins / matches;
            // 最低勝率は0.333（MTGルールに準拠）
            totalWinRate += Math.max(winRate, 0.333);
            opponentCount++;
          }
          break;
        }
      }
    }

    if (opponentCount === 0) {
      return 0.333;
    }

    return totalWinRate / opponentCount;

  } catch (e) {
    Logger.log("calculateOpponentWinRate エラー: " + e.message);
    return 0.333;
  }
}

/**
 * プレイヤーの勝率をシートに更新します
 * @param {string} playerId - プレイヤーID
 */
function updateOpponentWinRate(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    const opponentWinRate = calculateOpponentWinRate(playerId);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[indices["プレイヤーID"]] === playerId) {
        const rowNumber = i + 1;
        playerSheet.getRange(rowNumber, indices["OMW%"] + 1).setValue(opponentWinRate);
        break;
      }
    }
  } catch (e) {
    Logger.log(`updateOpponentWinRate エラー (${playerId}): ` + e.message);
  }
}

/**
 * すべての参加中プレイヤーの勝率を更新します
 */
function updateAllOpponentWinRates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const playerId = row[indices["プレイヤーID"]];
      const status = row[indices["参加状況"]];

      if (status === PLAYER_STATUS.ACTIVE) {
        const opponentWinRate = calculateOpponentWinRate(playerId);
        const rowNumber = i + 1;
        playerSheet.getRange(rowNumber, indices["OMW%"] + 1).setValue(opponentWinRate);
      }
    }

    Logger.log("すべてのプレイヤーの勝率を更新しました。");
  } catch (e) {
    Logger.log("updateAllOpponentWinRates エラー: " + e.message);
  }
}

/**
 * 現在の順位表を表示します
 */
function showStandings() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    // 参加中のプレイヤーのみ抽出
    const activePlayers = data
      .slice(1)
      .filter(row => row[indices["参加状況"]] === PLAYER_STATUS.ACTIVE)
      .map(row => {
        const playerId = row[indices["プレイヤーID"]];
        const matches = parseInt(row[indices["試合数"]]) || 0;
        const wins = parseInt(row[indices["勝数"]]) || 0;
        const opponentWinRate = parseFloat(row[indices["OMW%"]]) || 0;

        return {
          row: row,
          playerId: playerId,
          matchWinRate: matches > 0 ? wins / matches : 0,
          opponentWinRate: opponentWinRate
        };
      })
      .sort((a, b) => {
        // 勝点降順
        const pointsDiff = (b.row[indices["勝点"]] || 0) - (a.row[indices["勝点"]] || 0);
        if (pointsDiff !== 0) return pointsDiff;

        // 勝率降順
        const opponentDiff = b.opponentWinRate - a.opponentWinRate;
        if (Math.abs(opponentDiff) > 0.001) return opponentDiff;

        // 自己勝率降順
        const matchWinDiff = b.matchWinRate - a.matchWinRate;
        if (Math.abs(matchWinDiff) > 0.001) return matchWinDiff;

        // 試合数昇順（少ない方が上位）
        return (a.row[indices["試合数"]] || 0) - (b.row[indices["試合数"]] || 0);
      });

    if (activePlayers.length === 0) {
      ui.alert('順位表', '参加中のプレイヤーがいません。', ui.ButtonSet.OK);
      return;
    }

    let message = '【順位表】\n\n';
    message += '順位 | 名前 | 勝点 | 勝-敗 | OMW% | 試合数\n';
    message += '─'.repeat(50) + '\n';

    for (let i = 0; i < Math.min(activePlayers.length, 20); i++) {
      const player = activePlayers[i];
      const rank = i + 1;
      const name = player.row[indices["プレイヤー名"]];
      const points = player.row[indices["勝点"]] || 0;
      const wins = player.row[indices["勝数"]] || 0;
      const losses = player.row[indices["敗数"]] || 0;
      const matches = player.row[indices["試合数"]] || 0;
      const opponentRate = (player.opponentWinRate * 100).toFixed(1);

      message += `${rank}. ${name} | ${points}pt | ${wins}-${losses} | ${opponentRate}% | ${matches}試合\n`;
    }

    if (activePlayers.length > 20) {
      message += `\n...他 ${activePlayers.length - 20} 人`;
    }

    ui.alert('順位表', message, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("showStandings エラー: " + e.toString());
  }
}