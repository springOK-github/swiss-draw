/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview マッチングの管理を行うマネージャー
 * @author SpringOK
 */

// =========================================
// マッチング系
// =========================================

/**
 * 待機中のプレイヤーを抽出し、再戦履歴を厳格に考慮してマッチングを行います。
 * 過去に対戦した相手しかいない場合、マッチングを成立させずに待機させます。
 */
function matchPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  let lock = null;

  try {
    lock = acquireLock('マッチング実行');
    getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    const waitingPlayers = getWaitingPlayers();
    if (waitingPlayers.length < 2) {
      Logger.log(`警告: 現在待機中のプレイヤーは ${waitingPlayers.length} 人です。2人以上必要です。`);
      return;
    }

    let matches = [];
    let availablePlayers = [...waitingPlayers];
    let skippedPlayers = [];

    Logger.log("--- 厳格な再戦回避マッチング開始 (勝者優先) ---");
    while (availablePlayers.length >= 2) {
      const p1 = availablePlayers.shift();
      const p1Id = p1[playerIndices["プレイヤーID"]];
      const p1BlackList = getPastOpponents(p1Id);

      let p2Index = -1;
      for (let i = 0; i < availablePlayers.length; i++) {
        const p2Id = availablePlayers[i][playerIndices["プレイヤーID"]];
        if (!p1BlackList.includes(p2Id)) {
          p2Index = i;
          break;
        }
      }

      if (p2Index !== -1) {
        const p2 = availablePlayers.splice(p2Index, 1)[0];
        matches.push([p1Id, p2[playerIndices["プレイヤーID"]]]);
        Logger.log(`マッチング成立 (再戦なし): ${p1Id} vs ${p2[playerIndices["プレイヤーID"]]}`);
      } else {
        skippedPlayers.push(p1);
      }
    }

    skippedPlayers.push(...availablePlayers);

    if (skippedPlayers.length > 0) {
      Logger.log(`警告: ${skippedPlayers.length} 人のプレイヤーは適切な相手が見つからなかったため、待機を継続します。`);
    }

    // マッチング結果の反映
    if (matches.length > 0) {
      // プレイヤーの状態を更新
      const playerIdsToUpdate = matches.flat();
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        const playerId = row[playerIndices["プレイヤーID"]];
        if (playerIdsToUpdate.includes(playerId)) {
          playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
            .setValue(PLAYER_STATUS.IN_PROGRESS);
        }
      }

      // マッチングシートの現在の状態を取得
      const { indices: matchIndices, data: matchData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

      // 既存の卓の状態を確認
      const availableTables = [];
      const usedTables = new Set();

      for (let i = 1; i < matchData.length; i++) {
        const row = matchData[i];
        const tableNumber = row[matchIndices["卓番号"]];
        if (!tableNumber) continue;

        if (!row[matchIndices["ID1"]]) {
          // 空いている卓
          availableTables.push({ row: i, tableNumber: tableNumber });
        } else {
          usedTables.add(tableNumber);
        }
      }

      // マッチを卓に割り当て
      for (const match of matches) {
        const [p1Id, p2Id] = match;
        let targetRow = null;
        let tableNumber = null;

        // 勝者の前回使用した卓を確認
        const lastTable = getLastTableNumber(p1Id);
        if (lastTable) {
          const validation = validateTableNumber(lastTable);
          if (validation.isValid && !usedTables.has(lastTable)) {
            // 前回の卓が有効で空いている場合、その卓を使用
            const availableTableIndex = availableTables.findIndex(t => t.tableNumber === lastTable);
            if (availableTableIndex !== -1) {
              const table = availableTables.splice(availableTableIndex, 1)[0];
              targetRow = table.row;
              tableNumber = table.tableNumber;
              usedTables.add(tableNumber);
            }
          }
        }

        // 前回の卓が使えない場合は新しい卓を割り当て
        if (targetRow === null) {
          if (availableTables.length > 0) {
            // 既存の空き卓から割り当て
            const table = availableTables.shift();
            targetRow = table.row;
            tableNumber = table.tableNumber;
          } else {
            // 新しい卓を作成
            const newTableNumber = getNextAvailableTableNumber(inProgressSheet);
            tableNumber = newTableNumber;
            targetRow = matchData.length;
            // 新しい卓番号を設定
            inProgressSheet.getRange(targetRow + 1, 1).setValue(tableNumber);
          }
          usedTables.add(tableNumber);
        }

        // マッチを卓に割り当て
        inProgressSheet.getRange(targetRow + 1, 2, 1, 4).setValues([[
          p1Id,
          getPlayerName(p1Id),
          p2Id,
          getPlayerName(p2Id)
        ]]);
      }

      Logger.log(`マッチングが ${matches.length} 件成立しました。「${SHEET_IN_PROGRESS}」シートを確認してください。`);
      return matches.length;
    } else {
      Logger.log("警告: 新しいマッチングは成立しませんでした。");
      return 0;
    }

  } catch (e) {
    Logger.log("matchPlayers エラー: " + e.message);
    return 0;
  } finally {
    releaseLock(lock);
  }
}
