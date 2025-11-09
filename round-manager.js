/**
 * スイス方式トーナメントマッチングシステム
 * @fileoverview ラウンド管理 - ラウンドのライフサイクル管理
 * @author springOK
 */

// =========================================
// ラウンド管理
// =========================================

/**
 * 現在のラウンド番号を取得します
 * @returns {number} 現在のラウンド番号（0の場合は未開始）
 */
function getCurrentRound() {
    const properties = PropertiesService.getDocumentProperties();
    const currentRound = properties.getProperty('CURRENT_ROUND');
    return currentRound ? parseInt(currentRound, 10) : 0;
}

/**
 * 現在のラウンド番号を設定します
 * @param {number} roundNumber - 設定するラウンド番号
 */
function setCurrentRound(roundNumber) {
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperty('CURRENT_ROUND', roundNumber.toString());
    Logger.log(`現在のラウンドを ${roundNumber} に設定しました。`);
}

/**
 * 現在のラウンドが終了しているかチェックします
 * @returns {boolean} すべての対戦結果が記録されている場合true
 */
function isRoundComplete() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

    try {
        const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

        // データ行がない場合は完了とみなす
        if (data.length <= 1) {
            return true;
        }

        // すべての対戦に結果が記録されているかチェック
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const id1 = row[indices["ID1"]];
            const id2 = row[indices["ID2"]];
            const result = row[indices["結果"]];

            // 対戦が存在するのに結果が記録されていない場合
            if (id1 && !result) {
                return false;
            }
        }

        return true;
    } catch (e) {
        Logger.log("isRoundComplete エラー: " + e.message);
        return false;
    }
}

/**
 * 新しいラウンドを開始します
 * @returns {Object} { success: boolean, message: string, round: number }
 */
function startNewRound() {
    const ui = SpreadsheetApp.getUi();
    let lock = null;

    try {
        lock = acquireLock('新ラウンド開始');

        // 現在のラウンドが終了しているかチェック
        const currentRound = getCurrentRound();

        if (currentRound > 0 && !isRoundComplete()) {
            return {
                success: false,
                message: `ラウンド${currentRound}が終了していません。すべての対戦結果を記録してください。`
            };
        }

        // 参加中のプレイヤー数をチェック
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
        const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

        const activePlayers = playerData.slice(1).filter(row =>
            row[playerIndices["参加状況"]] === PLAYER_STATUS.ACTIVE
        );

        if (activePlayers.length < 2) {
            return {
                success: false,
                message: `参加中のプレイヤーが${activePlayers.length}人しかいません。2人以上必要です。`
            };
        }

        // 新しいラウンド番号
        const newRound = currentRound + 1;

        // 現在のラウンドシートをクリア
        const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
        if (inProgressSheet.getLastRow() > 1) {
            inProgressSheet.getRange(2, 1, inProgressSheet.getLastRow() - 1, inProgressSheet.getLastColumn()).clearContent();
        }

        // ラウンド番号を更新
        setCurrentRound(newRound);

        // マッチングを実行
        const matchCount = matchPlayersSwiss(newRound);

        if (matchCount === 0) {
            return {
                success: false,
                message: 'マッチングに失敗しました。'
            };
        }

        return {
            success: true,
            message: `ラウンド${newRound}を開始しました。${matchCount}組のマッチングが成立しました。`,
            round: newRound
        };

    } catch (e) {
        Logger.log("startNewRound エラー: " + e.message);
        return {
            success: false,
            message: "エラーが発生しました: " + e.toString()
        };
    } finally {
        releaseLock(lock);
    }
}

/**
 * ラウンド開始のUIラッパー関数
 */
function startNewRoundUI() {
    const ui = SpreadsheetApp.getUi();
    const currentRound = getCurrentRound();

    const confirmResponse = ui.alert(
        '新ラウンド開始',
        `現在: ラウンド${currentRound}\n\n` +
        `ラウンド${currentRound + 1}を開始しますか？`,
        ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
        ui.alert('処理をキャンセルしました。');
        return;
    }

    const result = startNewRound();

    if (result.success) {
        ui.alert('ラウンド開始', result.message, ui.ButtonSet.OK);
    } else {
        ui.alert('エラー', result.message, ui.ButtonSet.OK);
    }
}

/**
 * 現在のラウンド状況を表示します
 */
function showRoundStatus() {
    const ui = SpreadsheetApp.getUi();
    const currentRound = getCurrentRound();

    if (currentRound === 0) {
        ui.alert('ラウンド状況', 'トーナメントは未開始です。', ui.ButtonSet.OK);
        return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
    const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

    let totalMatches = 0;
    let completedMatches = 0;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const id1 = row[indices["ID1"]];
        const result = row[indices["結果"]];

        if (id1) {
            totalMatches++;
            if (result) {
                completedMatches++;
            }
        }
    }

    const isComplete = isRoundComplete();
    const status = isComplete ? '✅ 完了' : '⏳ 進行中';

    ui.alert(
        'ラウンド状況',
        `現在のラウンド: ${currentRound}\n` +
        `状態: ${status}\n\n` +
        `対戦数: ${completedMatches} / ${totalMatches}`,
        ui.ButtonSet.OK
    );
}

/**
 * トーナメントをリセットします（開発・テスト用）
 */
function resetTournament() {
    const ui = SpreadsheetApp.getUi();

    const confirmResponse = ui.alert(
        'トーナメントのリセット',
        '⚠️ すべてのラウンドデータと対戦履歴が削除されます。\n' +
        'プレイヤー情報はリセットされます。\n\n' +
        '本当に実行しますか？',
        ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
        ui.alert('処理をキャンセルしました。');
        return;
    }

    let lock = null;

    try {
        lock = acquireLock('トーナメントリセット');

        // ラウンド番号をリセット
        setCurrentRound(0);

        // プレイヤーシートの統計をリセット
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
        const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

        for (let i = 1; i < playerData.length; i++) {
            const rowNum = i + 1;
            playerSheet.getRange(rowNum, playerIndices["勝点"] + 1).setValue(0);
            playerSheet.getRange(rowNum, playerIndices["勝数"] + 1).setValue(0);
            playerSheet.getRange(rowNum, playerIndices["敗数"] + 1).setValue(0);
            playerSheet.getRange(rowNum, playerIndices["消化試合数"] + 1).setValue(0);
            playerSheet.getRange(rowNum, playerIndices["参加状況"] + 1).setValue(PLAYER_STATUS.ACTIVE);
        }

        // 対戦履歴シートをクリア
        const historySheet = ss.getSheetByName(SHEET_HISTORY);
        if (historySheet.getLastRow() > 1) {
            historySheet.getRange(2, 1, historySheet.getLastRow() - 1, historySheet.getLastColumn()).clearContent();
        }

        // 現在のラウンドシートをクリア
        const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
        if (inProgressSheet.getLastRow() > 1) {
            inProgressSheet.getRange(2, 1, inProgressSheet.getLastRow() - 1, inProgressSheet.getLastColumn()).clearContent();
        }

        ui.alert('リセット完了', 'トーナメントをリセットしました。', ui.ButtonSet.OK);
        Logger.log("トーナメントをリセットしました。");

    } catch (e) {
        ui.alert("エラーが発生しました: " + e.toString());
        Logger.log("resetTournament エラー: " + e.toString());
    } finally {
        releaseLock(lock);
    }
}
