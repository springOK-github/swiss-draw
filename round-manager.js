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

        // 勝率を更新（ラウンド2以降）
        if (newRound > 1) {
            updateAllOpponentWinRates();
        }

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

    // トーナメントが終了しているかチェック
    const tournamentStatus = getTournamentStatus();
    if (tournamentStatus === TOURNAMENT_STATUS.FINISHED) {
        ui.alert(
            'トーナメント終了済み',
            'このトーナメントは既に終了しています。\n新しいラウンドは開始できません。',
            ui.ButtonSet.OK
        );
        return;
    }

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

    if (!result.success) {
        ui.alert('エラー', result.message, ui.ButtonSet.OK);
    }
}

/**
 * トーナメントの状態を取得します
 * @returns {string} トーナメントの状態（進行中 or 終了）
 */
function getTournamentStatus() {
    const properties = PropertiesService.getDocumentProperties();
    const status = properties.getProperty('TOURNAMENT_STATUS');
    return status || TOURNAMENT_STATUS.IN_PROGRESS;
}

/**
 * トーナメントの状態を設定します
 * @param {string} status - 設定する状態
 */
function setTournamentStatus(status) {
    const properties = PropertiesService.getDocumentProperties();
    properties.setProperty('TOURNAMENT_STATUS', status);
    Logger.log(`トーナメント状態を ${status} に設定しました。`);
}

/**
 * トーナメントを終了します
 * - OMW%を最終更新
 * - トーナメント状態を「終了」に設定
 * - 以降のラウンド開始を禁止
 */
function finishTournament() {
    const ui = SpreadsheetApp.getUi();
    let lock = null;

    try {
        const status = getTournamentStatus();
        
        // 既に終了している場合はOMW%再計算のみ実行
        if (status === TOURNAMENT_STATUS.FINISHED) {
            const confirmResponse = ui.alert(
                'OMW%再計算',
                'トーナメントは既に終了しています。\n\n' +
                'OMW%を再計算しますか？\n' +
                '（対戦結果を修正した後に実行してください）',
                ui.ButtonSet.YES_NO
            );
            
            if (confirmResponse !== ui.Button.YES) {
                ui.alert('処理をキャンセルしました。');
                return;
            }
            
            lock = acquireLock('OMW%再計算');
            
            // OMW%を再計算
            updateAllOpponentWinRates();
            
            ui.alert(
                'OMW%再計算完了',
                'OMW%の再計算が完了しました。\n\n' +
                '最新の順位は「🏅 順位表示」から確認できます。',
                ui.ButtonSet.OK
            );
            return;
        }

        // 現在のラウンドが完了しているかチェック
        if (!isRoundComplete()) {
            ui.alert(
                'ラウンド未完了',
                '現在のラウンドが完了していません。\nすべての対戦結果を記録してから終了してください。',
                ui.ButtonSet.OK
            );
            return;
        }

        const confirmResponse = ui.alert(
            'トーナメント終了確認',
            'トーナメントを終了しますか？\n\n' +
            '終了後は新しいラウンドを開始できなくなります。\n' +
            'OMW%が最終更新されます。',
            ui.ButtonSet.YES_NO
        );

        if (confirmResponse !== ui.Button.YES) {
            ui.alert('処理をキャンセルしました。');
            return;
        }

        lock = acquireLock('トーナメント終了');

        // OMW%を最終更新
        updateAllOpponentWinRates();

        // トーナメント状態を終了に設定
        setTournamentStatus(TOURNAMENT_STATUS.FINISHED);

        ui.alert(
            'トーナメント終了',
            'トーナメントが正常に終了しました。\n\n' +
            '最終順位は「🏅 順位表示」から確認できます。',
            ui.ButtonSet.OK
        );

    } catch (e) {
        Logger.log("finishTournament エラー: " + e.message);
        ui.alert('エラー', 'トーナメント終了中にエラーが発生しました: ' + e.message, ui.ButtonSet.OK);
    } finally {
        releaseLock(lock);
    }
}
