/**
 * スイス方式トーナメントマッチングシステム
 * @fileoverview アプリケーション層 - 初期化・設定・排他制御
 * @author springOK
 */

// =========================================
// システム初期化・メニュー
// =========================================

/**
 * スプレッドシートを開いたときにカスタムメニューを作成します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 スイス方式トーナメント')
    .addItem('⚙️ シートの初期設定', 'setupSheets')
    .addSeparator()
    .addItem('🎯 新ラウンド開始', 'startNewRoundUI')
    .addItem('🏁 トーナメント終了', 'finishTournament')
    .addItem('🏅 順位表示', 'showStandings')
    .addSeparator()
    .addItem('➕ プレイヤーを追加する', 'registerPlayer')
    .addItem('❌ プレイヤーをドロップアウトさせる', 'dropoutPlayer')
    .addSeparator()
    .addItem('✅ 対戦結果の記録', 'promptAndRecordResult')
    .addItem('🔧 対戦結果の修正', 'correctMatchResult')
    .addSeparator()
    .addItem('⚙️ 最大卓数の設定', 'configureMaxTables')
    .addToUi();
}

/**
 * スプレッドシートを初期化し、必要なシートとヘッダーを作成します。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. プレイヤーシート
  let playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  if (!playerSheet) {
    playerSheet = ss.insertSheet(SHEET_PLAYERS);
  }
  playerSheet.clear();
  const playerHeaders = REQUIRED_HEADERS[SHEET_PLAYERS];
  playerSheet.getRange(1, 1, 1, playerHeaders.length).setValues([playerHeaders])
    .setFontWeight("bold").setBackground("#c9daf8").setHorizontalAlignment("center");
  // 幅の調整
  playerSheet.setColumnWidth(1, 100);  // プレイヤーID
  playerSheet.setColumnWidth(2, 150);  // プレイヤー名
  playerSheet.setColumnWidth(3, 60);   // 勝点
  playerSheet.setColumnWidth(7, 100);  // 勝率
  playerSheet.setColumnWidth(8, 100);  // 参加状況
  playerSheet.setColumnWidth(9, 150);  // 最終対戦日時

  // 2. 対戦履歴シート
  let historySheet = ss.getSheetByName(SHEET_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_HISTORY);
  }
  historySheet.clear();
  const historyHeaders = REQUIRED_HEADERS[SHEET_HISTORY];
  historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders])
    .setFontWeight("bold").setBackground("#fce5cd").setHorizontalAlignment("center");
  historySheet.setColumnWidth(1, 100);  // 対戦ID
  historySheet.setColumnWidth(2, 80);   // ラウンド
  historySheet.setColumnWidth(3, 150);  // 日時

  // 3. 現在のラウンドシート
  let inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  if (!inProgressSheet) {
    inProgressSheet = ss.insertSheet(SHEET_IN_PROGRESS);
  }
  inProgressSheet.clear();
  const inProgressHeaders = REQUIRED_HEADERS[SHEET_IN_PROGRESS];
  inProgressSheet.getRange(1, 1, 1, inProgressHeaders.length).setValues([inProgressHeaders])
    .setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  inProgressSheet.setColumnWidth(1, 80);   // ラウンド
  inProgressSheet.setColumnWidth(2, 80);   // 卓番号
  inProgressSheet.setColumnWidth(7, 150);  // 結果

  // ラウンド番号を初期化
  setCurrentRound(0);

  // トーナメント状態を進行中に初期化
  setTournamentStatus(TOURNAMENT_STATUS.IN_PROGRESS);

  ui.alert('初期設定完了', 'シートの初期設定が完了しました。\n\n「新ラウンド開始」からトーナメントを開始できます。', ui.ButtonSet.OK);
  Logger.log("シートの初期設定が完了しました。");
}

// =========================================
// システム設定管理
// =========================================

/**
 * 現在の最大卓数を取得します。
 * PropertiesServiceに保存されている値、なければデフォルト値を返します。
 * @returns {number} 最大卓数
 */
function getMaxTables() {
  const properties = PropertiesService.getDocumentProperties();
  const savedMaxTables = properties.getProperty('MAX_TABLES');

  if (savedMaxTables) {
    return parseInt(savedMaxTables, 10);
  }

  // デフォルト値
  return TABLE_CONFIG.MAX_TABLES;
}

/**
 * 最大卓数を設定します。
 * @param {number} maxTables - 設定する最大卓数
 */
function setMaxTables(maxTables) {
  const properties = PropertiesService.getDocumentProperties();
  properties.setProperty('MAX_TABLES', maxTables.toString());
  Logger.log(`最大卓数を ${maxTables} に設定しました。`);
}

/**
 * 最大卓数の設定をユーザーに促すダイアログを表示します。
 */
function configureMaxTables() {
  const ui = SpreadsheetApp.getUi();
  const currentMaxTables = getMaxTables();
  const currentRound = getCurrentRound();

  // トーナメント開始後は変更不可
  if (currentRound > 0) {
    ui.alert(
      '設定変更不可',
      `トーナメント開始後は卓数を変更できません。\n\n` +
      `現在のラウンド: ${currentRound}\n` +
      `現在の最大卓数: ${currentMaxTables}卓`,
      ui.ButtonSet.OK
    );
    return;
  }

  const response = ui.prompt(
    '最大卓数の設定',
    `現在の最大卓数: ${currentMaxTables}卓\n\n` +
    `新しい最大卓数を入力してください（1～200）：`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('設定をキャンセルしました。');
    return;
  }

  const input = response.getResponseText().trim();

  // 入力検証
  if (!/^\d+$/.test(input)) {
    ui.alert('エラー', '数字のみで入力してください。', ui.ButtonSet.OK);
    return;
  }

  const newMaxTables = parseInt(input, 10);

  // 範囲検証
  if (newMaxTables < 1 || newMaxTables > 200) {
    ui.alert('エラー', '最大卓数は1～200の範囲で入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 確認ダイアログ
  const confirmResponse = ui.alert(
    '設定の確認',
    `最大卓数を ${currentMaxTables}卓 → ${newMaxTables}卓 に変更します。\n\n` +
    'よろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('設定をキャンセルしました。');
    return;
  }

  // 設定を保存
  setMaxTables(newMaxTables);

  ui.alert(
    '設定完了',
    `最大卓数を ${newMaxTables}卓 に設定しました。`,
    ui.ButtonSet.OK
  );
}

// =========================================
// 排他制御
// =========================================

// ロックの最大待機時間（ミリ秒）
const LOCK_TIMEOUT = 30000; // 30秒

/**
 * スプレッドシートの排他ロックを取得します。
 * @param {string} lockName - ロックの名前（操作の種類を識別）
 * @returns {LockService.Lock} 取得したロック
 * @throws {Error} ロックが取得できない場合
 */
function acquireLock(lockName) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(LOCK_TIMEOUT);

  if (!success) {
    throw new Error(
      '他のユーザーが操作中です。\n' +
      'しばらく待ってから再度お試しください。\n' +
      `(${lockName})`
    );
  }

  return lock;
}

/**
 * ロックを解放します。
 * @param {LockService.Lock} lock - 解放するロック
 */
function releaseLock(lock) {
  if (lock) {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log('ロックの解放に失敗: ' + e.toString());
    }
  }
}