/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview スプレッドシートの初期設定とメニュー関連の機能
 * @author SpringOK
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを作成します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🃏 ポケモンマッチング')
    .addItem('シートの初期設定', 'setupSheets')
    .addSeparator()
    .addItem('プレイヤー登録', 'registerPlayer')
    .addItem('対戦結果の記録', 'promptAndRecordResult')
    .addSeparator()
    .addItem('プレイヤーを休憩にする', 'setPlayerResting')
    .addItem('休憩から復帰させる', 'returnPlayerFromResting')
    .addSeparator()
    .addItem('プレイヤーをドロップアウトさせる', 'dropoutPlayer')
    .addSeparator()
    .addItem('⚙️ 最大卓数の設定', 'configureMaxTables')
    .addToUi();
}

/**
 * スプレッドシートを初期化し、必要なシートとヘッダーを作成します。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

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
  playerSheet.setColumnWidth(1, 100);
  playerSheet.setColumnWidth(5, 100);
  playerSheet.setColumnWidth(6, 150);

  // 2. 対戦履歴シート
  let historySheet = ss.getSheetByName(SHEET_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_HISTORY);
  }
  historySheet.clear();
  const historyHeaders = REQUIRED_HEADERS[SHEET_HISTORY];
  historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders])
    .setFontWeight("bold").setBackground("#fce5cd").setHorizontalAlignment("center");
  historySheet.setColumnWidth(1, 150);

  // 3. マッチングシート
  let inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  if (!inProgressSheet) {
    inProgressSheet = ss.insertSheet(SHEET_IN_PROGRESS);
  }
  inProgressSheet.clear();
  const inProgressHeaders = REQUIRED_HEADERS[SHEET_IN_PROGRESS];
  inProgressSheet.getRange(1, 1, 1, inProgressHeaders.length).setValues([inProgressHeaders])
    .setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  inProgressSheet.setColumnWidth(3, 80);

  Logger.log("シートの初期設定が完了しました。");
}

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

/**
 * テーブル番号の妥当性を検証します。
 * @param {number} tableNumber - 検証するテーブル番号
 * @returns {boolean} 妥当な場合はtrue、そうでない場合はfalse
 */
function validateTableNumber(tableNumber) {
  const maxTables = getMaxTables();
  // ... maxTables を使って検証
}