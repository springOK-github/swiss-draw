/**
 * スイス方式トーナメントマッチングシステム
 * @fileoverview 共有ユーティリティ - シート操作とUI共通処理
 * @author springOK
 */

// =========================================
// シート操作ユーティリティ
// =========================================

/**
 * シートの構造を取得し、ヘッダー行の検証を行います。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート
 * @param {string} sheetName シート名（定数から取得）
 * @returns {Object} { headers: 配列, indices: オブジェクト, data: 2次元配列 }
 */

function getSheetStructure(sheet, sheetName) {
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) {
    throw new Error(`シート「${sheetName}」にデータがありません。`);
  }

  const headers = data[0].map(h => String(h).trim());
  const indices = {};
  const missing = [];

  const requiredHeaders = REQUIRED_HEADERS[sheetName];
  if (!requiredHeaders) {
    throw new Error(`シート「${sheetName}」の必須ヘッダー定義が見つかりません。`);
  }

  requiredHeaders.forEach(required => {
    const idx = headers.indexOf(required);
    if (idx === -1) {
      missing.push(required);
    } else {
      indices[required] = idx;
    }
  });

  if (missing.length > 0) {
    throw new Error(`シート「${sheetName}」に必須ヘッダーが不足しています: ${missing.join(", ")}`);
  }

  return { headers, indices, data };
}

/**
 * プレイヤーIDから名前を取得します
 * @param {string} playerId プレイヤーID
 * @returns {string} プレイヤー名。見つからない場合はIDをそのまま返します
 */
function getPlayerName(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

  for (let i = 1; i < data.length; i++) {
    if (data[i][indices["プレイヤーID"]] === playerId) {
      return data[i][indices["プレイヤー名"]] || playerId;
    }
  }
  return playerId;
}

// =========================================
// UI共通ユーティリティ
// =========================================

/**
 * プレイヤーIDの入力を受け付ける共通関数
 * @param {string} title - プロンプトのタイトル
 * @param {string} message - プロンプトのメッセージ
 * @returns {string|null} 整形されたプレイヤーID、キャンセル時はnull
 */
function promptPlayerId(title, message) {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return null;
  }

  const rawId = response.getResponseText().trim();

  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
    return null;
  }

  return PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));
}


