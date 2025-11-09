/**
 * スイス方式トーナメントマッチングシステム
 * @fileoverview システム定数 - シート名、ステータス、卓設定
 * @author springOK
 */

const SHEET_PLAYERS = "プレイヤー";
const SHEET_HISTORY = "対戦履歴";
const SHEET_IN_PROGRESS = "現在のラウンド";
const PLAYER_ID_PREFIX = "P";
const ID_DIGITS = 3; // IDの数字部分の桁数 (例: P001なら3)
const PLAYER_STATUS = {
  ACTIVE: "参加中",
  DROPPED: "終了"
};
// 卓に関する設定
const TABLE_CONFIG = {
  MAX_TABLES: 50,      // デフォルトの最大卓数（PropertiesServiceで上書き可能、範囲: 1-200）
  MIN_TABLE_NUMBER: 1  // 最小卓番号
};

// スイス方式の設定
const SWISS_CONFIG = {
  POINTS_WIN: 3,       // 勝利時の勝点
  POINTS_DRAW: 0,      // 引き分け時の勝点（敗北と同じ扱い）
  POINTS_LOSS: 0,      // 敗北時の勝点
  POINTS_BYE: 3        // バイ（不戦勝）時の勝点
};

const REQUIRED_HEADERS = {
  [SHEET_PLAYERS]: ["プレイヤーID", "プレイヤー名", "勝点", "勝数", "敗数", "試合数", "OMW%", "参加状況", "最終対戦日時"],
  [SHEET_HISTORY]: ["対戦ID", "ラウンド", "日時", "卓番号", "ID1", "プレイヤー1", "ID2", "プレイヤー2", "勝者名", "結果"],
  [SHEET_IN_PROGRESS]: ["ラウンド", "卓番号", "ID1", "プレイヤー1", "ID2", "プレイヤー2", "結果"]
};