# スイス方式トーナメント マッチングシステム

**開発者向けドキュメント**

このドキュメントはシステムの内部構造と開発ワークフローを説明します。
使用方法については [README.md](../README.md) を参照してください。

---

## アーキテクチャ概要

Google Apps Script (GAS) ベースのスイス方式トーナメントマッチングシステム。スプレッドシートを UI 兼データベースとして使用。

### コアコンポーネント

#### ドメイン層

- **player-domain.js**: プレイヤードメイン
  - プレイヤー操作: 登録・ドロップアウト
  - 順位表示: 勝点順の順位表、勝率（OMW%）計算
  - 統計管理: 勝点・勝敗数・試合数の管理
- **match-domain.js**: 対戦ドメイン

  - スイス方式マッチング: `matchPlayersSwiss()` - 同勝点マッチング、再戦回避、バイ処理
  - 対戦結果記録: 勝敗・引き分け・バイの記録、統計更新
  - 対戦結果修正: `correctMatchResult()` - 誤記録の修正

- **round-manager.js**: ラウンド管理
  - ラウンド制御: `startNewRound()` - 新ラウンド開始、マッチング実行
  - ラウンド状態管理: 現在ラウンド番号の取得・設定
  - トーナメント終了処理: `finishTournament()` - トーナメント終了

#### 共通層

- **shared.js**: 共有ユーティリティ
  - シート操作: `getSheetStructure()` - ヘッダー検証とデータ取得
  - プレイヤー名取得: `getPlayerName()` - ID から名前を解決
  - UI 共通処理: `promptPlayerId()` - プレイヤー ID 入力プロンプト

#### アプリケーション層

- **app.js**: アプリケーション層
  - システム初期化: `onOpen()` - カスタムメニュー、`setupSheets()` - シート作成
  - システム設定: `getMaxTables()`, `setMaxTables()` - 最大卓数管理
  - 排他制御: `acquireLock()`, `releaseLock()` - 同時操作防止

#### その他

- **constants.js**: システム定数（シート名、ステータス、卓設定、スイス方式設定）
- **test-utils.js**: テストデータ生成

### 3 つのシート構造

1. **プレイヤー**: プレイヤーマスタ（ID、名前、勝点、勝数、敗数、試合数、OMW%、参加状況、最終対戦日時）
2. **対戦履歴**: 完了した対戦の記録（対戦 ID、ラウンド、日時、卓番号、両者 ID・名前、勝者名、結果）
3. **現在のラウンド**: 進行中のラウンドの対戦（ラウンド、卓番号、両者 ID・名前、結果）

## 重要な設計パターン

### 状態遷移フロー

プレイヤーは 2 つの状態を持つ:

- `参加中` (ACTIVE): トーナメントに参加中
- `終了` (DROPPED): ドロップアウト済み（以降のラウンドに参加しない）

**プレイヤーステータスは定数で管理**:

```javascript
const PLAYER_STATUS = {
  ACTIVE: "参加中",
  DROPPED: "終了",
};
```

### ロック機構の必須パターン

複数ユーザーの同時操作を防ぐため、**すべての書き込み処理でロック取得が必須**:

```javascript
let lock = null;
try {
  lock = acquireLock("操作名");
  // 処理内容
} catch (e) {
  Logger.log("エラー: " + e.message);
} finally {
  releaseLock(lock);
}
```

**デッドロック回避**: `player-domain.js` では複数ロックを固定順序で取得（状態変更 → 対戦結果）。

### マッチングアルゴリズム

1. **待機中プレイヤーのソート優先順位**:

- 勝点降順で並べ、同点は勝数降順 → 試合数昇順で整列
- 勝点が同じプレイヤーはラウンドごとにシャッフルしてペアリング

2. **再戦回避**:

   - 対戦履歴を Map/Set でキャッシュし、O(1) で過去対戦相手をチェック
   - 未対戦相手のみマッチング
   - 全員が過去対戦者の場合は**マッチングを成立させず待機継続**

3. **バイ（不戦勝）処理**:

   - 奇数人数の場合、勝点が最も低いプレイヤーに自動的にバイを付与
   - バイを受けたプレイヤーは勝利扱いで 3 勝点を獲得
   - バイは対戦履歴に記録されるが、OMW%の計算からは除外

4. **卓番号の割り当て**:

- 各ラウンド開始時に卓番号を 1 から順に採番
- Bye の卓番号も同じカウンタで記録し、履歴に残す

5. **パフォーマンス最適化**:
   - 全データを一括取得してキャッシュ（シートアクセス最小化）
   - プレイヤー名を Map でキャッシュ（O(1) 取得）
   - 対戦履歴を Map<PlayerId, Set<OpponentId>> で構築（O(1) 検索）
   - インラインソートで中間関数呼び出しを削減

### マッチングの実行タイミング

- `startNewRound()` が `matchPlayersSwiss()` を呼び出し、ラウンドごとにマッチングを生成
- プレイヤー登録や対戦結果の記録ではマッチングを再計算せず、次のラウンド開始時に反映

### データ構造の検証パターン

すべてのシート操作は `getSheetStructure()` 経由でヘッダー検証:

```javascript
const { indices, data } = getSheetStructure(sheet, SHEET_PLAYERS);
const playerId = row[indices["プレイヤーID"]];
```

- `REQUIRED_HEADERS` 定数で必須列を定義
- 列インデックスを動的に取得（列順変更に対応）

## 開発ワークフロー

### Clasp セットアップ

```bash
npm install -g @google/clasp
clasp login
clasp clone "スクリプトID"  # GAS プロジェクトから取得
clasp push --watch  # 自動アップロード推奨
```

### Git + Clasp 並行運用

- `.clasp.json` は `.gitignore` に追加（各開発者が個別に clasp clone）
- コミットメッセージは**日本語**で記述
- `clasp push` と `git push` を併用して GAS とリポジトリを同期

### テスト実行

スプレッドシート上でカスタムメニューから実行:

1. 「シートの初期設定」でシート構造を作成
2. スクリプトエディタから `registerTestPlayers()` を実行してテストデータ生成

## プロジェクト固有の規約

### 命名規則

- プレイヤー ID: `P` + 3 桁数字（例: `P001`）- `ID_DIGITS` 定数で制御
- 対戦 ID: `T` + 4 桁数字（例: `T0001`）
- 卓番号: 1 ～ 200（`getMaxTables()` で動的取得、デフォルト: 50）

### UI 入力規則

ユーザーインターフェースでは**数字部分のみ**入力を要求:

```javascript
// ユーザーが「1」と入力 → システムで「P001」に整形
const playerId =
  PLAYER_ID_PREFIX +
  Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));
```

### ロケール設定

- タイムゾーン: `Asia/Tokyo`（`appsscript.json`）
- 日時フォーマット: `yyyy/MM/dd HH:mm:ss`

### ログとエラーハンドリング

- すべての catch ブロックで `Logger.log()` にエラー詳細を記録
- ユーザーには `ui.alert()` で簡潔なメッセージを表示
- データ不整合検出時は警告ログを出力しつつ処理継続

## 重要な注意事項

1. **ロック取得順序の厳守**: 複数ロック取得時はデッドロック防止のため順序固定
2. **カスタムメニュー関数はロック管理不要**: 内部の実処理関数でロック取得
3. **数値の型変換**: `parseInt()` 使用時は必ず基数 `10` を指定
4. **パフォーマンス最適化の維持**: `matchPlayersSwiss()` は Map/Set キャッシュで最適化済み。過去対戦相手のチェックは関数内で完結しており、外部ヘルパー関数は使用しない。

## 外部依存関係

- Google Apps Script V8 ランタイム
- SpreadsheetApp サービス
- LockService（排他制御）
- Utilities（日時フォーマット、文字列整形）

## 貢献ガイド

### Pull Request の提出

1. このリポジトリをフォーク
2. 機能ブランチを作成 (`git checkout -b feature/amazing-feature`)
3. 変更をコミット (`git commit -m 'feat: 素晴らしい機能を追加'`)
4. ブランチにプッシュ (`git push origin feature/amazing-feature`)
5. Pull Request を作成

### コミットメッセージ規約

Conventional Commits (https://www.conventionalcommits.org/ja/) に準拠:

- `feat:` 新機能
- `fix:` バグ修正
- `docs:` ドキュメントのみの変更
- `refactor:` リファクタリング
- `perf:` パフォーマンス改善
- `test:` テスト追加・修正

### コードレビューのポイント

- ✅ ロック機構が適切に使用されているか
- ✅ エラーハンドリングが適切か
- ✅ ログ出力が適切か
- ✅ パフォーマンスへの影響を考慮しているか
- ✅ ドキュメントが更新されているか

## 参考リソース

- 詳細な使用方法: `README.md`
- テスト関数例: `test-utils.js`
- GAS API ドキュメント: https://developers.google.com/apps-script
