# lht-cmn

`lht-cmn` は `local-html-tools` 全体で共有する UI コンポーネント基盤です。

- Version: `v20260308`
- License: Apache License 2.0 (`lht-cmn/LICENSE`)
- Copyright: Toshiki Iga

## ライセンスと帰属

- `lht-cmn` 自体は Apache License 2.0 で配布します。
- デザイン方針は Material Design 3 の設計原則を参照します。
- 実装技術として、`lht-cmn` は必要に応じて Material Web（`@material/web`）を優先利用します。
- Material Web のライセンスは Apache License 2.0 です。
- 帰属情報の詳細は `lht-cmn/NOTICE` に記載します。

## 目的

`local-html-tools` では、入力・選択・ヘルプ・コピー・メニューなどの UI を複数ページで繰り返し実装してきました。  
`lht-cmn` はこの重複を減らし、UI を `lht-*` Web Components として共通化するためのレイヤーです。

## 基本方針

- デザイン基準は Material Design 3
- 画面側の公開 UI レイヤーは常に `lht-*` とし、`md-*` を直接使わせない
- `lht-*` は self-contained を原則とし、アプリ側へ `md-*` の登録責務を漏らさない
- 実現手段として Material Web を優先利用してよい
- ただし Material 利用の有無に関わらず、`lht-*` は最低保証で壊れないことを優先する
- 内部実装として許可する型は次の 2 つに限定する:
  - `md-*` 優先 + fallback
  - 完全自前実装
- 「Material 依存で fallback なし」は原則避け、採る場合は README に明示する
- fallback は Material の完全再現ではなく、公開 API の最低保証に留める

## メリット

- 画面ごとの重複実装を削減できる
- 見た目と挙動（必須表示、ヘルプ表示、フォーカス時の挙動）を統一できる
- 変更点を `lht-cmn` に集約でき、保守・レビューがしやすくなる
- 単一HTML生成前提でも、開発時の部品再利用性を維持できる
- 変更点が局所化され、生成AIが誤って別画面を壊す確率が下がる
- UI規約が `lht-*` に集約され、提案が毎回同じ型で出せる
- レビュー時に「画面の見た目差分」より「共通部品の差分」を見ればよくなり、判断が速くなる

## 運用方針（重要）

- 画面側（`docs/*-src.html`）は `lht-*` を利用し、`md-*` 直接実装の追加は原則避ける
- `lht-cmn/js/components.js` を共通コンポーネントの正本とする
- `lht-cmn/css/components.css` を実運用スタイルの正本とする
- `lht-cmn/` 配下（特に `js/components.js` / `css/components.css`）の変更は、必ずユーザーの明示許可を得てから実施する
- `md3/` は段階的にリファレンス用途へ縮退し、実運用スタイルは `lht-cmn` に集約する

## 構成

- `lht-cmn/js/components.js`
  - 共通 Web Components 定義
- `lht-cmn/css/components.css`
  - 上記コンポーネントの共通スタイル
- `lht-cmn/catalog/index.html`
  - 実表示と HTML 利用例を並べて確認するコンポーネントカタログ

### コンポーネント一覧

| コンポーネント | できること | 内部構造（概要） |
|---|---|---|
| `lht-help-tooltip` | `(i)` ヘルプアイコンとツールチップを1タグで配置できる | `md-icon-button` が使える環境ではそれを利用し、未定義時はネイティブ `button` fallback を生成。タグ内HTMLをツールチップ本文へ差し込む |
| `lht-text-field-help` | ラベル付き入力（単一行/複数行）とフォーカス時ヘルプ表示を共通化できる | `md-outlined-text-field` が使える環境ではそれを利用し、未定義時はネイティブ `input` / `textarea` fallback を生成。属性（`field-id`/`label`/`type`/`rows` など）を透過する |
| `lht-select-help` | ラベル付きドロップダウンとヘルプ表示を共通化できる | 内部で `md-outlined-select` を生成し、`<script type="application/json" slot="options">` から `md-select-option` を構築 |
| `lht-switch-help` | スイッチ + ラベル + ヘルプを1セットで配置できる | `md-switch` が使える環境ではそれを利用し、未定義時は `input.md-switch-input + span.md-switch` fallback を生成。`on-change` 指定時はグローバル関数を呼び出す |
| `lht-command-block` | コマンド表示枠とコピー操作（単一/二重ボタン）を共通化できる | `md-icon-button` が使える環境ではそれを利用し、未定義時はネイティブ `button` fallback を生成。クリック時に Clipboard API（不可時は `textarea` フォールバック）でコピー |
| `lht-page-menu` | 右上メニュー（トップへ戻る等）を共通化できる | 内部でメニューボタン + パネル + リンクを生成。外側クリックで自動クローズ |
| `lht-index-card-link` | `docs/index` のカードリンクを統一フォーマットで記述できる | 内部でカードDOM（`a` + タイトル + 説明 + 矢印/バッジ）を生成し、`variant`/`target`/外部リンク判定を吸収 |
| `lht-file-select` | ファイル選択UI（Filledボタン + ファイル名表示）を共通化できる | 内部で hidden `input[type=file]` とトリガUIを生成。`md-filled-button` が使える環境ではそれを利用し、未定義時はフォールバックボタンで動作維持。公開イベントで open ownership を制御できる |
| `lht-loading-overlay` | 処理中オーバーレイ（スピナー + メッセージ）を共通化できる | `active` で表示制御し、`aria-live`/`aria-hidden` を同期。必要に応じて `aria-busy` 更新と操作無効化も連動 |
| `lht-toast` | 一時通知トースト（コピー完了など）を共通化できる | `active` と `show()/hide()` で表示制御し、`role="status"` / `aria-live="polite"` を標準化。`window.showToast` が未定義なら自動補完 |
| `lht-error-alert` | エラー/警告/情報表示を共通化できる | `variant="error|warning|info"` と `active` / `show()/hide()` で表示制御し、variant ごとに `role` / `aria-live` を標準化（`clear()` は補助） |
| `lht-input-mode-toggle` | 入力モード切替（file/source ラジオ）を共通化できる | 既定ID（`inputModeFile`/`inputModeSource`）を維持しつつ、`source/file` ブロックの `md-hidden` 切替を自動化できる |
| `lht-preview-output` | プレビュー表示 + コピー導線を共通化できる | `preview` 枠とコピーボタンを1タグで提供し、`copy-target-id` 指定で既存出力要素からのコピーにも対応 |

## 利用方法

HTML から次を読み込みます。

- `../../lht-cmn/css/components.css`
- `../../lht-cmn/js/components.js`

開発時は上記ファイルを参照し、最終的な配布物はビルド時に単一HTMLへインライン化します。

ページ固有の見た目調整は各画面側の CSS で実施し、共通的な DOM 生成と振る舞いは `lht-cmn` 側で管理します。

## Integration Contract

- アプリ側の公開 UI レイヤーは `lht-*` とし、`md-*` の登録有無を前提に分岐しない
- `lht-*` が内部で `md-*` を使う場合も、その依存解決責務は `lht-cmn` 側に置く
- 各コンポーネントは次のどちらかで実装する:
  - `md-*` 優先 + fallback
  - 完全自前実装
- 例外的に Material 依存を残す場合は、その事実と理由を README に明示する
- fallback の責務は「公開 API の最低保証」であり、Material の完全再現ではない
- ID の責務:
  - `field-id` / `switch-id` / `input-id` / `button-id` / `file-name-id` などの公開 ID はアプリ側が指定する
  - `lht-cmn` はその ID を、公開 API の対象となる内部要素へ引き継ぐ
- 初期化ライフサイクル:
  - 各 `lht-*` は `connectedCallback()` 完了時に `data-initialized="true"` を付与する
  - それまでは `lht-cmn/css/components.css` が pre-upgrade flash を抑止する
- 公開 API:
  - 属性、公開イベント、公開メソッドとして README に書いたものだけを契約対象とする
  - 内部 DOM の細部は fallback 契約で明記したものを除き非公開とみなす
- CSS 拡張点:
  - アプリ側 CSS は `lht-*` タグ自身、README に明記した公開 class、テーマ変数の上書きに寄せる
  - 内部生成 DOM のタグ構造に直接依存する拡張は、fallback 契約で明記された箇所に限定する

## 適用ルール

- `lht-text-field-help` を使う場合は、`label` と `help-text` の設定を「できない理由がない限り」行う
- `lht-select-help` を使う場合も、`label` と `help-text` の設定を「できない理由がない限り」行う
- `lht` 前提の形へ揃える:
  - 外側の旧ラベル（`label + Required + (i) + :`）は整理する（全画面で対応完了した時点で、この項目はREADMEから削除する）
  - 入力系は `lht-text-field-help` / `lht-select-help` / `lht-switch-help` 側に `label` と `help-text` を集約する
  - 必須指定は可能な限りコンポーネント側（`required`）へ寄せる
- 例外にする場合は、対象画面側に理由を残す（表示密度・既存互換・重複説明の回避など）

### コンポーネント設計規約（表示制御とAPI）

- 表示/非表示を持つ `lht-*` は、表示状態の正本属性を `active` とする
- 表示制御メソッドは `show()` / `hide()` を標準とする
- `clear()` は「内容を消して非表示にする」用途でのみ任意追加する（`show/hide` の代替にはしない）
- `active` の更新時は `aria-hidden` を同期する
- 初期化完了後は `data-initialized="true"` を付与する
- `lht-cmn/css/components.css` は `data-initialized="true"` が付くまで未初期化コンテンツを `visibility:hidden` で隠し、pre-upgrade flash を防ぐ

### Fallback / Parity Table

| コンポーネント | Material 未読込時 | fallback | 現在の最小保証 |
|---|---|---|---|
| `lht-select-help` | 動作する | ネイティブ `select` | `value` / `required` / `disabled` / `change` / `setOptions()` / `getValue()` / `setValue()` |
| `lht-text-field-help` | 動作する | ネイティブ `input` / `textarea` | `value` / `required` / `disabled` / `placeholder` / `min` / `max` / `step` / `rows` |
| `lht-switch-help` | 動作する | `input.md-switch-input + span.md-switch` | `checked` / `change` / `switch-id` / ラベル表示 |
| `lht-file-select` | 動作する | ネイティブ `button` + hidden `input[type=file]` | `input-id` / `button-id` / `file-name-id` / `before-open` / `change` / `multiple` / `disabled` |
| `lht-command-block` | 動作する | ネイティブ `button` | `command-id` / 単一・二重 copy button / click-to-copy |

## ドロップダウン置換手順（`lht-select-help`）

1. 基本は `lht-select-help` を使い、`field-id` / `label` / `help-text` を設定する
2. 選択肢は `lht-select-help` に対して宣言する
   - `<script type="application/json" slot="options">[...]</script>` を使用する
3. `lht-select-help` で `<option>` 子要素は使用しない（後方互換運用は終了）
4. 既存JS互換のため、DOM参照ID（`document.getElementById(...)`）は変更しない

## カード共通化（`lht-index-card-link`）

トップ `index` のリンクカードは、基本的に `lht-index-card-link` で共通化します。

- 目的:
  - カードDOM（`a + title + desc + arrow`）の型を固定する
  - 見た目と挙動をコンポーネント側へ集約する

### 主な属性

- `href`（必須）: 遷移先
- `title`（必須）: タイトル
- `desc`（必須）: 説明文
- `icon`（任意）: タイトル先頭に出すアイコン文字（例: `🧰`）
- `variant`: `default | simple | external`
- `arrow`: `auto | none`
- `target` / `rel`: 必要時に指定（`external` / 外部URL / `_blank` は自動補完あり）
- `badge`: バッジ文字列
- `desc-lines`: 説明文の行数クランプ（数値）

### 使用例

```html
<lht-index-card-link
  href="git/git-branch-diff.html"
  icon="🧰"
  title="Git ブランチ比較"
  desc="2つのブランチ差分を表示するコマンドを生成します。"
  variant="default"
  desc-lines="3">
</lht-index-card-link>
```

## LHT リファレンス

### `lht-help-tooltip`

- 用途: `(i)` ヘルプ表示
- 主な属性: `label`, `wide`, `placement`
- fallback:
  - `md-icon-button` 未読込時はネイティブ `button.md-help-icon-button--fallback` を内部生成する
  - hover / focus-within による tooltip 表示契約は Material / fallback の両方で共通
  - anchor 用の最小 CSS (`position`, `overflow`, tooltip visibility) は `lht-help-tooltip` 側に同梱し、アプリ側の追加 tooltip CSS を前提にしない
- placement:
  - `placement="auto|left|right|top|bottom"` を指定できる
  - 既定値は `auto`
  - `auto` は active 時に viewport overflow が最小になる向きを選び、必要に応じて位置を clamp する

### `lht-text-field-help`

- 用途: テキスト/数値/複数行入力 + フォーカス時ヘルプ
- 主な属性: `field-id`, `label`, `help-text`, `hide-delay-ms`, `type`, `placeholder`, `value`, `rows`, `min`, `max`, `step`, `required`, `disabled`, `field-class`
- fallback:
  - `md-outlined-text-field` 未読込時はネイティブ `input` / `textarea` を内部生成する
  - `rows` 指定時は `textarea` fallback を優先する
  - fallback 時の `help-text` は field 下部の supporting text として表示し、`focus` で表示・`blur` 後 `hide-delay-ms` で非表示にする
  - fallback 時も `title` 属性は補助的に維持する

### `lht-select-help`

- 用途: セレクト入力 + フォーカス時ヘルプ
- 主な属性: `field-id`, `label`, `help-text`, `hide-delay-ms`, `value`, `required`, `disabled`, `field-class`
- fallback:
  - `md-outlined-select` 未読込時はネイティブ `select` を内部生成する
  - fallback 時の `help-text` は field 下部の supporting text として表示し、`focus` で表示・`blur` 後 `hide-delay-ms` で非表示にする
  - fallback 時も `title` 属性は補助的に維持する
- 選択肢定義: `<script type="application/json" slot="options">[...]</script>`
- 補助メソッド:
  - `setOptions([{ value, label, selected?, disabled? }], { preserveValue? })`
  - `getValue()`
  - `setValue(value)`
- lifecycle メモ:
  - declarative options の有無は初期化開始時点で判定する
  - `script[slot="options"]` を利用した場合、初期化時に JSON を読んで内部 option へ反映したあと script ノードは消費される
  - declarative options 利用時は child `option` 監視による再 hydration を行わない
- dynamic options:
  - 動的更新は `setOptions([...])` を正規ルートとする
  - 既定では `preserveValue: true` と同等に動作し、同じ `value` が次の options に残っていれば選択を維持する
  - 既存選択値が次の options に存在しない場合は空に戻す
  - 動的更新と declarative JSON / 手動 child mutation の混在は避ける

### `lht-switch-help`

- 用途: スイッチ + ラベル + ヘルプ
- 主な属性: `switch-id`, `label`, `help-label`, `help-wide`, `checked`, `on-change`
- fallback:
  - `md-switch` 未読込時は `input.md-switch-input + span.md-switch` を内部生成する
  - `switch-id` は fallback 時も checkbox input に付与される
  - `checked` 状態と `change` イベントは Material / fallback の両方で利用できる

### `lht-command-block`

- 用途: コマンド表示 + コピーUI
- 主な属性: `command-id`, `copy-buttons`（`single` / `dual`）
- fallback:
  - `md-icon-button` 未読込時はネイティブ `button.md-copy-button--fallback` を内部生成する
  - コピー動作は Material / fallback の両方で共通

### `lht-page-menu`

- 用途: 右上メニュー（戻るリンク等）
- 主な属性: `home-href`, `home-label`

### `lht-page-hero`

- 用途: ページ先頭の見出しブロック（タイトル + 補助説明 + ヘルプ + メニュー）
- 主な属性: `title`, `subtitle`, `icon`, `help-label`, `help-wide`, `menu-home-href`, `menu-home-label`, `no-menu`
- 本文スロット: ヘルプポップアップに表示する説明HTML

### `lht-index-card-link`

- 用途: `docs/index` 用カードリンク
- 主な属性: `href`, `title`, `desc`, `icon`, `variant`, `arrow`, `target`, `rel`, `badge`, `desc-lines`

### `lht-file-select`

- 用途: ファイル選択UI（ボタン + hidden file input + ファイル名表示）
- 主な属性: `input-id`, `button-id`, `file-name-id`, `accept`, `button-label`, `placeholder`, `file-label`, `multiple`, `disabled`, `show-file-name`, `auto-open`
- 公開イベント:
  - `lht-file-select:before-open`
    - button click 時に発火する cancellable event
    - `auto-open="false"` のときは発火のみ行い、内部 `input.click()` は実行しない
    - `auto-open` が既定 `true` の場合でも `preventDefault()` で内部 open を抑止できる
  - `lht-file-select:change`
    - hidden file input の `change` 後に発火する
    - `detail.names` と `detail.files` で選択結果を参照できる

### `lht-loading-overlay`

- 用途: ファイル読み込みなどの非同期処理中オーバーレイ（indeterminate loading）
- 主な属性: `active`, `text`, `busy-target-id`, `disable-target-ids`
- 補助メソッド: `setActive(boolean)`, `isActive()`, `waitForNextPaint()`
- ARIAルール:
  - 常時 `role="status"` と `aria-live="polite"` を持つ
  - `active` に応じて `aria-hidden` を `false/true` へ同期する
  - `busy-target-id` 指定時は対象へ `aria-busy` を `true/false` で同期する
- 推奨フロー:
  1. `overlay.setActive(true)` で開始
  2. `await overlay.waitForNextPaint()` で先に描画を確定
  3. 重い処理を実行
  4. `finally` で `overlay.setActive(false)` を必ず実行

### `lht-toast`

- 用途: コピー完了などの短時間通知（toast/snackbar）
- 主な属性: `active`, `text`, `duration-ms`
- 補助メソッド: `show(message?, durationMs?)`, `hide()`
- ARIAルール:
  - 常時 `role="status"` と `aria-live="polite"` を持つ
  - 常時 `aria-atomic="true"` を持つ
- 運用メモ:
  - ページ側に `<lht-toast id="toast"></lht-toast>` を1つ配置して使う
  - 既存コードが `window.showToast(...)` を呼ぶ場合、未定義時は `lht-toast` 側が自動補完する

### `lht-error-alert`

- 用途: 画面内エラー/警告/情報表示の共通化（`errorText` パターンの置換）
- 主な属性: `text`, `active`, `variant`
- 補助メソッド: `show(message?)`, `hide()`, `clear()`, `isVisible()`
- ARIAルール:
  - `variant="error"` は `role="alert"` と `aria-live="assertive"`
  - `variant="warning|info"` は `role="status"` と `aria-live="polite"`
  - 常時 `aria-atomic="true"` を持つ
  - 表示状態に応じて `aria-hidden` を同期する

### `lht-input-mode-toggle`

- 用途: `file/source` 入力切替ラジオUIの共通化（music系の重複置換）
- 主な属性: `name`, `group-label`, `file-id`, `source-id`, `file-label`, `source-label`, `default-mode`, `source-target-id`, `file-target-id`, `on-change`, `disabled`
- 補助メソッド: `getMode()`, `setMode(mode)`, `applyModeUi()`
- 互換メモ:
  - 既定の `file-id` / `source-id` は `inputModeFile` / `inputModeSource`
  - 既存JSが `document.getElementById("inputModeFile")` 等を参照していても置換しやすい

### `lht-preview-output`

- 用途: プレビュー表示とコピー導線の共通化（`preview + copyBtn` パターンの置換）
- 主な属性: `preview-id`, `copy-button-id`, `copy-target-id`, `placeholder`, `copy-label`, `copy-aria-label`, `preview-tag`, `no-copy`
- 補助メソッド: `getText()`, `setText(text)`, `copy(targetId?)`, `clear()`
- 運用メモ:
  - 既定の `preview-id` / `copy-button-id` は `previewText` / `copyBtn`
  - `copy-target-id` を指定すると、プレビュー枠とは別要素のテキストをコピーできる

## Appendix

### Appendix A: Material Web 置換の実施手順（実装メモ）

`*-src.html` を `lht-*` 前提へ寄せるときの、実務上の手順メモです。

1. 置換対象を `*-src.html` 上で特定する
2. 既存の生HTML部品を `lht-*`（内部的には Material Web または自前実装）へ置換する
3. 状態取得/保存ロジックを `selected` / `value` ベースへ揃える
4. 見た目差分（角丸、高さ、フォーカスリング、余白）を CSS トークンで吸収する
5. 単一HTMLビルドを実行して動作確認する

### Appendix B: 置換対応表（内部実装の目安）

- テキスト入力: `md-outlined-text-field`
- テキストエリア: `md-outlined-text-field type="textarea"`
- セレクト: `md-outlined-select` + `md-select-option`
- トグル: `md-switch`
- アイコンボタン: `md-icon-button`
- ヘルプ `(i)`: `lht-help-tooltip`
- フィールド活性時ヘルプ表示: `lht-text-field-help`
- スイッチ + ヘルプ: `lht-switch-help`
- コマンド表示 + コピー: `lht-command-block`
- 右上メニュー: `lht-page-menu`

### Appendix C: テーマ色運用メモ

- フォーカス、選択、強調は `primary` 系（`--md-sys-color-primary`）を基準にする
- `secondary` は `primary` と競合しない範囲で使う。迷ったら `primary` に寄せる
- フォーカスリング色はコンポーネント間で統一する
- Material Web の色変更は、まず `:root` の `--md-sys-*` を調整し、個別上書きは最小限にする

### Appendix D: tooltip 実装制約メモ

- `@material/web@2.4.1` では `md-tooltip` が同梱されないため、`lht-help-tooltip` は `md-tooltip-group` + `md-tooltip-content` ベースで運用する

### Appendix E: ドロップダウンでよくあるミスと回避方法

`lht-select-help` は `md-outlined-select` を内部利用するため、単一HTML化や依存読込順の影響を受けやすいです。  
以下のミスが、ドロップダウン崩れ（選択肢がただのテキストになる等）を起こしやすいです。

1. `md-outlined-select` が未定義のまま初期化される
- 症状:
  - 選択UIが表示されず、選択肢テキストだけが並ぶ
- 回避:
  - 標準配置の Material Web バンドル（`lht-cmn/vendor/material-web-outlined-text-field.bundle.js`）を `lht-cmn/js/components.js` より前に配置する
  - `lht-cmn` 側のフォールバック（ネイティブ `select`）が効く実装を維持する

2. `lht-select-help` の選択肢定義が不正
- 症状:
  - 選択肢が空になる / 既定値が反映されない
- 回避:
  - `<script type="application/json" slot="options">[...]</script>` の JSON を必ず配列で定義する
  - `value` と `label` を明示する
  - 既定値は `selected: true` と `value` の整合を取る

3. `field-id` を変えて既存JS参照が壊れる
- 症状:
  - `document.getElementById(...)` が `null` になり、初期化やイベント登録で失敗する
- 回避:
  - 置換時も DOM 参照ID（`field-id`）は既存IDを維持する

4. 単一HTML化でインラインスクリプトが壊れる
- 症状:
  - `Unexpected end of input`
  - バンドル内文字列が壊れ、`popover` などの警告が連鎖する
- 回避:
  - ビルド時に `</script>` を `<\\/script>` へエスケープする
  - 文字列置換でJSを差し込む場合は `replace` の関数置換を使い、`$` 展開事故を避ける

5. CSSの責務が混在して見た目が崩れる
- 症状:
  - ドロップダウンの幅・余白・フォーカス装飾がページごとに不揃い
- 回避:
  - 基本スタイルは `lht-cmn/css/components.css` に集約する
  - 画面側CSSはレイアウト差分（余白・配置）に限定する
