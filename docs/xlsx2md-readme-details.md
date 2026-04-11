# xlsx2md README 詳細

> 補足
> このアプリは `miku-xlsx2md` に名称変更され、リポジトリは <https://github.com/igapyon/miku-xlsx2md> に移動した。
> 原則として、参照先は移転先リポジトリとする。

文書の役割は段階的に分かれています。

- `README.md`
  - 入口説明と全体像
- `docs/xlsx2md-spec.md`
  - 人間が意図と設計方針を理解するための上位仕様
- `docs/xlsx2md-impl-spec.md`
  - 現行実装の詳細仕様と再実装・保守のための文書
- `docs/xlsx-formula-subset.md`
  - 数式対応範囲の補助文書

git に入れない実データや一時検証用の `.xlsx` は `local-data/` に配置します。このディレクトリは `.gitignore` 対象です。

`local-data/` で使う一部サンプルの取得元メモ:

- Microsoft Create planner / tracker templates: <https://excel.cloud.microsoft/create/ja/planner-tracker-templates/>

## 現在の状態

- 実装済み
- `.xlsx` をブラウザ内で読み込み、全シートを自動で Markdown 化できる
- 出力モードとして `display / raw / both` を切り替えられる
- formatting mode として `plain / github` を切り替えられる
- table detection mode として `balanced / border` を切り替えられる
- 画像、表、地の文、リストを扱う
- Excel グラフは画像ではなく、タイトル・系列・参照範囲・副軸などの意味情報として Markdown に出せる
- DrawingML 図形は、要約ではなくタグ名・属性名を保った raw 寄りの階層 Markdown として出せる
- 単純図形 (`rect / line / arrow / text box`) は、raw ダンプを残したまま SVG アセットとしても出せる
- 数式は `cached value` 優先で、必要時のみ自前 evaluator が補完する
- `github` formatting mode では、対応する Excel の `bold / italic / strike / underline` とセル内改行を Markdown / HTML へ反映できる
- セルのハイパーリンクは、対応できる範囲で外部リンク / ブック内リンクとして Markdown へ反映できる
- rich text / Markdown rendering は、`markdown escape -> rich text parser -> plain/github formatter -> table escape` へ段階分離済み
- `styledText.parts` は `text / escaped` と `rawText` を持つ内部表現へ進めている
- fixture と `local-data` を使った検証を継続している

- 未対応または将来改善
- 専用の `セクション分割ブロック` は未導入で、レイアウト中心シートは現在も通常の表検出と地の文抽出をベースに扱う
- カレンダー / ボード / ダッシュボード系の自然な分解は改善余地がある
- dynamic array / spill や複雑な Excel 表現は段階的対応の途中である
- Markdown 記号を含む生文字の escape は段階的に整理中である
- ブック内リンクは現時点では対象シート先頭アンカーへのリンクを基本とし、セル位置そのものへの厳密ジャンプは未対応である

## 主な特徴

- 変換対象は常に全シート
- `.xlsx` 選択後に自動で Markdown 生成まで進む
- Markdown を主表示とし、解析系 UI は補助情報として扱う
- 保存ファイル名は安全側へサニタイズする
- レイアウト中心シートは、見た目再現ではなく `表 / リスト / 画像 / 補助セクション` への分解を優先する
- グラフや図形についても、見た目再現より意味情報や raw metadata を Markdown へ落とすことを優先する

## 方針

- 文書は `docs/` 配下に配置し、既存の `docs/text/` には入れない
- 入出力はブラウザ内で完結し、サーバ送信を前提にしない
- 最初の用途は「Excel 設計書を Markdown へ変換する」に置く
- 配布形態は Single-file Web App とする
- UI は `lht-cmn/` を利用する
- 実装の正本ソースは TypeScript とする
- 自動テストを前提とする
- 出力モードは `display / raw / both` を扱う
- formatting mode は `plain / github` を扱う
- table detection mode は `balanced / border` を扱う
- `raw` / `both` で保存する Markdown や ZIP は、ファイル名にモードサフィックスを付けて区別できるようにする
- `Markdown+assets ZIP保存` は、既定で `連結 Markdown 1 本 + assets/` の構成とする
- ZIP entry timestamp は、再現性のため固定値を使う
- 保存ファイル名は安全側へサニタイズし、シート名の空白や一部記号をそのまま使わないことがある
- 表候補外の縦並び短文は、地の文として連結せず箇条書きへ変換することがある
- レイアウト中心シートは、見た目再現ではなく `表 / リスト / 画像 / 補助セクション` を優先して扱う
- 空欄の多いフォーム風の罫線領域は、現時点では情報脱落を避けるため表として残すことがある。一方で、横に広く疎なフォーム領域は表候補から外して地の文やセクションへ寄せることもある
- 数式は `cached value` を優先し、自前 evaluator は未計算保存や欠損時の補助として扱う
- 数式診断では、`resolved / fallback / unsupported` に加えて、`cached / ast / legacy / formula` の解決経路を表示する
- 数式セルの `cached value` 判定は `<v>` 要素の有無を基準とし、`<v></v>` のような空文字キャッシュも `cached` として扱う
- rich text / `<br>` / Markdown 記号の扱いは、`plain` と `github` の責務を分けて段階的に整理する
- 装飾系は、GitHub 互換の Markdown / HTML へ自然に落とせるものを優先し、自然に落とせない Excel 固有表現は当面無理に再現しない
- ハイパーリンクは GitHub 上で自然に見える Markdown リンクを優先し、リンクセル由来の下線は重ねて出さない
- formatter は `plain` と `github` を別モジュールとして分け、renderer は mode dispatch に寄せる

## 使い方

1. `xlsx2md.html` を開いて `.xlsx` ファイルを選択する
2. 読み込み後、自動で全シートの Markdown が生成される
3. 必要に応じて `display / raw / both` と変換オプションを調整し、Markdown または ZIP を保存する

補足:

- `plain`: 装飾を落として素朴なテキストへ寄せる
- `github`: `bold / italic / strike / underline` とセル内改行を GitHub 互換の Markdown / HTML へ寄せる
- ハイパーリンクは `[text](url)` または workbook 内アンカーへのリンクとして出す
- `balanced`: 既定の表検出モード
- `border`: 罫線のある領域からだけ表を検出し、borderless fallback 検知を抑えるモード

## UI の考え方

- Markdown を主表示とし、解析サマリーはその下に置く
- `解析サマリー`
- `変換設定`
- `表候補スコア`
- `数式診断`
  はアコーディオンで普段は閉じる
- 診断系は通常利用では補助情報として扱う

## Markdown 出力上の補助識別子

- 全シート連結 Markdown では、各チャンク先頭に HTML コメントで一意識別子を付けてよい
- この識別子は分割 Markdown ファイル名そのものではなく、チャンク識別子として扱う
- コメント中の識別子は、連結 Markdown 内で元シート断片を追跡しやすくするための補助メタ情報である
- 例:

```markdown
<!-- workbook_001_SheetName -->
```

## Formatting Mode

- `plain`
  - Excel の文字装飾を落として Markdown 化する
  - セル内改行は空白へ正規化する
- `github`
  - 対応する Excel の文字装飾を GitHub 互換の Markdown / HTML へ寄せる
  - `bold -> **...**`
  - `italic -> *...*`
  - `strike -> ~~...~~`
- `underline -> <ins>...</ins>`
- ただし hyperlink セルはリンク自体で視認できるため、underline は追加で出さない
  - セル内改行は `<br>` として扱う

## Table Detection Mode

- `balanced`
  - 既定の表検出モード
  - border seed と value seed の両方を使って表候補を作る
  - 罫線のない dense な値ブロックも表候補になりうる
- `border`
  - 罫線のある表候補を優先するモード
  - borderless fallback 検知を抑える
  - 非罫線ベースの誤検知が辛い workbook / sheet 向け

最小の実ファイル回帰として、`tests/fixtures/table/table-border-priority-sample01.xlsx` では同一データに対して

- `balanced`: 表として検出
- `border`: narrative として扱う

差分を固定している。

ただし、Markdown 記号を含む生文字の escape は現時点では段階的整備の途中である。
設計整理のメモは [rich-text-markdown-rendering.md](./rich-text-markdown-rendering.md) を参照。

内部構成メモ:

- `markdown-escape.ts`
  - 生文字の Markdown / HTML セーフティを担当
- `rich-text-parser.ts`
  - `text / lineBreak / styledText` token 列へ変換
- `rich-text-plain-formatter.ts`
  - `plain` 向け描画
- `rich-text-github-formatter.ts`
  - `github` 向け描画
- `markdown-table-escape.ts`
  - table cell 専用 escape

## 出力モード

- `display`: Excel の表示値寄りで Markdown を出力する標準モード
- `raw`: Excel 内部値を優先して Markdown を出力するモード
- `both`: 表示値を本文に出しつつ、必要に応じて `[raw=...]` を補助表示するモード

## モード選択の目安

- `display`: 人間が Excel を見ながら生成 AI と内容を共有したいとき
- `raw`: 内部値や未加工値を確認したいとき
- `both`: 表示値と内部値の差分を比較しながら確認したいとき

## 出力例

同じセルでも、出力モードによって次のように表現が変わります。

```markdown
date
display: 2024/1/1
raw: 45292
both: 2024/1/1 [raw=45292]

currency
display: ¥1,024,768
raw: 1024768
both: ¥1,024,768 [raw=1024768]

fraction
display: 3/4
raw: 0.75
both: 3/4 [raw=0.75]

formula-date
display: 2024/3/17
raw: 45368
both: 2024/3/17 [raw=45368]
```

## 保存名の扱い

- 保存ファイル名は Workbook 名、シート順、シート名から組み立てる
- シート名の空白や一部記号は、保存名では `_` に寄せることがある
- ZIP 保存時の Markdown は、シートごとの分割ではなく Workbook 単位の連結 Markdown 1 本を標準とする
- ZIP entry timestamp は `2025-01-01 00:00:00` の固定値を使う
- これは「本当の作成日時」を表すためではなく、同じ入力から同じ ZIP バイナリを作りやすくするため
- 例:
  - シート名 `A B-東京&大阪.01`
  - 保存名 `edge-weird-sheetname-sample01_001_A_B-東京_大阪.01.md`
  - 連結 Markdown 保存名 `edge-weird-sheetname-sample01.md`

## レイアウト中心シートの扱い

- `イベント プランナー`、`月間プランナー`、`老後資金プランナー` のようなレイアウト中心シートでは、巨大な 1 表へ吸い込まず `セクション / 表 / リスト / 画像` のまとまりを優先する
- 現在は最小実装として、縦方向に大きな空白がある場所でセクションを分け、Markdown 上では `---` で区切る
- セクション先頭の短い単行テキストは、見出し候補として `###` へ昇格することがある
- Excel グラフは `## グラフ` セクションで、anchor・タイトル・種別・系列・参照範囲・副軸を Markdown 化する
- DrawingML 図形は `## 図形` セクションで、`xdr:*` / `a:*` のタグ名・属性名を保った階層ダンプとして Markdown 化する
- 単純図形 (`rect / line / arrow / text box`) は、図形セクションの raw ダンプを保持したまま `SVG` アセット参照も併記する
- ただしフォーム風領域、カレンダー、ボード、ダッシュボードの完全再現はまだ対象外で、今後の改善余地として扱う

## 既知の制約

- 未対応寄り
- `spill` / dynamic array は最小対応の段階で、`A1#` のような表現や runtime 入口はあるが、実データ fixture による十分な検証はこれから
- 配列定数や `space intersection` は最小対応済みだが、Excel 互換としてはまだ限定的
- rich text / `<br>` / Markdown 記号 escape は段階的整備の途中で、専用 renderer の責務分離は今後の検討課題
- 専用の `セクション分割ブロック` は未導入で、レイアウト中心シートも現在は通常の表検出と地の文抽出を優先する
- フォーム風領域は現時点では保守的に扱う
  - 表として残す場合がある
  - 横に広く疎なものは narrative / section 側へ寄せる場合がある
- 図形は raw metadata のダンプに対応し、単純図形 (`rect / line / arrow / text box`) は SVG アセット化にも対応している
- SmartArt は現時点では意味解釈対象にせず、fallback 扱いとする
- グラフは当面、タイトル・系列・参照範囲・副軸などのテキストメタデータ出力を標準とし、SVG 化は着手しない
- DrawingML の接続先解釈や複雑図形の完全意味解釈は未対応
- カレンダー / ボード / ダッシュボード系は、通常表としての自然な表現が難しく、現在は `セクション / 表 / リスト / 画像` への分解を優先している
- Excel 数式全体を完全再実装しているわけではなく、通常は `cached value` を主に使い、自前 evaluator は補助として働く

- 実装済み寄り
- `display / raw / both` の切替
- 数式診断での `resolved / fallback / unsupported` と `cached / ast / legacy / formula` 表示
- 空文字の `cached value` を `cached` とみなす扱い

## 図形 SVG 対応の進め方

- DrawingML 図形は、まず `anchor / name / prstGeom / text / extents / rawEntries` を安定して抽出し、Markdown の `## 図形` セクションへ raw metadata を残すことを優先する
- そのうえで、SVG 化は `prstGeom` ごとに段階的に追加する
- 新しい図形を対応させるときは、先に fixture を追加して `rawEntries` と Markdown 出力を固定し、その後で `office-drawing.ts` に renderer を足す
- SVG 化は完全再現を前提にせず、まずは簡略化した形状でよい
  - 例: `flowChartDecision -> diamond`
  - 例: `flowChartInputOutput -> parallelogram`
  - 例: `rightArrow -> polygon`
- 既存図形との同型は、できるだけ同じ renderer を再利用する
  - 例: `rect` と `flowChartProcess`
  - 例: `roundRect` と `flowChartTerminator`
- connector、単純矩形、単純矢印のような図形から先に対応し、callout や複雑図形は後段で扱う
- SmartArt や複雑図形は、当面は raw metadata と fallback を優先し、SVG は無理に広げない

## 現在の構成イメージ

```text
.
├── README.md
├── docs/
│   ├── local-data-review.md
│   ├── xlsx2md-spec.md
│   ├── xlsx2md-impl-spec.md
│   ├── xlsx2md-readme-details.md
│   └── xlsx-formula-subset.md
├── xlsx2md-src.html
├── xlsx2md.html
├── references/
│   └── MS-XLSX-parser-grammar.abnf
├── src/
│   └── xlsx2md/
│       ├── css/
│       │   └── app.css
│       ├── ts/
│       │   ├── core.ts
│       │   ├── main.ts
│       │   └── formula/
│       │       ├── tokenizer.ts
│       │       ├── parser.ts
│       │       └── evaluator.ts
│       └── js/
│           ├── core.js
│           ├── main.js
│           └── formula/
│               ├── tokenizer.js
│               ├── parser.js
│               └── evaluator.js
├── tests/
│   ├── xlsx2md-main.test.js
│   ├── xlsx2md-formula-parser.test.js
│   └── fixtures/
│       ├── README.md
│       ├── display/
│       ├── merge/
│       ├── formula/
│       ├── narrative/
│       ├── image/
│       ├── named-range/
│       └── edge/
└── local-data/
    └── ...
```
