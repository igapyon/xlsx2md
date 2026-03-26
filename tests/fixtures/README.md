# xlsx2md Fixtures

`docs/xlsx2md/tests/fixtures/` は、`xlsx2md` の実ファイルベース回帰テスト用 `.xlsx` 置き場です。

## ルール

- できるだけ `1ファイル = 1目的` にする
- `1シート` で済むものは `1シート` にする
- 値は少なく、意図は明確にする
- 先頭付近に「何を試すサンプルか」が分かる見出しセルを置く
- 可能ならファイル名に `sample01` を付ける
- 通常は普通に Excel 保存して `cached value` を残す
- 未計算保存を試す場合だけ、別ファイルへ分ける
- Excel 保存由来の環境依存メタデータはできるだけ落とす

## Excel 保存メタデータの注意

一部の `.xlsx` は、Excel 保存時に `xl/workbook.xml` や `docProps/*.xml` へ環境依存メタデータを含むことがある。

- これは `xlsx2md` の実装由来ではなく、保存元 Excel が埋めるメタデータである
- 変換処理では通常参照しないが、fixture としてはノイズなので残さない方がよい
- 特に Git 管理する fixture では、ローカル絶対パス、作成者、更新者、作成日時、更新日時、保存アプリ情報が残ることがあるので削除を推奨する
- fixture では、コメント、threaded comments、person 情報、custom properties、external links、connections、embeddings、VBA、printer settings も原則残さない

対処方法:

1. `.xlsx` を ZIP として展開する
2. `xl/workbook.xml` から `x15ac:absPath` を含む `mc:AlternateContent` を削除する
3. `docProps/core.xml` から `dc:creator`、`cp:lastModifiedBy`、`dcterms:created`、`dcterms:modified` を削除する
4. `docProps/app.xml` から `Application`、`AppVersion` を削除する
5. ZIP として再圧縮して `.xlsx` を置き換える

確認方法:

- `unzip -p <file.xlsx> xl/workbook.xml | rg 'x15ac:absPath|absPath url='`
- `unzip -p <file.xlsx> docProps/core.xml | rg 'dc:creator|cp:lastModifiedBy|dcterms:created|dcterms:modified'`
- `unzip -p <file.xlsx> docProps/app.xml | rg '<Application>|<AppVersion>'`
- `unzip -l <file.xlsx> | rg 'xl/comments[0-9]*\.xml|xl/threadedComments/|xl/persons/person\.xml|docProps/custom\.xml|xl/externalLinks/|xl/connections\.xml|xl/embeddings/|xl/vbaProject\.bin|xl/printerSettings/'`

該当があれば、fixture 作成後に一度この確認を行う。

## 既存 fixture

| ファイル | 主目的 | 対応章 | 主に確認する症状 |
| --- | --- | --- |
| `xlsx2md-basic-sample01.xlsx` | 総合サンプル | `xlsx2md-spec.md` 6, 7, 10, 13 | 表と地の文の崩れ、基本 Markdown 差分 |
| `display/display-format-sample01.xlsx` | 表示形式 | `xlsx2md-spec.md` 12 | `display / raw / both` の見え方差分 |
| `formula/formula-basic-sample01.xlsx` | 基本数式 | `xlsx2md-spec.md` 10, 11 | `cached / resolved / fallback` の差分 |
| `formula/formula-crosssheet-sample01.xlsx` | 複数シート参照 | `xlsx2md-spec.md` 10, 19 | sheet 参照解決漏れ、参照先ずれ |
| `formula/formula-shared-sample01.xlsx` | shared formula | `xlsx2md-spec.md` 10, 19 | shared formula 展開漏れ、連番列の崩れ |
| `formula/formula-spill-sample01.xlsx` | dynamic array / spill | `xlsx2md-spec.md` 19 | spill 解決漏れ、`A1#` 参照崩れ |
| `link/hyperlink-basic-sample01.xlsx` | ハイパーリンク | `xlsx2md-spec.md` 5, 19 | 外部リンク / ブック内リンクの保持、リンクセルの下線出力 |
| `rich/rich-text-github-sample01.xlsx` | rich text / 文字装飾 | `xlsx2md-spec.md` 5, 6 | GitHub 互換の `bold / italic / strike / underline` 出力 |
| `rich/rich-markdown-escape-sample01.xlsx` | Markdown 記号 / escape | `xlsx2md-spec.md` 5, 6 | Markdown 記号を含む文字列と rich text の混在 |
| `rich/rich-usecase-sample01.xlsx` | rich text + hyperlink 実用例 | `xlsx2md-spec.md` 5, 6, 19 | 表セル内 rich text、外部リンク、改行、取消線つき補足文の安定出力 |
| `merge/merge-pattern-sample01.xlsx` | 結合セル | `xlsx2md-spec.md` 13 | `[MERGED←] / [MERGED↑]` の崩れ |
| `merge/merge-multiline-sample01.xlsx` | 結合セル内改行 | `xlsx2md-spec.md` 13 | multiline merged text の parse と Markdown 正規化 |
| `table/table-basic-sample01.xlsx` | 隣接表(縦) | `xlsx2md-spec.md` 7, 8 | 縦に密接した独立表の誤結合 |
| `table/table-basic-sample02.xlsx` | 隣接表(横) | `xlsx2md-spec.md` 7, 8 | 横に密接した独立表の誤結合 |
| `table/table-basic-sample03.xlsx` | 隣接表(縦横) | `xlsx2md-spec.md` 7, 8 | 4表密集時の過剰な一体検出 |
| `table/table-basic-sample11.xlsx` | 方眼紙表(単体) | `xlsx2md-spec.md` 7, 8, 13 | merge 多用の方眼紙風表の取りこぼし |
| `table/table-basic-sample12.xlsx` | 方眼紙表(縦) | `xlsx2md-spec.md` 7, 8, 13 | merge 多用の方眼紙風 2 表の誤結合 |
| `table/table-basic-sample13.xlsx` | 方眼紙表(縦横) | `xlsx2md-spec.md` 7, 8, 13 | merge 多用の方眼紙風 4 表の誤結合 |
| `table/table-basic-sample14.xlsx` | 方眼紙表(結合漏れ) | `xlsx2md-spec.md` 7, 8, 13 | merge 多用表で一部だけ結合漏れがある場合の崩れ |
| `table/table-basic-sample15.xlsx` | 方眼紙表(縦結合混在) | `xlsx2md-spec.md` 7, 8, 13 | merge 多用表で縦結合が混じる場合の崩れ |
| `table/table-border-priority-sample01.xlsx` | border モード | `xlsx2md-spec.md` 7, 8 | borderless dense block の誤検知と `border` 差分 |

### ルート直下

- `xlsx2md-basic-sample01.xlsx`
  - 総合サンプル
  - 地の文、表、結合セル、shared formula、表示形式をまとめて確認する
  - 対応章: `xlsx2md-spec.md` 6, 7, 10, 13
  - 主に確認する症状: 表と地の文の崩れ、基本 Markdown 差分

### `display/`

- `display-format-sample01.xlsx`
  - 表示形式専用
  - 数値、通貨、会計、日付、時刻、パーセンテージ、分数、指数、文字列、和暦を確認する
  - 対応章: `xlsx2md-spec.md` 12
  - 主に確認する症状: `display / raw / both` の見え方差分

### `formula/`

- `formula-basic-sample01.xlsx`
  - 基本数式専用
  - `ref`、算術、`IF`、`SUM`、`COUNTIF`、`TEXT`、`DATE`、`VALUE` を確認する
  - 対応章: `xlsx2md-spec.md` 10, 11
  - 主に確認する症状: `cached / resolved / fallback` の差分
- `formula-crosssheet-sample01.xlsx`
  - 複数シート参照専用
  - sheet 参照と日本語シート名参照を確認する
  - 対応章: `xlsx2md-spec.md` 10, 19
  - 主に確認する症状: sheet 参照解決漏れ、参照先ずれ
- `formula-shared-sample01.xlsx`
  - shared formula 専用
  - オートフィル由来の連番列を確認する
  - 対応章: `xlsx2md-spec.md` 10, 19
  - 主に確認する症状: shared formula 展開漏れ、連番列の崩れ
- `formula-spill-sample01.xlsx`
  - dynamic array / spill 専用
  - `_xlfn.SEQUENCE(3)` と `SUM(C4#)` の cached value を含めた最小ケース
  - 対応章: `xlsx2md-spec.md` 19
  - 主に確認する症状: spill 由来セルの取り扱い、`#` 参照の保持

### `rich/`

- `rich-text-github-sample01.xlsx`
  - rich text / 文字装飾専用
  - セル全体装飾と部分装飾を使って、GitHub 互換の `bold / italic / strike / underline` 変換を確認する
  - `A9` は italic ベースに strike を一部上乗せした混在ケース、`B9` は `italic+strike` のセル全体装飾ケースとして使う
  - `A12` は `underline / strike / bold` を同一セル内で混在させた日本語ケース、`B12` は `bold + italic + underline` の複合装飾ケースとして使う
  - 対応章: `xlsx2md-spec.md` 5, 6
  - 主に確認する症状: 文字装飾の脱落、run 境界での空白欠落、表セル内装飾の崩れ
- `rich-markdown-escape-sample01.xlsx`
  - Markdown 記号 / escape 専用
  - `*`, `_`, `~~`, `#`, `-`, `1.`, link 風、image 風、backtick、`<tag>`, `|`, `\` を含む文字列を確認する
  - セル全体装飾、部分装飾、セル内改行、表セルを同時に含める
  - 対応章: `xlsx2md-spec.md` 5, 6
  - 主に確認する症状: Markdown 記号の誤解釈、`github` モードでの `<br>` 変換、表セル内の崩れ
  - 期待観点:
    - narrative では backtick や image 風文字列を literal のまま出す
    - table では `|` だけを優先的に escape し、他の文字列は過剰変換しない
    - `path\\to\\file` や `&lt;tag&gt;` のような断片が mode 差をまたいでも安定する
- `rich-usecase-sample01.xlsx`
  - rich text + hyperlink の実用寄りケース
  - 外部リンク付きの表で、説明列や補足列に部分装飾、セル内改行、取消線つき補足文を含めて確認する
  - `github` では `bold / italic / strike / underline / <br>` と Markdown リンクが共存することを確認する
  - `plain` では装飾を落としつつ、リンクラベルと文面が素直なテキストとして残ることを確認する
  - 対応章: `xlsx2md-spec.md` 5, 6, 19
  - 主に確認する症状: 表セル内 rich text の崩れ、リンクセルの出力、`github/plain` 差分、取消線つき補足文の扱い

### `link/`

- `hyperlink-basic-sample01.xlsx`
  - ハイパーリンク専用
  - 外部 URL とブック内リンクを、narrative と table の両方で確認する
  - 対応章: `xlsx2md-spec.md` 5, 19
  - 主に確認する症状: 外部リンクの保持、ブック内リンクの保持、リンクセルの下線が Markdown 出力へ過剰反映されないこと

### `merge/`

- `merge-pattern-sample01.xlsx`
  - 結合セル専用
  - 横結合、縦結合、2x2 結合と `[MERGED←] / [MERGED↑]` を確認する
  - 対応章: `xlsx2md-spec.md` 13
  - 主に確認する症状: `[MERGED←] / [MERGED↑]` の崩れ
- `merge-multiline-sample01.xlsx`
  - 結合セル内改行専用
  - 2x2 結合セルの先頭セルに multiline text を入れて確認する
  - parse では改行を保持し、Markdown では現状空白へ正規化されることを前提にする
  - 対応章: `xlsx2md-spec.md` 13
  - 主に確認する症状: multiline merged text の脱落、Markdown 正規化時の崩れ

### `table/`

- `table-basic-sample01.xlsx`
  - 独立した表が縦に密接しているケース
  - 見出し行や注記行を挟まずに上下へ並ぶ 2 表を確認する
  - 対応章: `xlsx2md-spec.md` 7, 8
  - 主に確認する症状: 縦に密接した独立表の誤結合
- `table-basic-sample02.xlsx`
  - 独立した表が横に密接しているケース
  - 表の間に補助列の文字セルがあっても別表として扱えるか確認する
  - 対応章: `xlsx2md-spec.md` 7, 8
  - 主に確認する症状: 横に密接した独立表の誤結合、補助列の narrative 混入
- `table-basic-sample03.xlsx`
  - 独立した表が縦横に密接して 4 表並ぶケース
  - 2x2 配置の全体を 1 つの大きな表として誤検出しないか確認する
  - 対応章: `xlsx2md-spec.md` 7, 8
  - 主に確認する症状: 4表密集時の過剰な一体検出
- `table-basic-sample11.xlsx`
  - 方眼紙風に merge を多用した単表ケース
  - 見た目上は広い方眼紙でも 1 つの表として抽出できるか確認する
  - 対応章: `xlsx2md-spec.md` 7, 8, 13
  - 主に確認する症状: merge 多用の方眼紙風表の取りこぼし
- `table-basic-sample12.xlsx`
  - 方眼紙風に merge を多用した表が縦に 2 つ並ぶケース
  - 説明セルを挟んでも上下の表を別表として扱えるか確認する
  - 対応章: `xlsx2md-spec.md` 7, 8, 13
  - 主に確認する症状: merge 多用の方眼紙風 2 表の誤結合
- `table-basic-sample13.xlsx`
  - 方眼紙風に merge を多用した表が縦横に 4 つ並ぶケース
  - 2x2 配置でも各表を独立して検出できるか確認する
  - 対応章: `xlsx2md-spec.md` 7, 8, 13
  - 主に確認する症状: merge 多用の方眼紙風 4 表の誤結合
- `table-basic-sample14.xlsx`
  - 方眼紙風に merge を多用した単表で、一部に結合漏れセルがあるケース
  - 多少の merge 崩れがあっても表全体を 1 表として扱えるか確認する
  - 対応章: `xlsx2md-spec.md` 7, 8, 13
  - 主に確認する症状: merge 多用表で一部だけ結合漏れがある場合の崩れ
- `table-basic-sample15.xlsx`
  - 方眼紙風に merge を多用した単表で、備考列に縦結合が混ざるケース
  - `MERGED↑` を含む表でも Markdown 表として壊れないか確認する
  - 対応章: `xlsx2md-spec.md` 7, 8, 13
  - 主に確認する症状: merge 多用表で縦結合が混じる場合の崩れ
- `table-border-priority-sample01.xlsx`
  - 罫線のない dense な 2x2 値ブロック
  - `balanced` では表候補になりやすく、`border` では narrative に落ちる差分を確認する
  - 対応章: `xlsx2md-spec.md` 7, 8
  - 主に確認する症状: 非罫線 fallback による誤検知

## 作成予定 fixture

### `formula/formula-basic-sample01.xlsx`

- 目的: 基本数式
- 対応章: `xlsx2md-spec.md` 10, 11
- 主に確認する症状: `cached / resolved / fallback` の差分
- 構成: 1シート
- 含めたい式:
  - `=A1`
  - `=A1+B1`
  - `IF`
  - `SUM`
  - `COUNTIF`
  - `TEXT`
  - `DATE`
  - `VALUE`
- セル配置案:
  - `A1`: `基本数式サンプル`
  - `A3`: `base1`
  - `B3`: `10`
  - `A4`: `base2`
  - `B4`: `5`
  - `A5`: `ref`
  - `B5`: `=B3`
  - `A6`: `arith`
  - `B6`: `=B3+B4`
  - `A7`: `if`
  - `B7`: `=IF(B3>B4,"OK","NG")`
  - `A8`: `sum`
  - `B8`: `=SUM(B3:B4)`
  - `A9`: `countif`
  - `B9`: `=COUNTIF(B3:B4,">7")`
  - `A10`: `text`
  - `B10`: `=TEXT(B3,"0000")`
  - `D3`: `date`
  - `E3`: `=DATE(2024,3,17)`
  - `D4`: `value_num`
  - `E4`: `=VALUE("1,234.5")`
  - `D5`: `value_date`
  - `E5`: `=VALUE("2024/03/17")`

### `formula/formula-crosssheet-sample01.xlsx`

- 目的: 複数シート参照
- 対応章: `xlsx2md-spec.md` 10, 19
- 主に確認する症状: sheet 参照解決漏れ、参照先ずれ
- 構成: 2シート以上
- セル配置案:
  - シート:
    - `Sheet1`
    - `Sheet2`
    - `日本語シート`
  - `Sheet2!A1`: `1`
  - `Sheet2!B1`: `2`
  - `Sheet2!A2`: `3`
  - `Sheet2!B2`: `4`
  - `Sheet2!B3`: `CrossValue`
  - `日本語シート!C4`: `日本語参照値`
  - `Sheet1!A1`: `複数シート参照サンプル`
  - `Sheet1!A3`: `sheet2_ref`
  - `Sheet1!B3`: `=Sheet2!B3`
  - `Sheet1!A4`: `jp_sheet_ref`
  - `Sheet1!B4`: `='日本語シート'!C4`
  - `Sheet1!A5`: `sum_range`
  - `Sheet1!B5`: `=SUM(Sheet2!A1:B2)`
- 補足:
  - 余裕があれば空白入りシート名も追加する

### `formula/formula-shared-sample01.xlsx`

- 目的: shared formula
- 対応章: `xlsx2md-spec.md` 10, 19
- 主に確認する症状: shared formula 展開漏れ、連番列の崩れ
- 構成: 1シート
- セル配置案:
  - `A1`: `No`
  - `B1`: `連番`
  - `A2:A11`: `1..10`
  - `B2`: `1`
  - `B3`: `=B2+1`
  - `B3:B11`: オートフィル
  - `D1`: `shared formula サンプル`
- 補足:
  - コピー貼り付けではなく、Excel のオートフィルで増やす

  - `E4`: `=SUM(C4#)`
- 補足:
  - Excel で実際に spill させて保存する
  - もし `=A4:A6` だけで spill しない場合は、Microsoft 365 / Excel for web で dynamic array が有効な状態で作成する
  - 可能なら worksheet XML の `f@ref` が残ることを確認したい

### `named-range/named-range-sample01.xlsx`

- 目的: `definedNames`
- 対応章: `xlsx2md-spec.md` 10, 19
- 主に確認する症状: 名前定義解決漏れ、scope 誤解決
- 構成: 2シート
- セル配置案:
  - シート:
    - `Summary`
    - `Other`
  - `Summary!A1`: `definedNames サンプル`
  - `Summary!A3`: `BaseName元`
  - `Summary!B3`: `Base`
  - `Summary!A4`: `BaseRange1`
  - `Summary!B4`: `10`
  - `Summary!A5`: `BaseRange2`
  - `Summary!B5`: `20`
  - workbook スコープ名:
    - `BaseName=Summary!$B$3`
    - `BaseRange=Summary!$B$4:$B$5`
  - `Other!A1`: `LocalCross元`
  - `Other!B2`: `CrossRef`
  - sheet スコープ名:
    - `LocalCross=Other!$B$2`
  - `Summary!D3`: `=BaseName`
  - `Summary!D4`: `=SUM(BaseRange)`
  - `Other!D2`: `=LocalCross`

### `narrative/narrative-vs-table-sample01.xlsx`

- 目的: 地の文と表の判定
- 対応章: `xlsx2md-spec.md` 6, 7, 8
- 主に確認する症状: 表への過剰吸い込み、narrative 二重出力
- 構成: 1シート
- セル配置案:
  - `A1`: 太字 `地の文と表の判定`
  - `A3`: `この設計書は受注入力画面を説明する。`
  - `A4`: `外部システムとの連携条件を以下に示す。`
  - `A5`: `本文は罫線なしのままにする。`
  - `A7`: 太字 `項目一覧`
  - `B8:F11`: 罫線あり表
  - `B8:F8`: `項番 / 項目名称 / 物理名 / 初期値 / 備考`
  - `B9:F11`: 2-3行のデータ
  - `A13`: `※注記: この表はサンプルです。`

### `image/image-basic-sample01.xlsx`

- 目的: 画像抽出
- 対応章: `xlsx2md-spec.md` 14, 15
- 主に確認する症状: 画像 asset 抽出漏れ、anchor ずれ
- 構成: 1シート
- セル配置案:
  - `A1`: `画像抽出サンプル`
  - `B3:F6`: 簡単な表
  - `A7`: `画像サンプル`
  - 画像1枚目: `C8` 付近
  - 画像2枚目: 置けるなら `F8` または `C15` 付近
- 補足:
  - 1枚はサイズ変更してもよい

### `image/image-basic-sample02.xlsx`

- 目的: 画像とグラフの共存確認
- 対応章: `xlsx2md-spec.md` 5, 14, 15
- 主に確認する症状: image / chart の混在崩れ、drawing rels 誤解決
- 構成: 1シート
- セル配置案:
  - `A1`: `画像とグラフの共存サンプル`
  - `B3:F6`: 簡単な元データ表
    - 例: `項目 / 値A / 値B`
  - 画像1枚目: `H3` 付近
  - グラフ1個目: `B9` 付近
- 確認したいこと:
  - `image` と `chart` が同じ sheet の drawing に共存しても壊れないこと
  - 画像は `## 画像`
  - グラフは `## グラフ`
    として別々に抽出できること
- 補足:
  - グラフは棒グラフ 1 個で十分
  - 画像とグラフのアンカーが離れている方が見やすい

### `chart/chart-basic-sample01.xlsx`

- 目的: グラフメタデータ抽出の最小確認
- 対応章: `xlsx2md-spec.md` 5, 15, 19
- 主に確認する症状: chart title / series / type の欠落
- 構成: 1シート
- シート名案:
  - `chart-basic`
- セル配置案:
  - `A1`: `グラフ基本サンプル`
  - `B3:D7`: 元データ表
  - `B3:D3`: `項目 / 系列A / 系列B`
  - `B4:D7`: 3-4行のデータ
  - グラフ1個目: `B10` 付近
- グラフ条件:
  - タイトルあり
  - 系列は 2 本
  - 棒グラフ 1 個
- 確認したいこと:
  - anchor
  - title
  - chart type
  - series name
  - category range ref
  - value range ref
  が Markdown に出ること
- 補足:
  - まずは画像キャッシュではなく chart XML から意味情報を拾えることが目的

### `chart/chart-mixed-sample01.xlsx`

- 目的: 複合グラフ・多系列確認
- 主に確認する症状: 複合グラフ種別判定漏れ、副軸系列の欠落
- 構成: 1シート
- シート名案:
  - `chart-mixed`
- セル配置案:
  - `A1`: `複合グラフサンプル`
  - `B3:E8`: 元データ表
  - `B3:E3`: `項目 / 売上 / 割引額 / 利益率`
  - `B4:E8`: 4-5行のデータ
  - グラフ1個目: `B10` 付近
- グラフ条件:
  - 棒 + 折れ線 の複合グラフ
  - 可能なら副軸あり
- 確認したいこと:
  - `chartType` が単一種類ではなく複数種類を持つ場合でも壊れないこと
  - 系列 3 本程度でも Markdown 化できること
- 補足:
  - 副軸そのものの完全意味解釈は初段では不要
  - まずは `複合` と系列参照が落ちないことを確認したい

### `shape/shape-basic-sample01.xlsx`

- 目的: 図形存在時の安全確認
- 構成: 1シート
- シート名案:
  - `shape-basic`
- セル配置案:
  - `A1`: `図形サンプル`
  - `B3:E6`: 簡単な表
  - `H3` 付近: テキストボックス
  - `H8` 付近: 矢印
  - `K3` 付近: 四角形
- 確認したいこと:
  - 図形が workbook 内に存在しても parse 全体が壊れないこと
  - 画像でも chart でもない drawing 要素を、少なくとも安全に無視または別扱いできること
- 補足:
  - 初段では図形の意味抽出までは求めない
  - まずは「壊れない」「誤って画像扱いしない」を確認する

### `shape/shape-flowchart-sample01.xlsx`

- 目的: フローチャート系図形の raw dump / SVG / 図ブロック clustering 確認
- 構成: 1シート
- シート名案:
  - `shape-flowchart`
- セル配置案:
  - `A1`: `フローチャート図形サンプル`
  - `B3:E6`: 簡単な表
  - `H3` 付近: 開始/終了
  - `K3` 付近: 処理
  - `N3` 付近: 条件判断
  - `Q3` 付近: データ
  - 各図形の間を connector / 矢印で結ぶ
- 確認したいこと:
  - `a:prstGeom@prst` が flowchart 系でも raw dump に残ること
  - 単純図形として SVG 化できるものがあれば assets に出ること
  - 近接配置された図形群が 1 つの `図ブロック` にまとまること
- 補足:
  - 図形数は 4-6 個程度で十分
  - テキスト入り図形を 1-2 個含めると確認しやすい

### `shape/shape-block-arrow-sample01.xlsx`

- 目的: ブロック矢印系図形の raw dump / SVG / 図ブロック clustering 確認
- 構成: 1シート
- シート名案:
  - `shape-block-arrow`
- セル配置案:
  - `A1`: `ブロック矢印サンプル`
  - `H3` 付近: 右矢印
  - `K3` 付近: 左右矢印
  - `N3` 付近: 上矢印
  - `Q3` 付近: U ターン矢印 または 曲線矢印
  - 可能なら `H8` 付近に別系統の矢印を 1-2 個
- 確認したいこと:
  - 矢印系 `prstGeom` が raw dump に残ること
  - 単純矢印として SVG 化できるものがあれば assets に出ること
  - 上段のまとまりと下段のまとまりが別 `図ブロック` になり得ること
- 補足:
  - connector ではなく、Office の「ブロック矢印」を優先する
  - 方向や形の違いが出るように 4-6 個程度置く

### `shape/shape-callout-sample01.xlsx`

- 目的: 吹き出し系図形の raw dump / テキスト抽出 / 図ブロック clustering 確認
- 構成: 1シート
- シート名案:
  - `shape-callout`
- セル配置案:
  - `A1`: `吹き出しサンプル`
  - `H3` 付近: 角丸吹き出し
  - `K3` 付近: 楕円吹き出し
  - `N3` 付近: 雲形吹き出し
  - 各吹き出しには短いテキストを入れる
  - 可能なら `H8` 付近に注釈用の別吹き出しを追加
- 確認したいこと:
  - `a:t` が図形テキストとして raw dump に残ること
  - 吹き出し系 shape でも図形テキストを Markdown で読めること
  - 近接した吹き出し群が 1 つの `図ブロック` にまとまること
- 補足:
  - 初段では吹き出し形状そのものの完全 SVG 再現までは求めない
  - まずは raw dump とテキスト抽出を重視する

### `edge/edge-empty-sample01.xlsx`

- 目的: 空系の境界
- 構成: 1シート
- セル配置案:
  - `A1`: `空系境界サンプル`
  - `C7`: `only-value`
- 補足:
  - それ以外は空のままにする
  - 罫線、結合、画像は入れない

### `edge/edge-weird-sheetname-sample01.xlsx`

- 目的: ファイル名サニタイズ
- 構成: 1シート
- 補足:
  - Excel は `/ \ ? * : [ ]` をシート名に使えない
  - そのため、Excel で許される範囲で揺れやすい名前を使う
- シート名案:
  - `A B-東京&大阪.01`
- セル配置案:
  - `A1:D1`: `項番 / 名称 / 値 / 備考`
  - `A2:D4`: 2-3行のデータ

## 補足

- さらに細かい方針や広いバックログは [TODO.md](../../../TODO.md) を参照
