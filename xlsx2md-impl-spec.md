# xlsx2md Implementation Specification

## 1. 文書の位置づけ

この文書は、`xlsx2md` の現行実装が実際にどのように動作するかを整理するための実装仕様書である。

主な目的は次の通りである。

- `docs/xlsx2md/src/xlsx2md/ts/*.ts` の現行実装に基づく挙動を整理する
- `README.md`、`xlsx2md-spec.md`、`xlsx-formula-subset.md` に分散している実装前提の情報を、実装準拠の観点で補助する
- 将来の仕様検討や実装差分確認の際に、どこが「現行挙動」で、どこが「将来構想」かを切り分けやすくする

本書は、上位方針や理想仕様を定めるための文書ではなく、現時点の実装挙動を記述する文書である。

したがって、上位方針や将来構想については [xlsx2md-spec.md](./xlsx2md-spec.md) を優先し、本書では実装済みの範囲と実際の処理順を優先して記述する。

### 1.1 関連文書との役割分担

`xlsx2md` 関連文書の役割分担は概ね次の通りである。

- [README.md](./README.md)
  - 概要、使い方、現在の状態、既知の制約を把握するための入口文書
- [xlsx2md-spec.md](./xlsx2md-spec.md)
  - 上位仕様、設計方針、将来構想を整理する文書
- [xlsx2md-impl-spec.md](./xlsx2md-impl-spec.md)
  - 現行実装の実際の挙動、内部モデル、処理順を記述する文書
- [xlsx-formula-subset.md](./xlsx-formula-subset.md)
  - Excel 数式サブセットおよび evaluator / resolver 周辺の検討を補助する文書

実装差分の確認や挙動確認を行う場合は、本書と TypeScript 実装を優先して参照する。

## 2. 対象範囲

### 2.1 対応入力

現行実装が主対象とする入力は、Excel Workbook 形式の `.xlsx` ファイルである。

入力はブラウザ上でローカルファイルとして読み込まれ、ZIP 展開後に内部 XML を参照して解析する。

### 2.2 非対応入力

現行実装では、少なくとも以下は対象外である。

- `.xls`
- `.csv`
- 外部 Workbook を前提とする完全参照解決

また、Excel のすべての機能を完全再現することは目的としていない。

### 2.3 本文書の基準

本書の記述基準は次の通りである。

- 現行コードの正本は `docs/xlsx2md/src/xlsx2md/ts/` 配下の TypeScript 実装とする
- `js/` 配下および `xlsx2md.html` はビルド生成物として扱う
- 実装と既存文書に差分がある場合、本書では現行実装を基準に記述する
- 未対応事項や将来検討事項は、実装済み仕様と分けて記述する

また、本書では特に断りがない限り、`display` モードを標準的な出力モードとして扱う。

## 3. 全体処理フロー

`xlsx2md` の現行実装における全体処理フローは概ね次の通りである。

1. ユーザーがブラウザ UI で `.xlsx` ファイルを選択する
2. 入力ファイルを ArrayBuffer として読み込む
3. `.xlsx` を ZIP として展開し、必要な XML ファイルを取得する
4. `workbook.xml`、worksheet XML、rels、`sharedStrings.xml`、`styles.xml` などを解析する
5. Workbook / Sheet / Cell / Merge / Table / Image / Chart / Shape の内部モデルを構築する
6. 数式セルについて、`cached value`、AST evaluator、従来 resolver などを用いて解決を試みる
7. Sheet ごとに表候補、地の文ブロック、補助セクションを抽出する
8. `display / raw / both` の出力モードに応じて Markdown を組み立てる
9. 画面上に解析サマリー、表候補スコア、数式診断、Markdown プレビューを表示する
10. 必要に応じて Markdown 一括保存または ZIP 保存を行う

この流れのうち、解析本体は主に `core.ts`、画面表示や操作導線は主に `main.ts` が担う。

## 4. Workbook 解析仕様

### 4.1 ZIP / XML 読み込み

`.xlsx` は ZIP アーカイブとして扱う。

現行実装では、入力された ArrayBuffer を直接走査し、中央ディレクトリおよびローカルヘッダを読んで各エントリを取得する。圧縮方式は少なくとも次を扱う。

- store
- deflate-raw

ZIP 展開後は、必要なエントリを `Map<string, Uint8Array>` として保持し、XML 系エントリは UTF-8 文字列へデコードしたうえで DOM として解析する。

XML 解析は `DOMParser` を用い、namespace 接頭辞そのものではなく localName を基準に辿る実装を併用する。

### 4.2 workbook / worksheet / rels の解決

Workbook の起点は `xl/workbook.xml` とする。

現行実装では、少なくとも次の手順で workbook / worksheet の対応付けを行う。

1. `xl/workbook.xml` を読み込む
2. `sheet` 要素を列挙し、Workbook 上のシート順・シート名を取得する
3. `xl/_rels/workbook.xml.rels` を読み込み、`r:id` から worksheet XML への相対パスを解決する
4. 各 worksheet XML を個別に解析する

rels の解決では、相対パス、`..`、同一ディレクトリ基準の正規化を行う。

worksheet 側でも同様に `.rels` を参照し、少なくとも次の関連リソースを解決する。

- table
- drawing

drawing 経由で、さらに画像、グラフ、図形を辿る。

### 4.3 sharedStrings / styles / definedNames

#### sharedStrings

`xl/sharedStrings.xml` が存在する場合、`si` 要素を順番に読み取り、共有文字列配列を構築する。

文字列抽出では、単純な `t` 要素だけでなく rich text 断片も連結対象とする。一方で phonetic 系の要素は読み飛ばす。

#### styles

`xl/styles.xml` が存在する場合、少なくとも以下を読み取る。

- border
- numFmt
- cellXfs

これにより、各セルに対して少なくとも次のスタイル情報を紐付ける。

- `borders`
  - `top`
  - `bottom`
  - `left`
  - `right`
- `numFmtId`
- `formatCode`

`styles.xml` が存在しない場合は、`General` 相当の既定値を使う。

#### definedNames

`workbook.xml` の `definedName` 要素を読み取り、名前定義一覧を構築する。

現行実装では、少なくとも次を保持する。

- `name`
- `formulaText`
- `localSheetName`

`localSheetId` を持つ名前は sheet scope name として扱い、持たない名前は workbook scope name として扱う。

また、`_xlnm.` で始まる組み込み名前は通常の名前定義一覧から除外する。

## 5. Sheet モデル仕様

### 5.1 cells

各 worksheet から抽出したセルは、Sheet モデル内で `cells` 配列として保持する。

現行実装のセルモデルは、少なくとも次の情報を持つ。

- `address`
- `row`
- `col`
- `valueType`
- `rawValue`
- `outputValue`
- `formulaText`
- `resolutionStatus`
- `resolutionSource`
- `cachedValueState`
- `styleIndex`
- `borders`
- `numFmtId`
- `formatCode`
- `formulaType`
- `spillRef`

このうち `rawValue` はセル内部値寄り、`outputValue` は表示値寄りまたは Markdown 出力寄りの値として保持する。

また、数式セルについては `formulaText` の有無により通常セルと区別される。

### 5.2 merges

結合セル範囲は `merges` 配列として保持する。

各要素は、少なくとも次を持つ。

- `startRow`
- `startCol`
- `endRow`
- `endCol`
- `ref`

`mergeCell` 要素の `ref` を起点として、矩形範囲へ展開して保持する。

### 5.3 tables

Excel table として定義されている領域は `tables` 配列として保持する。

各要素は、少なくとも次を持つ。

- `sheetName`
- `name`
- `displayName`
- `start`
- `end`
- `columns`
- `headerRowCount`
- `totalsRowCount`

この情報は、構造化参照の解決や表関連メタ情報の補助に利用される。

ただし、Markdown 出力上の表は table 定義だけで決まるのではなく、別途の表候補検出も併用する。

### 5.4 images

埋め込み画像は `images` 配列として保持する。

各要素は、少なくとも次を持つ。

- `sheetName`
- `filename`
- `path`
- `anchor`
- `data`
- `mediaPath`

`anchor` は drawing のアンカーセルを A1 形式へ変換したものとする。

`path` は Markdown および ZIP 出力で参照する相対パスであり、`assets/<sheet>/...` 形式を用いる。

### 5.5 charts

drawing 内で検出された chart は `charts` 配列として保持する。

各要素は、少なくとも次を持つ。

- `sheetName`
- `anchor`
- `chartPath`
- `title`
- `chartType`
- `series`

`series` は系列ごとの参照情報を持ち、少なくとも次を含む。

- `name`
- `categoriesRef`
- `valuesRef`
- `axis`

グラフの解釈は限定的であり、画像化や Excel 相当描画を行うのではなく、構造情報の抽出を主目的とする。

### 5.6 shapes

drawing 内で検出された図形は `shapes` 配列として保持する。

各要素は、少なくとも次を持つ。

- `sheetName`
- `anchor`
- `name`
- `kind`
- `text`
- `widthEmu`
- `heightEmu`
- `elementName`
- `anchorElementName`
- `rawEntries`

`kind` は、図形要素や `prstGeom`、`txBox` などの情報をもとに推定される簡易分類である。

例:

- テキストボックス
- 長方形
- 直線矢印コネクタ

`rawEntries` は、drawing XML 上の属性やテキストを平坦化して保持したキー・値一覧であり、現行実装では図形の詳細情報を Markdown へ出す際の基礎データとしても利用する。

## 6. セル値処理仕様

### 6.1 通常セル

通常セルでは、セル型とスタイル情報に応じて `rawValue` と `outputValue` を構築する。

現行実装では、少なくとも次の型を区別する。

- shared string (`t="s"`)
- inline string (`t="inlineStr"`)
- boolean (`t="b"`)
- string (`t="str"`)
- error (`t="e"`)
- 数値または型未指定

基本方針は次の通りである。

- `rawValue` は XML 上の値を基準に保持する
- `outputValue` は可能な範囲で表示値寄りに整形する
- 文字列セルでは、基本的に `rawValue` と `outputValue` は同じになる
- 数値や日付系セルでは、書式に応じて `outputValue` が整形されることがある

### 6.2 表示形式の適用

数値系セルについては、`numFmtId` と `formatCode` をもとに表示形式を適用する。

現行実装で重点的に扱うのは、少なくとも次の系統である。

- 整数・小数
- 桁区切り
- 通貨
- 会計
- 日付
- 時刻
- パーセンテージ
- 分数
- 指数
- 一部の和暦・ユーザー定義書式

表示形式適用は完全な Excel 互換ではなく、Markdown 出力上で実用上重要な表示を優先する。

たとえば、次のような処理を行う。

- 日付シリアル値を日付文字列へ変換する
- `%` を含む書式はパーセンテージ表示へ寄せる
- `?/?` 系書式は分数表示へ寄せる
- scientific notation 相当の書式は指数表記へ寄せる
- 会計書式のゼロ値は `¥ -` のような表示を返すことがある

数式セルについても、解決後の値に対して同様の表示形式適用を試みる。

### 6.3 outputMode ごとの出力値

Markdown 生成時のセル出力は、`outputMode` により切り替わる。

現行実装で扱うモードは次の 3 つである。

- `display`
- `raw`
- `both`

各モードの扱いは次の通りである。

#### display

- `outputValue` を優先して出力する
- 人間が Excel 上で見ている表示値に近い形を優先する

#### raw

- `rawValue` を優先して出力する
- 表示形式適用前の内部値寄りの情報を確認したい場合に用いる

#### both

- `outputValue` を本文値として出力する
- `rawValue` が `outputValue` と異なる場合に限り、`[raw=...]` を補助表示する
- `rawValue` と `outputValue` が同じ場合は、補助表示を付けない

この切り替えは Markdown テーブル、地の文、数式を含む各種セル出力に共通で使われる。

## 7. 数式セル処理仕様

### 7.1 基本方針

数式セルは、可能な限り値として解決する。ただし、Excel 数式全体の完全互換を目的とするものではなく、Markdown 化に必要な範囲で実用的に補完する方針を採る。

現行実装では、数式セルごとに少なくとも次を保持する。

- `formulaText`
- `resolutionStatus`
- `resolutionSource`
- `cachedValueState`
- `rawValue`
- `outputValue`

また、数式セルの最終出力では、値の解決そのものと、表示形式の再適用の両方を考慮する。

### 7.2 cached value の扱い

現行実装では、数式セルに `v` 要素が存在する場合、それを `cached value` ありとして扱う。

この判定は `v` 要素の中身が非空かどうかではなく、`v` 要素の存在有無で決まる。

したがって、次のようなケースも `cached value` 採用とみなす。

- `<f>...</f><v></v>`

この場合、`cachedValueState` は `present_empty` となり、`resolutionSource` は `cached_value` となる。

一方で、`v` 要素自体が存在しない数式セルは cached 不在とみなし、以降の自前解決へ進む。

### 7.3 解決順序

現行実装における数式解決順序は、概ね次の通りである。

1. `cached value`
2. AST evaluator
3. 従来の文字列ベース resolver
4. 式文字列保持

実際の内部メタでは、少なくとも次の経路を区別する。

- `cached_value`
- `ast_evaluator`
- `legacy_resolver`
- `formula_text`
- `external_unsupported`

`resolutionStatus` は「最終状態」を表し、`resolutionSource` は「どの経路でその状態に至ったか」を表す。

### 7.4 AST evaluator

AST evaluator は、`docs/xlsx2md/src/xlsx2md/ts/formula/` 配下の tokenizer / parser / evaluator を用いた解決系である。

現行実装では、数式文字列を AST 化し、必要に応じて次のような参照解決コンテキストを与えて評価する。

- セル参照
- 範囲参照
- 名前定義
- sheet scope name
- 構造化参照
- 現在セル参照

AST evaluator で解決できた場合、`resolutionSource` は `ast_evaluator` となる。

AST evaluator は優先度上、cached 不在時に最初に試行される自前解決経路である。

### 7.5 legacy resolver

legacy resolver は、AST evaluator で解決できない場合に使う従来の文字列ベース解決系である。

現行実装では、少なくとも次のような処理を含む。

- 直接参照の解決
- `IF`、`IFERROR`、論理式の解決
- 文字列連結
- 数値関数・日付関数・文字列関数の一部解決
- 集計関数・条件付き集計関数の一部解決
- `VLOOKUP` / `HLOOKUP` / `XLOOKUP` / `MATCH` / `INDEX` の一部解決
- 算術式への簡易置換と評価

legacy resolver で解決できた場合、`resolutionSource` は `legacy_resolver` となる。

現時点では、legacy resolver は後方互換の安全装置であると同時に、実戦データ上で実際に使われる解決経路でもある。

### 7.6 fallback_formula

cached 不在であり、AST evaluator と legacy resolver のいずれでも解決できない場合、数式セルは式文字列保持へフォールバックする。

この場合、少なくとも次の状態となる。

- `resolutionStatus = fallback_formula`
- `resolutionSource = formula_text`

出力値は空欄にせず、式文字列をそのまま保持する。

### 7.7 外部参照

外部 Workbook を参照する式は、現行実装では解決対象外とする。

たとえば、式文字列に `[...]` を含む external workbook 参照がある場合、少なくとも次の状態として扱う。

- `resolutionStatus = unsupported_external`
- `resolutionSource = external_unsupported`

この場合、外部参照の再解決は行わず、式文字列保持を優先する。

### 7.8 shared formula

shared formula は worksheet XML の `f` 要素における `t="shared"` と `si` をもとに扱う。

現行実装では、少なくとも次のように処理する。

- 基底となる shared formula を `si` ごとに記録する
- 後続セルでは、基底式とセル位置の差分を使って参照アドレスを平行移動する
- 展開後の式文字列を `formulaText` として扱う

これにより、オートフィル由来の shared formula でも通常の数式セルと同様の解決フローへ流せるようにする。

### 7.9 名前定義・構造化参照

現行実装では、名前定義と構造化参照を数式解決の対象に含める。

#### 名前定義

- workbook scope name を解決対象とする
- sheet scope name を解決対象とする
- 単一セル参照だけでなく、範囲参照も扱う

#### 構造化参照

- table 定義から列位置を求める
- 列参照を実セル範囲へ展開する
- 対応可能な範囲では、範囲行列として AST evaluator や aggregate 解決へ渡す

これにより、`definedNames` や table 参照を含む実務系 Workbook に対して、完全ではないが実用的な解決を目指す。

## 8. 数式診断仕様

### 8.1 診断対象

数式診断は、Sheet 内の `cells` のうち `formulaText` を持ち、かつ `resolutionStatus !== null` のセルを対象とする。

したがって、診断一覧に載るのは「数式セルであり、解決状態を持つセル」である。

この一覧は「再計算が必要だったセルの一覧」ではなく、「数式セルの解決状態一覧」である。

そのため、`cached value` をそのまま採用した数式セルも、数式診断の対象に含まれる。

### 8.2 status

数式診断では、少なくとも次の `status` を扱う。

- `resolved`
- `fallback_formula`
- `unsupported_external`

各意味は次の通りである。

#### resolved

最終的に値として解決できたことを表す。

ただし、この `resolved` は解決経路を区別しない。つまり、次のどちらでも `resolved` になりうる。

- `cached value` 採用
- AST evaluator または legacy resolver による自前解決

#### fallback_formula

値として解決できず、式文字列保持へフォールバックしたことを表す。

#### unsupported_external

外部参照など、現行実装の解決対象外として扱ったことを表す。

### 8.3 source

`source` は、どの経路で現在の出力値に到達したかを示す補助メタ情報である。

現行実装では、少なくとも次を扱う。

- `cached_value`
- `ast_evaluator`
- `legacy_resolver`
- `formula_text`
- `external_unsupported`

各意味は次の通りである。

#### cached_value

保存済み `cached value` を採用したことを表す。

#### ast_evaluator

AST evaluator により解決したことを表す。

#### legacy_resolver

従来の文字列ベース resolver により解決したことを表す。

#### formula_text

式文字列保持へフォールバックしたことを表す。

#### external_unsupported

外部参照などの非対応経路として扱ったことを表す。

`status` と `source` は別軸であり、たとえば `status = resolved` かつ `source = cached_value` のような組み合わせをとる。

### 8.4 cachedValueState

`cachedValueState` は、数式セルの `v` 要素の状態を区別するための内部メタ情報である。

現行実装では、少なくとも次を扱う。

- `present_nonempty`
- `present_empty`
- `absent`

各意味は次の通りである。

#### present_nonempty

`v` 要素が存在し、中身が空文字ではない。

#### present_empty

`v` 要素が存在し、中身が空文字である。

この場合でも、現行実装では `cached value` ありとして扱う。

#### absent

`v` 要素が存在しない。

この場合は cached 不在とみなし、AST evaluator または legacy resolver などの自前解決へ進む。

`cachedValueState` は現在の診断 UI では主表示項目ではないが、挙動説明や内部確認には重要な補助情報である。

### 8.5 診断表示

画面 UI の数式診断では、少なくとも次の観点を表示する。

- 全体件数
- `source` 別件数
  - `cached`
  - `ast`
  - `legacy`
  - `formula`
- `status` 別件数
  - `resolved`
  - `fallback`
  - `unsupported`
- シートごとの数式診断一覧

個別の診断カードでは、少なくとも次を表示する。

- セルアドレス
- `source` ピル
- `status` ピル
- `formulaText => outputValue`

この表示により、たとえば次のような区別が可能になる。

- `cached` でそのまま採用された数式
- 自前評価で補完された数式
- 式文字列保持へ落ちた数式

なお、現行表示では `cachedValueState` は一覧カード上へ直接は出していない。

## 9. 表検出仕様

### 9.1 基本方針

現行実装の表検出は、worksheet 全体から表候補を自動抽出し、スコアリングにより採用可否を判定する方式である。

この表検出は、Excel table 定義そのものに完全依存するのではなく、セルの値配置と罫線情報から「表らしい矩形領域」を推定する。

そのため、ここでいう表候補は「Markdown テーブルとして出力すべき領域候補」であり、Workbook 内の table object と必ずしも一致しない。

また、現行実装では理想仕様上の記述よりも、連結成分ベースの領域抽出が強く効く。

### 9.2 seed cell

表候補抽出の起点となるセルを `seed cell` と呼ぶ。

現行実装では、少なくとも次のどちらかを満たすセルを seed cell として採用する。

- `outputValue` が空でない
- いずれかの辺に罫線がある

逆に、値も罫線も持たないセルは表候補抽出の対象外とする。

この設計により、罫線主体の表だけでなく、値が密に並ぶ領域も候補になりうる。

### 9.3 連結成分の抽出

seed cell を対象に、上下左右の 4 近傍で隣接しているセルを連結成分としてまとめる。

現行実装では、概ね次の流れで処理する。

1. seed cell の位置マップを作る
2. 未訪問の seed cell から探索を開始する
3. 上下左右に隣接する seed cell をたどって同一成分へまとめる
4. 成分全体の最小行・最小列・最大行・最大列を求める
5. その外接矩形を表候補領域として扱う

このため、現行実装の表候補は「罫線だけで形成された矩形」ではなく、「値または罫線を持つセルの隣接成分を矩形化した領域」である。

また、外接矩形の行数または列数が 2 未満のものは表候補から除外する。

### 9.4 スコアリング

表候補の採用可否は、重み付きスコアリングで判定する。

現行実装で使う主な観点は次の通りである。

- 最小グリッド条件
  - 2 行 x 2 列以上
- 罫線の存在
- 非空セル密度
- 先頭行のヘッダらしさ
- 結合セル過多による減点
- 長文中心領域による減点

少なくとも次のような加点・減点を行う。

- 2x2 以上の矩形である
- 罫線セルの比率が一定以上である
- 密度が高い
- 先頭行に短い非数値文字列が複数ある
- 結合セルが多すぎる
- 長文中心かつ密度が低い

このスコアリングは学習ベースではなく、固定重みのヒューリスティックである。

### 9.5 表候補採用条件

スコアが閾値以上の候補だけを表として採用する。

現行実装では、閾値は固定値であり、候補ごとに少なくとも次を保持する。

- `startRow`
- `startCol`
- `endRow`
- `endCol`
- `score`
- `reasonSummary`

採用後は、この矩形領域を Markdown テーブルへ変換する。

表として採用されなかったセル群は、後続の地の文抽出側へ回る。

なお、現行実装では次の点に注意が必要である。

- 「表検出の主要手掛かりは罫線」とする上位方針に対し、実装上は seed cell と連結成分の影響が大きい
- そのため、値が密に並ぶが罫線が弱い領域も表候補になりうる
- 一方で、巨大レイアウト系シートでは通常表検出だけでは自然な結果にならないことがある

## 10. 地の文抽出仕様

### 10.1 narrative block

地の文抽出は、表として採用されなかったセル群を対象に行う。

現行実装では、少なくとも次の条件を満たすセルを地の文候補とする。

- `outputValue` が空でない
- 採用済み表候補の内部に含まれない

そのうえで、行番号ごとにセルをまとめ、同一行内では列順に並べて 1 行のテキストへ連結する。

行テキスト生成の基本方針は次の通りである。

- 同一行のセル値を左から右へ連結する
- セル間は半角スペースでつなぐ
- 空文字は除外する

その後、行単位のテキストを近接性に基づいて narrative block へまとめる。

現行実装では、少なくとも次の条件で block を分割する。

- 前行との行差が 1 を超える
- 開始列の差が一定以上大きい

この block は、後続で段落またはリストとしてレンダリングされる。

### 10.2 リスト化条件

現行実装では、narrative block 内の各行を `items` として持ち、一定条件を満たす連続行をリストとして Markdown 箇条書きへ変換する。

リスト候補判定はヒューリスティックであり、少なくとも次のような観点を用いる。

- 1 行の長さが極端に短すぎず長すぎない
- 先頭列が短いマーカーや状態値として見える
- 複数セルからなり、後続セルを本文として読める
- 1 セルだけでも、短文の独立項目として読みやすい
- 見出しや段落末尾だけに見える行は除外する

さらに、候補行が連続しており、一定件数以上まとまっている場合にリストブロックとして扱う。

現行実装では、少なくとも次のような出力を行う。

- 通常箇条書き
  - `- ...`
- チェック済み項目
  - `- [x] ...`
- 未チェック項目
  - `- [ ] ...`

行列構造やセル値の並びを使って、マーカー列と本文列を分離する場合がある。

### 10.3 表との優先関係

地の文抽出は、表検出の後に行う。

そのため、あるセルが採用済み表候補の内部に含まれる場合、そのセルは地の文候補には入らない。

優先関係は概ね次の通りである。

1. 表候補として採用された領域
2. 表外セルから作る narrative block
3. narrative block の中でのリスト化判定

この順序により、同じセルを表と地の文の両方へ重複出力しないようにしている。

一方で、現行実装は「採用済み表候補を除外した残り」を地の文へ回す方式であるため、表検出の結果が地の文の見え方へ直接影響する。

たとえば、次のようなケースでは注意が必要である。

- 値の密な領域が表として採用され、本来地の文に近い内容が表へ吸われる
- 逆に、表候補閾値に届かない領域が narrative block として出力される

この点は、現行実装の表検出ヒューリスティックと地の文抽出ヒューリスティックが密接に結びついていることを意味する。

## 11. 結合セル仕様

### 11.1 代表セル

結合セル範囲は Markdown 上で Excel と同じ見た目には再現しない。

現行実装では、結合範囲の左上セルを代表セルとして扱う。

そのため、結合範囲内では少なくとも次のような役割分担になる。

- 左上セル
  - 元の値を保持する
- 左上以外のセル
  - 補助トークンへ置き換える

結合範囲情報そのものは `merges` として保持され、表描画時に行列へ反映する。

### 11.2 `[MERGED←]` / `[MERGED↑]`

現行実装では、結合セル展開時に補助トークンを用いる。

- 同一行内で左上セルの右側にあるセル
  - `[MERGED←]`
- 下方向に展開されるセル
  - `[MERGED↑]`

複合結合範囲でも、左上セル以外はこのどちらかのトークンへ置き換える。

この処理は、表候補から生成した 2 次元行列に対して後段で適用する。

したがって、内部セルモデルの `outputValue` 自体が補助トークンへ書き換わるのではなく、Markdown 用の表行列に対して補助トークンが注入される。

## 12. 画像仕様

### 12.1 抽出対象

画像抽出は drawing XML をたどって行う。

現行実装では、worksheet に紐づく drawing を解決し、少なくとも次のアンカー要素を対象とする。

- `oneCellAnchor`
- `twoCellAnchor`

この中で `blip` を持つ要素を画像として扱い、関連する media ファイルを抽出する。

画像抽出は OCR や画像内容理解を目的とするものではなく、埋め込み画像の存在と位置を Markdown へ持ち込むことを目的とする。

### 12.2 アンカー

画像位置は、drawing の `from` 要素に含まれる `row` と `col` をもとに A1 形式のアンカーへ変換して保持する。

このアンカーは、少なくとも次の用途に使う。

- 画像情報の内部保持
- Markdown 上の `### 画像NNN (Anchor)` 見出し

ピクセル単位やオフセット単位の再現は主目的ではなく、セルアンカー単位の保持を優先する。

### 12.3 Markdown 出力

画像が存在する場合、Markdown 末尾付近に `## 画像` セクションを追加する。

各画像について、少なくとも次を出力する。

- `### 画像NNN (Anchor)`
- `- File: ...`
- Markdown の画像リンク

出力先パスは、少なくとも `assets/<sheet>/image_XXX.ext` 形式で扱う。

また、ZIP 出力では Markdown 本体とともに画像ファイルも含める。

## 13. グラフ仕様

### 13.1 抽出対象

グラフ抽出も drawing XML を起点として行う。

現行実装では、anchor 内の `graphicFrame` から chart 参照をたどり、chart XML を読み込んでグラフ情報を構築する。

画像と同様、少なくとも `oneCellAnchor` / `twoCellAnchor` に紐づく位置情報を利用する。

### 13.2 保持情報

グラフについては、少なくとも次を保持する。

- `sheetName`
- `anchor`
- `chartPath`
- `title`
- `chartType`
- `series`

`series` では、系列名、カテゴリ参照、値参照、軸種別などを持つ。

この保持方針は、Excel グラフの見た目再現よりも、どの系列がどのセル範囲を参照しているかを可読な形で残すことを重視したものである。

### 13.3 Markdown 出力

グラフが存在する場合、Markdown 末尾付近に `## グラフ` セクションを追加する。

各グラフについて、少なくとも次を出力する。

- `### グラフNNN (Anchor)`
- タイトル
- 種別
- 系列一覧
- 必要に応じて categories / values の参照範囲

副軸系列などは、系列メタ情報として補助的に出力する。

## 14. 図形仕様

### 14.1 抽出対象

図形抽出も drawing XML を起点とする。

現行実装では、anchor 内に存在する少なくとも次の図形要素を対象とする。

- `sp`
- `cxnSp`

ただし、同じ anchor 内でも、画像やグラフとして扱う要素は図形抽出から除外する。

そのため、画像・グラフ・図形の分類は drawing 内で排他的に行われる。

### 14.2 kind / text / ext / rawEntries

図形では、詳細な再描画ではなく、構造情報と属性の保持を優先する。

少なくとも次を抽出対象とする。

- `kind`
  - 形状種別や text box 判定に基づく簡易分類
- `text`
  - 図形内部のテキスト
- `widthEmu`
- `heightEmu`
- `elementName`
- `anchorElementName`
- `rawEntries`

`kind` は、要素名、`txBox` 属性、`prstGeom` などを使って決める。

`rawEntries` は、anchor 以下の XML を平坦化して、`key / value` の配列として保持したものである。

この `rawEntries` により、図形 XML の属性やテキスト内容を Markdown 上でも追跡しやすくしている。

### 14.3 Markdown 出力

図形が存在する場合、Markdown 末尾付近に `## 図形` セクションを追加する。

各図形について、少なくとも次を出力する。

- `### 図形NNN (Anchor)`
- `anchorElement`
- `element`
- `rawEntries` の各項目

現行実装の Markdown 出力は、図形の意味解釈や自然文要約ではなく、XML 由来の構造情報を確認できる形を優先する。

したがって、図形セクションは人間向けの最終表現というより、追跡可能性と解析確認のための補助情報としての性格が強い。

## 15. Markdown 組み立て仕様

### 15.1 セクション順

現行実装では、Sheet 単位の Markdown を次の大きな順で組み立てる。

1. シート見出し
2. ソース情報
3. 本文
4. 補助セクション
  - グラフ
  - 図形
  - 画像

本文セクション内部では、地の文ブロックと表ブロックを同一の配列へ集め、行番号・列番号ベースで並べ替えて出力する。

このため、地の文と表の相対順は、Sheet 上の位置関係に基づいて決まる。

### 15.2 見出し命名

現行実装では、少なくとも次の見出し命名規則を用いる。

- Sheet 見出し
  - `# <sheet name>`
- 表
  - `### 表NNN (Range)`
- 画像
  - `### 画像NNN (Anchor)`
- グラフ
  - `### グラフNNN (Anchor)`
- 図形
  - `### 図形NNN (Anchor)`

表は役割推定名ではなく、連番を基本とする。

地の文ブロック自体には通常個別見出しを付けず、本文の流れの中へ直接配置する。

### 15.3 テーブル描画

表候補として採用された領域は、2 次元行列へ変換したうえで Markdown テーブルとして描画する。

現行実装では、概ね次の処理順を取る。

1. 候補矩形からセル行列を構築する
2. `outputMode` に応じてセル文字列を得る
3. 必要に応じて trim を行う
4. 結合セル補助トークンを適用する
5. 必要に応じて空行・空列を除外する
6. Markdown テーブル記法へ変換する

ヘッダ行の扱いはオプションで制御される。

- `treatFirstRowAsHeader = true`
  - 先頭行をヘッダとして扱う
- `false`
  - 空ヘッダ行を持つ通常表として扱う

セル値中の `|` はエスケープし、改行は `<br>` に変換する。

### 15.4 補助セクション

グラフ、図形、画像は本文本体とは別セクションとして末尾へ追加する。

現行実装での並びは次の通りである。

1. `## グラフ`
2. `## 図形`
3. `## 画像`

これらは Sheet 上の位置へ完全に差し込むのではなく、補助情報として末尾側へまとめて出力する。

そのため、アンカー位置は各節見出しに残すが、本文の流れへ厳密に埋め込むわけではない。

## 16. 出力ファイル仕様

### 16.1 ファイル名

Sheet ごとの Markdown ファイル名は、少なくとも次の要素から組み立てる。

- Workbook 名
- Sheet の物理順
- Sheet 名
- 必要に応じて outputMode サフィックス

基本形は次の通りである。

- `<workbook>_<sheetIndex>_<sheetName>.md`

`raw` または `both` モードでは、少なくとも次のようなサフィックスを付ける。

- `_raw`
- `_both`

Workbook 名と Sheet 名は、ファイル名として不安定な文字や空白を整理したうえでサニタイズする。

### 16.2 all-in-one Markdown

画面上のダウンロード機能では、Sheet ごとの Markdown を 1 つに連結した all-in-one Markdown を生成できる。

現行実装では、各 Sheet の Markdown 断片の間にチャンク識別用コメントを挿入して連結する。

ファイル名は、少なくとも次の形式を用いる。

- `<workbook>.md`
- `<workbook>_raw.md`
- `<workbook>_both.md`

これにより、Sheet 単位ファイルとは別に、Workbook 全体を一括確認できる。

### 16.3 ZIP 出力

ZIP 出力では、Workbook 単位の連結 Markdown 1 本と補助アセットをまとめて格納する。

現行実装では、少なくとも次を含む。

- `output/<combined-file>.md`
- `output/assets/...` 配下の画像
- `output/assets/...` 配下の図形 SVG

ZIP の保存名には Workbook 名と、必要に応じて outputMode サフィックスを付ける。

少なくとも次のような形式を用いる。

- `<workbook>_xlsx2md_export.zip`
- `<workbook>_xlsx2md_export_raw.zip`
- `<workbook>_xlsx2md_export_both.zip`

現行実装では、グラフは Markdown 本文内へメタ情報として書き出す。一方で図形は、Markdown 内の metadata 出力に加えて、SVG 生成対象のものについては個別アセットとして ZIP 内へ追加される。

## 17. UI 上の表示仕様

### 17.1 解析サマリー

画面 UI では、Workbook 全体および Sheet ごとの解析サマリーを表示する。

現行実装では、少なくとも次の情報を表示対象とする。

- Workbook 名
- Sheet 数
- outputMode
- 表数
- 地の文ブロック数
- 結合セル数
- 画像数
- 数式診断件数
- 解析セル数

このサマリーは、Markdown 出力そのものではなく、変換結果の概況を確認するための補助表示である。

### 17.2 表候補スコア

表候補スコア表示では、採用された表候補についてスコアと理由を表示する。

現行実装では、少なくとも次を表示する。

- 全体件数
- strong 件数
- candidate 件数
- Sheet ごとの件数
- 各表候補の Range
- 各表候補の score
- `reasonSummary`

表示ラベルはスコア帯によって少なくとも次を用いる。

- `strong`
- `candidate`
- `unknown`

この UI は、表検出の妥当性を人手確認するためのデバッグ・補助表示としての性格が強い。

### 17.3 数式診断

数式診断表示では、数式セルの解決状態と解決経路を一覧化する。

現行実装では、少なくとも次の表示要素を持つ。

- 全体件数
- `source` 別件数
  - `cached`
  - `ast`
  - `legacy`
  - `formula`
- `status` 別件数
  - `resolved`
  - `fallback`
  - `unsupported`
- Sheet ごとの数式診断件数
- 各セルの診断カード

各診断カードでは、少なくとも次を表示する。

- セルアドレス
- `source` ピル
- `status` ピル
- `formulaText => outputValue`

この表示により、「cached をそのまま採用した数式」と「自前解決が必要だった数式」を見分けやすくする。

## 18. 内部メタ情報

現行実装では、Markdown 本文へそのまま出さない情報も内部メタとして保持する。

主なものは次の通りである。

### Workbook / Sheet レベル

- Sheet 順
- worksheet XML path
- sharedStrings
- definedNames

### Cell レベル

- `valueType`
- `rawValue`
- `outputValue`
- `styleIndex`
- `borders`
- `numFmtId`
- `formatCode`
- `formulaText`
- `formulaType`
- `spillRef`
- `resolutionStatus`
- `resolutionSource`
- `cachedValueState`

### Range / 構造レベル

- merge range
- 表候補 range
- 表候補 score
- 表候補の理由一覧

### Drawing レベル

- 画像の `mediaPath`
- グラフの `chartPath`
- 図形の `elementName`
- 図形の `anchorElementName`
- 図形の `rawEntries`

これらの内部メタ情報は、次の目的で用いられる。

- Markdown 出力値の決定
- 数式解決
- 診断表示
- デバッグや将来機能の足場

## 19. 既知の制約

現行実装には、少なくとも次の制約がある。

- Excel の見た目をピクセル単位で再現するものではない
- 表検出はヒューリスティックであり、値配置や罫線の条件によって過検出・過小検出が起こりうる
- レイアウト中心のシートでは、自然なセクション分割よりも通常表検出が優先される場合がある
- 数式解決は実務上重要なサブセットを対象としており、Excel 数式の完全互換ではない
- 画像、グラフ、図形は構造情報保持を主目的としており、意味理解や再描画を行わない
- グラフや図形は本文中へ厳密に埋め込まず、補助セクションとして末尾寄りに出力する
- 外部 Workbook 参照は解決しない
- `display` モードの表示形式適用は重点パターン中心であり、Excel の全書式を再現するわけではない

また、実装が内部メタを多く保持している一方で、そのすべてを UI や Markdown へ直接出しているわけではない。

## 20. 未対応・今後検討

現時点で未対応、または今後の検討対象として整理できるものは少なくとも次の通りである。

### 表検出・レイアウト

- レイアウト中心シートに対する専用分解ルール
- `カレンダー / ボード / ダッシュボード系` シートの別カテゴリ扱い
- 弱い表候補に対する警告や段階的運用
- 表候補スコアリングへの列型一貫性や空白関係の強化反映

### 数式

- Excel 数式の未対応サブセット拡張
- dynamic array / spill の完全対応
- lambda 系や future function への対応
- AST evaluator と legacy resolver の役割整理とさらなる AST 側への移行

### 出力・診断

- 数式診断の追加フィルタや対話的絞り込み
- `cachedValueState` の UI 表示要否
- グラフや図形のより自然な Markdown 表現
- 画像・図形・グラフの本文側への配置改善

### ドキュメント

- 上位仕様との継続的な整合
- 実データレビュー結果の impl-spec への反映ルール整理

## 21. 実装ファイル対応

現行実装における主要な責務とファイル対応は概ね次の通りである。

### 解析本体

- `docs/xlsx2md/src/xlsx2md/ts/core.ts`
  - ZIP 展開
  - workbook / worksheet 解析
  - sharedStrings / styles / definedNames
  - cell / merge / table / image / chart / shape モデル構築
  - 数式解決
  - 表検出
  - 地の文抽出
  - Markdown 組み立て
  - ZIP 出力

### 数式サブシステム

- `docs/xlsx2md/src/xlsx2md/ts/formula/tokenizer.ts`
  - 数式トークナイズ
- `docs/xlsx2md/src/xlsx2md/ts/formula/parser.ts`
  - AST 構築
- `docs/xlsx2md/src/xlsx2md/ts/formula/evaluator.ts`
  - AST 評価

### UI

- `docs/xlsx2md/src/xlsx2md/ts/main.ts`
  - 画面操作
  - オプション取得
  - 解析サマリー表示
  - 表候補スコア表示
  - 数式診断表示
  - ダウンロード / ZIP 保存

- `docs/xlsx2md/src/xlsx2md/css/app.css`
  - `xlsx2md` 画面固有の見た目

### テスト

- `docs/xlsx2md/tests/xlsx2md-main.test.js`
  - 実ファイルベース回帰テスト
  - Workbook 解析、Markdown 生成、画像・グラフ・図形・数式まわりの確認

- `docs/xlsx2md/tests/xlsx2md-formula-parser.test.js`
  - tokenizer / parser / evaluator の単体寄り確認

### 生成物

- `docs/xlsx2md/src/xlsx2md/js/*.js`
  - TypeScript からの生成物
- `docs/xlsx2md/xlsx2md.html`
  - single-file Web App の生成物

本書の記述基準は TypeScript 実装であり、生成物はその反映結果として扱う。

## 22. 実装参考コード

本章は、`impl-spec` だけで再実装可能性を上げるための補助資料である。

- 正本は `docs/xlsx2md/src/xlsx2md/ts/core.ts` とする
- ここでは、再実装時に骨格となる代表的な型定義と関数断片を掲載する
- 全量転載ではなく、章ごとの理解と再実装の足場になる範囲へ絞る

### 22.1 主要型定義

- 役割: Workbook / Sheet / Cell / 数式診断の最小モデルを固定する
- 入力: なし。再実装時の基底型として使う
- 出力: TypeScript の type 定義
- 前後関係: 後続のすべての関数断片は、これらの型を前提にしている

```ts
type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;
type FormulaResolutionSource = "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;
type CachedValueState = "present_nonempty" | "present_empty" | "absent" | null;

type ParsedCell = {
  address: string;
  row: number;
  col: number;
  valueType: string;
  rawValue: string;
  outputValue: string;
  formulaText: string;
  resolutionStatus: FormulaResolutionStatus;
  resolutionSource: FormulaResolutionSource;
  cachedValueState: CachedValueState;
  styleIndex: number;
  borders: BorderFlags;
  numFmtId: number;
  formatCode: string;
  formulaType: string;
  spillRef: string;
};

type ParsedSheet = {
  name: string;
  index: number;
  path: string;
  cells: ParsedCell[];
  merges: MergeRange[];
  tables: ParsedTable[];
  images: ParsedImageAsset[];
  charts: ParsedChartAsset[];
  shapes: ParsedShapeAsset[];
  maxRow: number;
  maxCol: number;
};

type ParsedWorkbook = {
  name: string;
  sheets: ParsedSheet[];
  sharedStrings: string[];
  definedNames: {
    name: string;
    formulaText: string;
    localSheetName: string | null;
  }[];
};
```

### 22.2 数式セル読込と cached 判定

`extractCellOutputValue(...)` は、数式セルの初期状態を決める入口である。特に `cachedValueState` と `resolutionStatus` / `resolutionSource` の初期値をここで確定する。

- 役割: 1 セル分の XML から `rawValue` / `outputValue` / 数式メタを決める
- 入力: `cellElement`, `sharedStrings`, `cellStyle`, `formulaOverride`
- 出力: `ParsedCell` 構築前の中間オブジェクト
- 前後関係: `parseWorksheet(...)` から呼ばれ、その後の再解決フェーズの初期状態になる

```ts
function extractCellOutputValue(
  cellElement: Element,
  sharedStrings: string[],
  cellStyle: CellStyleInfo,
  formulaOverride = ""
): {
  valueType: string;
  rawValue: string;
  outputValue: string;
  formulaText: string;
  resolutionStatus: FormulaResolutionStatus;
  resolutionSource: FormulaResolutionSource;
  cachedValueState: CachedValueState;
} {
  const type = (cellElement.getAttribute("t") || "").trim();
  const valueNode = cellElement.getElementsByTagName("v")[0] || null;
  const valueText = getTextContent(valueNode);
  const formulaText = formulaOverride || getTextContent(cellElement.getElementsByTagName("f")[0]);
  const cachedValueState: CachedValueState = !formulaText
    ? null
    : !valueNode
      ? "absent"
      : valueText === ""
        ? "present_empty"
        : "present_nonempty";
  if (formulaText) {
    const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
    if (/\[[^\]]+\.xlsx\]/i.test(normalizedFormula)) {
      return {
        valueType: type || "formula",
        rawValue: valueText || normalizedFormula,
        outputValue: normalizedFormula,
        formulaText: normalizedFormula,
        resolutionStatus: "unsupported_external",
        resolutionSource: "external_unsupported",
        cachedValueState
      };
    }
    if (valueNode) {
      const formattedValue = formatCellDisplayValue(valueText, cellStyle);
      return {
        valueType: type || "formula",
        rawValue: valueText,
        outputValue: formattedValue ?? valueText,
        formulaText: normalizedFormula,
        resolutionStatus: "resolved",
        resolutionSource: "cached_value",
        cachedValueState
      };
    }
    return {
      valueType: type || "formula",
      rawValue: normalizedFormula,
      outputValue: normalizedFormula,
      formulaText: normalizedFormula,
      resolutionStatus: "fallback_formula",
      resolutionSource: "formula_text",
      cachedValueState
    };
  }
  // ... 通常セル分岐
}
```

### 22.3 AST evaluator 呼び出し

`tryResolveFormulaExpressionWithAst(...)` は、`cached value` 不在時に最初に試される自前解決経路である。名前定義、構造化参照、spill 参照の解決コンテキストをここで与える。

- 役割: parser / evaluator を使って数式を AST ベースで評価する
- 入力: 数式文字列、現在シート名、セル参照 resolver、範囲 resolver、現在セルアドレス
- 出力: 解決できた場合の scalar 文字列、失敗時は `null`
- 前後関係: `tryResolveFormulaExpressionDetailed(...)` から、legacy resolver より先に呼ばれる

```ts
function tryResolveFormulaExpressionWithAst(
  expression: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string,
  resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
  currentAddress?: string
): string | null {
  const formulaApi = (globalThis as any).__xlsx2mdFormula;
  if (!formulaApi?.parseFormula || !formulaApi?.evaluateFormulaAst) {
    return null;
  }
  try {
    const ast = formulaApi.parseFormula(`=${expression}`);
    const evaluated = formulaApi.evaluateFormulaAst(ast, {
      resolveCell(ref: string, sheet: string | null) {
        return coerceFormulaAstScalar(resolveCellValue(sheet || currentSheetName, normalizeFormulaAddress(ref)));
      },
      resolveName(name: string) {
        const scopedRef = parseSheetScopedDefinedNameReference(name, currentSheetName);
        if (scopedRef) {
          const scopedValue = resolveDefinedNameScalarValue?.(scopedRef.sheetName, scopedRef.name) ?? null;
          if (scopedValue != null) {
            return coerceFormulaAstScalar(scopedValue);
          }
        }
        const scalarValue = resolveDefinedNameScalarValue?.(currentSheetName, name) ?? null;
        if (scalarValue != null) {
          return coerceFormulaAstScalar(scalarValue);
        }
        const rangeRef = resolveDefinedNameRangeRef?.(currentSheetName, name) ?? null;
        if (rangeRef && resolveRangeEntries) {
          return createFormulaAstRangeMatrix(
            rangeRef.sheetName,
            rangeRef.start,
            rangeRef.end,
            resolveRangeEntries
          );
        }
        return null;
      },
      resolveStructuredRef(table: string, column: string) {
        const rangeRef = resolveStructuredRangeRef?.(currentSheetName, `${table}[${column}]`) ?? null;
        if (!rangeRef || !resolveRangeEntries) {
          return null;
        }
        return createFormulaAstRangeMatrix(
          rangeRef.sheetName,
          rangeRef.start,
          rangeRef.end,
          resolveRangeEntries
        );
      }
      // ... resolveRange / resolveSpill / currentAddress など
    });
    return formulaApi.stringifyFormulaValue?.(evaluated) ?? String(evaluated ?? "");
  } catch {
    return null;
  }
}
```

### 22.4 地の文抽出

`extractNarrativeBlocks(...)` は、表として採用されなかったセルを行単位にまとめ、近接する行を narrative block として束ねる。

- 役割: 表外セルを paragraph / list 判定前の block へまとめる
- 入力: `sheet`, 採用済み `tables`, `options`
- 出力: `NarrativeBlock[]`
- 前後関係: `detectTableCandidates(...)` の結果を前提に動き、後段で `renderNarrativeBlock(...)` に渡される

```ts
function extractNarrativeBlocks(sheet: ParsedSheet, tables: TableCandidate[], options: MarkdownOptions = {}): NarrativeBlock[] {
  const rowMap = new Map<number, ParsedCell[]>();
  for (const cell of sheet.cells) {
    if (!cell.outputValue) continue;
    if (isCellInAnyTable(cell.row, cell.col, tables)) continue;
    const entries = rowMap.get(cell.row) || [];
    entries.push(cell);
    rowMap.set(cell.row, entries);
  }
  const rowNumbers = Array.from(rowMap.keys()).sort((a, b) => a - b);
  const blocks: NarrativeBlock[] = [];
  let current: NarrativeBlock | null = null;
  let previousRow = -100;

  for (const rowNumber of rowNumbers) {
    const cells = (rowMap.get(rowNumber) || []).slice().sort((a, b) => a.col - b.col);
    const rowSegments = splitNarrativeRowSegments(cells, options);
    for (const segment of rowSegments) {
      const rowText = segment.values.join(" ").trim();
      if (!rowText) continue;
      const startCol = segment.startCol;
      if (!current || rowNumber - previousRow > 1 || Math.abs(startCol - current.startCol) > 3) {
        current = {
          startRow: rowNumber,
          startCol,
          endRow: rowNumber,
          lines: [rowText],
          items: [{
            row: rowNumber,
            startCol,
            text: rowText,
            cellValues: segment.values
          }]
        };
        blocks.push(current);
      } else {
        current.lines.push(rowText);
        current.endRow = rowNumber;
        current.items.push({
          row: rowNumber,
          startCol,
          text: rowText,
          cellValues: segment.values
        });
      }
      previousRow = rowNumber;
    }
  }

  return blocks;
}
```

### 22.5 表候補検出

`detectTableCandidates(...)` は、seed cell の 4 近傍連結成分を外接矩形化し、ヒューリスティックなスコアで採用可否を決める。

- 役割: worksheet 全体から Markdown テーブル候補を抽出する
- 入力: `sheet`
- 出力: `TableCandidate[]`
- 前後関係: `convertSheetToMarkdown(...)` で最初に呼ばれ、地の文抽出と表描画の両方に影響する

```ts
function detectTableCandidates(sheet: ParsedSheet): TableCandidate[] {
  const seedCells = collectTableSeedCells(sheet);
  const positionMap = new Map<string, ParsedCell>();
  for (const cell of seedCells) {
    positionMap.set(`${cell.row}:${cell.col}`, cell);
  }
  const visited = new Set<string>();
  const candidates: TableCandidate[] = [];

  for (const cell of seedCells) {
    const key = `${cell.row}:${cell.col}`;
    if (visited.has(key)) continue;
    const queue = [cell];
    const component: ParsedCell[] = [];
    visited.add(key);
    while (queue.length > 0) {
      const current = queue.shift() as ParsedCell;
      component.push(current);
      for (const [rowDelta, colDelta] of [[1, 0], [-1, 0], [0, 1], [0, -1]]) {
        const nextKey = `${current.row + rowDelta}:${current.col + colDelta}`;
        const nextCell = positionMap.get(nextKey);
        if (!nextCell || visited.has(nextKey)) continue;
        visited.add(nextKey);
        queue.push(nextCell);
      }
    }

    const rows = component.map((entry) => entry.row);
    const cols = component.map((entry) => entry.col);
    const startRow = Math.min(...rows);
    const endRow = Math.max(...rows);
    const startCol = Math.min(...cols);
    const endCol = Math.max(...cols);
    const area = Math.max(1, (endRow - startRow + 1) * (endCol - startCol + 1));
    const density = component.filter((entry) => entry.outputValue.trim()).length / area;
    const rowCount = endRow - startRow + 1;
    const colCount = endCol - startCol + 1;
    if (rowCount < 2 || colCount < 2) {
      continue;
    }

    let score = 0;
    const reasons: string[] = [];
    const borderCells = component.filter((entry) => entry.borders.top || entry.borders.bottom || entry.borders.left || entry.borders.right);
    if (rowCount >= 2 && colCount >= 2) {
      score += TABLE_SCORE_WEIGHTS.minGrid;
      reasons.push(`2x2 以上 (+${TABLE_SCORE_WEIGHTS.minGrid})`);
    }
    if (borderCells.length >= Math.max(2, Math.ceil(component.length * 0.3))) {
      score += TABLE_SCORE_WEIGHTS.borderPresence;
      reasons.push(`罫線あり (+${TABLE_SCORE_WEIGHTS.borderPresence})`);
    }
    if (density >= 0.55) {
      score += TABLE_SCORE_WEIGHTS.densityHigh;
      reasons.push(`密度高 (+${TABLE_SCORE_WEIGHTS.densityHigh})`);
    }
    // ... ヘッダらしさ、結合セル過多、長文中心など

    if (score >= TABLE_SCORE_WEIGHTS.threshold) {
      candidates.push({
        startRow,
        startCol,
        endRow,
        endCol,
        score,
        reasonSummary: reasons
      });
    }
  }

  return candidates.sort((left, right) => {
    if (left.startRow !== right.startRow) return left.startRow - right.startRow;
    return left.startCol - right.startCol;
  });
}
```

### 22.6 Markdown 組み立てと ZIP 出力

`convertSheetToMarkdown(...)` は、表候補、地の文、補助セクションを統合して 1 Sheet 分の Markdown とサマリを返す。`createCombinedMarkdownExportFile(...)` と `createExportEntries(...)` は Workbook 単位の保存物を作る。

- 役割: sheet 単位 Markdown、連結 Markdown、ZIP entry を組み立てる
- 入力: `workbook`, `sheet`, `options` または `markdownFiles`
- 出力: `MarkdownFile`、連結 Markdown、ZIP entry 一覧
- 前後関係: 解析本体の最終段であり、UI のプレビューと保存の両方がこの結果を使う

```ts
function convertSheetToMarkdown(workbook: ParsedWorkbook, sheet: ParsedSheet, options: MarkdownOptions = {}): MarkdownFile {
  const charts = sheet.charts || [];
  const shapes = sheet.shapes || [];
  const shapeBlocks = extractShapeBlocks(shapes);
  const treatFirstRowAsHeader = options.treatFirstRowAsHeader !== false;
  const tables = detectTableCandidates(sheet);
  const narrativeBlocks = extractNarrativeBlocks(sheet, tables, options);
  const sectionBlocks = extractSectionBlocks(sheet, tables, narrativeBlocks);
  const formulaDiagnostics = sheet.cells
    .filter((cell) => !!cell.formulaText && cell.resolutionStatus !== null)
    .map((cell) => ({
      address: cell.address,
      formulaText: cell.formulaText,
      status: cell.resolutionStatus,
      source: cell.resolutionSource,
      outputValue: cell.outputValue
    }));

  // narrative / table を位置順に sections へ積む
  // groupedSections 単位で本文を作り、必要なら --- で区切る
  // chart / shape / image セクションを末尾へ追加する

  return {
    fileName: createOutputFileName(workbook.name, sheet.index, sheet.name, options.outputMode || "display"),
    sheetName: sheet.name,
    markdown,
    summary: {
      outputMode: options.outputMode || "display",
      sections: sectionBlocks.length,
      tables: tables.length,
      narrativeBlocks: narrativeBlocks.length,
      merges: sheet.merges.length,
      images: sheet.images.length,
      charts: charts.length,
      cells: sheet.cells.length,
      tableScores: tables.map((table) => ({
        range: formatRange(table.startRow, table.startCol, table.endRow, table.endCol),
        score: table.score,
        reasons: [...table.reasonSummary]
      })),
      formulaDiagnostics
    }
  };
}

function createCombinedMarkdownExportFile(workbook: ParsedWorkbook, markdownFiles: MarkdownFile[]): { fileName: string; content: string } {
  const outputMode = markdownFiles[0]?.summary.outputMode || "display";
  const suffix = outputMode === "display" ? "" : `_${outputMode}`;
  const fileName = `${String(workbook.name || "workbook").replace(/\.xlsx$/i, "")}${suffix}.md`;
  const content = markdownFiles
    .map((markdownFile) => `<!-- ${markdownFile.fileName.replace(/\.md$/i, "")} -->\n${markdownFile.markdown}`)
    .join("\n\n");
  return { fileName, content };
}

function createExportEntries(workbook: ParsedWorkbook, markdownFiles: MarkdownFile[]): ExportEntry[] {
  const entries: ExportEntry[] = [];
  if (markdownFiles.length > 0) {
    const combined = createCombinedMarkdownExportFile(workbook, markdownFiles);
    entries.push({
      name: `output/${combined.fileName}`,
      data: textEncoder.encode(`${combined.content}\n`)
    });
  }
  for (const sheet of workbook.sheets) {
    for (const image of sheet.images) {
      entries.push({
        name: `output/${image.path}`,
        data: image.data
      });
    }
    for (const shape of sheet.shapes || []) {
      if (!shape.svgPath || !shape.svgData) continue;
      entries.push({
        name: `output/${shape.svgPath}`,
        data: shape.svgData
      });
    }
  }
  return entries;
}
```

### 22.7 今後の付録追加候補

再実装可能性をさらに上げるには、次を順に本章へ追加するとよい。

- `parseWorkbook(...)` の worksheet / rels 解決断片
- shared formula 展開コード
- `renderNarrativeBlock(...)` とリスト化判定コード
- image / chart / shape 抽出コード
- `main.ts` 側の formula diagnostics 集計コード

### 22.8 Workbook 入口と shared formula 展開

Workbook 読み込みの入口は `parseWorkbook(...)`、sheet 単位の数式読み込みと shared formula 展開の入口は `parseWorksheet(...)` である。数式セルの `formulaOverride` はここで決まる。

- 役割: ZIP 展開後に workbook 全体を内部モデルへ変換し、sheet ごとの初期解析を行う
- 入力: `arrayBuffer`, `workbookName`、および worksheet XML / sharedStrings / styles
- 出力: `ParsedWorkbook` と、その構成要素としての `ParsedSheet[]`
- 前後関係: `extractCellOutputValue(...)`、drawing 抽出、再解決フェーズのすべての入口になる

```ts
async function parseWorkbook(arrayBuffer: ArrayBuffer, workbookName = "workbook.xlsx"): Promise<ParsedWorkbook> {
  const files = await unzipEntries(arrayBuffer);
  const workbookBytes = files.get("xl/workbook.xml");
  if (!workbookBytes) {
    throw new Error("xl/workbook.xml が見つかりません");
  }
  const sharedStrings = parseSharedStrings(files);
  const cellStyles = parseCellStyles(files);
  const rels = parseRelationships(files, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");
  const workbookDoc = xmlToDocument(decodeXmlText(workbookBytes));
  const sheetNodes = Array.from(workbookDoc.getElementsByTagName("sheet"));
  const sheetNames = sheetNodes.map((sheetNode, index) => sheetNode.getAttribute("name") || `Sheet${index + 1}`);
  const definedNames = parseDefinedNames(workbookDoc, sheetNames);
  const sheets = sheetNodes.map((sheetNode, index) => {
    const name = sheetNode.getAttribute("name") || `Sheet${index + 1}`;
    const relId = sheetNode.getAttribute("r:id") || "";
    const sheetPath = rels.get(relId) || "";
    return parseWorksheet(files, name, sheetPath, index + 1, sharedStrings, cellStyles);
  });
  const workbook = {
    name: workbookName,
    sheets,
    sharedStrings,
    definedNames
  };
  resolveSimpleFormulaReferences(workbook);
  resolveSimpleFormulaReferences(workbook);
  resolveSimpleFormulaReferences(workbook);
  return workbook;
}

function parseWorksheet(
  files: Map<string, Uint8Array>,
  sheetName: string,
  sheetPath: string,
  sheetIndex: number,
  sharedStrings: string[],
  cellStyles: CellStyleInfo[]
): ParsedSheet {
  const bytes = files.get(sheetPath);
  if (!bytes) {
    throw new Error(`シート XML が見つかりません: ${sheetPath}`);
  }
  const doc = xmlToDocument(decodeXmlText(bytes));
  const sharedFormulaMap = new Map<string, { address: string; formulaText: string }>();
  const cells = Array.from(doc.getElementsByTagName("c")).map((cellElement) => {
    const address = cellElement.getAttribute("r") || "";
    const position = parseCellAddress(address);
    const styleIndex = Number(cellElement.getAttribute("s") || 0);
    const cellStyle = cellStyles[styleIndex] || {
      borders: EMPTY_BORDERS,
      numFmtId: 0,
      formatCode: "General"
    };
    let formulaOverride = "";
    const formulaElement = cellElement.getElementsByTagName("f")[0] || null;
    const formulaType = formulaElement?.getAttribute("t") || "";
    const spillRef = formulaElement?.getAttribute("ref") || "";
    const sharedIndex = formulaElement?.getAttribute("si") || "";
    const formulaText = getTextContent(formulaElement);
    if (formulaType === "shared" && sharedIndex) {
      if (formulaText) {
        const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
        sharedFormulaMap.set(sharedIndex, { address, formulaText: normalizedFormula });
        formulaOverride = normalizedFormula;
      } else {
        const sharedBase = sharedFormulaMap.get(sharedIndex);
        if (sharedBase) {
          formulaOverride = translateSharedFormula(sharedBase.formulaText, sharedBase.address, address);
        }
      }
    }
    const output = extractCellOutputValue(cellElement, sharedStrings, cellStyle, formulaOverride);
    return {
      address,
      row: position.row,
      col: position.col,
      valueType: output.valueType,
      rawValue: output.rawValue,
      outputValue: output.outputValue,
      formulaText: output.formulaText,
      resolutionStatus: output.resolutionStatus,
      resolutionSource: output.resolutionSource,
      cachedValueState: output.cachedValueState,
      styleIndex,
      borders: cellStyle.borders,
      numFmtId: cellStyle.numFmtId,
      formatCode: cellStyle.formatCode,
      formulaType,
      spillRef
    } satisfies ParsedCell;
  });
  // merges / tables / images / charts / shapes を続けて構築する
}
```

shared formula の平行移動自体は `translateSharedFormula(...)` が担う。

- 役割: shared formula の基底式を対象セル位置へ平行移動する
- 入力: 基底式、基底セルアドレス、対象セルアドレス
- 出力: 対象セル用に平行移動した数式文字列
- 前後関係: `parseWorksheet(...)` 内で `formulaOverride` を作るために使う

```ts
function translateSharedFormula(baseFormulaText: string, baseAddress: string, targetAddress: string): string {
  const basePos = parseCellAddress(baseAddress);
  const targetPos = parseCellAddress(targetAddress);
  if (!basePos.row || !basePos.col || !targetPos.row || !targetPos.col) {
    return baseFormulaText;
  }
  const rowOffset = targetPos.row - basePos.row;
  const colOffset = targetPos.col - basePos.col;
  const normalized = String(baseFormulaText || "").replace(/^=/, "");
  const translated = normalized.replace(
    /(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)/g,
    (full, quotedSheet, plainSheet, qualifiedAddress, localAddress) => {
      const address = qualifiedAddress || localAddress;
      if (!address) return full;
      const shifted = shiftReferenceAddress(address, rowOffset, colOffset);
      if (qualifiedAddress) {
        const sheetPrefix = quotedSheet ? `'${quotedSheet}'` : plainSheet;
        return `${sheetPrefix}!${shifted}`;
      }
      return shifted;
    }
  );
  return translated.startsWith("=") ? translated : `=${translated}`;
}
```

### 22.9 リスト化と section block

地の文は block 化だけではなく、`renderNarrativeBlock(...)` で「4 件以上連続した list candidate を箇条書きへ変換する」という後段ロジックを持つ。

- 役割: narrative block を paragraph または Markdown list へ最終整形する
- 入力: `NarrativeBlock`
- 出力: Markdown 文字列
- 前後関係: `extractNarrativeBlocks(...)` の直後ではなく、`convertSheetToMarkdown(...)` の本文組み立て時に使う

```ts
function renderNarrativeBlock(block: NarrativeBlock): string {
  if (!block.items || block.items.length === 0) {
    return block.lines.join("\n");
  }
  const parts: string[] = [];
  let index = 0;
  while (index < block.items.length) {
    let runEnd = index;
    while (
      runEnd < block.items.length
      && isNarrativeListCandidate(block.items[runEnd])
      && (runEnd === index || block.items[runEnd].row === block.items[runEnd - 1].row + 1)
    ) {
      runEnd += 1;
    }
    const runLength = runEnd - index;
    if (runLength >= 4) {
      parts.push(block.items.slice(index, runEnd).map((item) => formatNarrativeListItem(item)).join("\n"));
      index = runEnd;
      continue;
    }
    let proseEnd = index;
    while (proseEnd < block.items.length) {
      const nextRunStart = proseEnd;
      let candidateEnd = nextRunStart;
      while (
        candidateEnd < block.items.length
        && isNarrativeListCandidate(block.items[candidateEnd])
        && (candidateEnd === nextRunStart || block.items[candidateEnd].row === block.items[candidateEnd - 1].row + 1)
      ) {
        candidateEnd += 1;
      }
      if (candidateEnd - nextRunStart >= 4) {
        break;
      }
      proseEnd += 1;
    }
    parts.push(block.items.slice(index, proseEnd).map((item) => item.text).join("\n"));
    index = proseEnd;
  }
  return parts.join("\n\n");
}
```

現在の `section block` は専用の意味分解器ではなく、narrative / table / image / chart のアンカー位置を縦ギャップで束ねる軽量 grouping である。

- 役割: 本文と補助要素のアンカーを、縦方向の大きな空白で section 単位にまとめる
- 入力: `sheet`, `tables`, `narrativeBlocks`
- 出力: `SectionBlock[]`
- 前後関係: `convertSheetToMarkdown(...)` の `groupedSections` 構築で使われ、`---` 区切りの有無に影響する

```ts
function extractSectionBlocks(sheet: ParsedSheet, tables: TableCandidate[], narrativeBlocks: NarrativeBlock[]): SectionBlock[] {
  const charts = sheet.charts || [];
  const anchors: Array<{ startRow: number; startCol: number; endRow: number; endCol: number }> = [];

  for (const block of narrativeBlocks) {
    anchors.push({
      startRow: block.startRow,
      startCol: block.startCol,
      endRow: block.endRow,
      endCol: Math.max(block.startCol, ...block.items.map((item) => item.startCol))
    });
  }
  for (const table of tables) {
    anchors.push({
      startRow: table.startRow,
      startCol: table.startCol,
      endRow: table.endRow,
      endCol: table.endCol
    });
  }
  for (const image of sheet.images) {
    const anchor = parseCellAddress(image.anchor);
    if (anchor.row > 0 && anchor.col > 0) {
      anchors.push({ startRow: anchor.row, startCol: anchor.col, endRow: anchor.row, endCol: anchor.col });
    }
  }
  for (const chart of charts) {
    const anchor = parseCellAddress(chart.anchor);
    if (anchor.row > 0 && anchor.col > 0) {
      anchors.push({ startRow: anchor.row, startCol: anchor.col, endRow: anchor.row, endCol: anchor.col });
    }
  }
  if (anchors.length === 0) {
    return [];
  }

  anchors.sort((left, right) => {
    if (left.startRow !== right.startRow) return left.startRow - right.startRow;
    return left.startCol - right.startCol;
  });

  const sections: SectionBlock[] = [];
  let current: SectionBlock | null = null;
  let previousEndRow = -100;
  const verticalGapThreshold = 4;

  for (const anchor of anchors) {
    const gap = anchor.startRow - previousEndRow;
    if (!current || gap > verticalGapThreshold) {
      current = {
        startRow: anchor.startRow,
        startCol: anchor.startCol,
        endRow: anchor.endRow,
        endCol: anchor.endCol
      };
      sections.push(current);
    } else {
      current.startRow = Math.min(current.startRow, anchor.startRow);
      current.startCol = Math.min(current.startCol, anchor.startCol);
      current.endRow = Math.max(current.endRow, anchor.endRow);
      current.endCol = Math.max(current.endCol, anchor.endCol);
    }
    previousEndRow = Math.max(previousEndRow, anchor.endRow);
  }

  return sections;
}
```

### 22.10 drawing 抽出

画像、グラフ、図形はすべて worksheet rels から drawing をたどる。画像とグラフは drawing rels を引き、図形は anchor 内の `sp` / `cxnSp` を直接読む。

- 役割: worksheet にぶら下がる drawing から image / chart / shape を抽出する
- 入力: `files`, `sheetName`, `sheetPath`
- 出力: `ParsedImageAsset[]`、`ParsedChartAsset[]`、`ParsedShapeAsset[]`
- 前後関係: `parseWorksheet(...)` から呼ばれ、後段の Markdown 補助セクションと ZIP 出力に直結する

```ts
function parseDrawingImages(
  files: Map<string, Uint8Array>,
  sheetName: string,
  sheetPath: string
): ParsedImageAsset[] {
  const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
  const imageAssets: ParsedImageAsset[] = [];
  let imageCounter = 1;

  for (const drawingPath of sheetRels.values()) {
    if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
    const drawingBytes = files.get(drawingPath);
    if (!drawingBytes) continue;
    const drawingDoc = xmlToDocument(decodeXmlText(drawingBytes));
    const drawingRels = parseRelationships(files, buildRelsPath(drawingPath), drawingPath);
    const anchors = getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(getElementsByLocalName(drawingDoc, "twoCellAnchor"));

    for (const anchor of anchors) {
      const from = getFirstChildByLocalName(anchor, "from");
      const colNode = getFirstChildByLocalName(from || anchor, "col");
      const rowNode = getFirstChildByLocalName(from || anchor, "row");
      const col = Number(getTextContent(colNode)) + 1;
      const row = Number(getTextContent(rowNode)) + 1;
      if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
        continue;
      }
      const blip = getElementsByLocalName(anchor, "blip")[0] || null;
      const embedId = blip?.getAttribute("r:embed") || blip?.getAttribute("embed") || "";
      const mediaPath = drawingRels.get(embedId) || "";
      if (!mediaPath) continue;
      const mediaBytes = files.get(mediaPath);
      if (!mediaBytes) continue;

      const extension = getImageExtension(mediaPath);
      const safeDir = createSafeSheetAssetDir(sheetName);
      const filename = `image_${String(imageCounter).padStart(3, "0")}.${extension}`;
      imageAssets.push({
        sheetName,
        filename,
        path: `assets/${safeDir}/${filename}`,
        anchor: `${colToLetters(col)}${row}`,
        data: new Uint8Array(mediaBytes),
        mediaPath
      });
      imageCounter += 1;
    }
  }

  return imageAssets;
}
```

`parseDrawingImages(...)` は、anchor から `blip` の `r:embed` を取り、drawing rels を通して `xl/media/*` へ到達する経路である。保存名と `assets/<sheet>/image_XXX.ext` の相対パスもこの関数で決める。

```ts
function parseDrawingCharts(
  files: Map<string, Uint8Array>,
  sheetName: string,
  sheetPath: string
): ParsedChartAsset[] {
  const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
  const charts: ParsedChartAsset[] = [];

  for (const drawingPath of sheetRels.values()) {
    if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
    const drawingBytes = files.get(drawingPath);
    if (!drawingBytes) continue;
    const drawingDoc = xmlToDocument(decodeXmlText(drawingBytes));
    const drawingRels = parseRelationships(files, buildRelsPath(drawingPath), drawingPath);
    const anchors = getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(getElementsByLocalName(drawingDoc, "twoCellAnchor"));

    for (const anchor of anchors) {
      const from = getFirstChildByLocalName(anchor, "from");
      const colNode = getFirstChildByLocalName(from || anchor, "col");
      const rowNode = getFirstChildByLocalName(from || anchor, "row");
      const col = Number(getTextContent(colNode)) + 1;
      const row = Number(getTextContent(rowNode)) + 1;
      if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
        continue;
      }

      const chartNode = getFirstChildByLocalName(anchor, "graphicFrame");
      const chartRef = getElementsByLocalName(chartNode || anchor, "chart")[0] || null;
      const relId = chartRef?.getAttribute("r:id") || chartRef?.getAttribute("id") || "";
      const chartPath = drawingRels.get(relId) || "";
      if (!chartPath) continue;
      const chartBytes = files.get(chartPath);
      if (!chartBytes) continue;
      const chartDoc = xmlToDocument(decodeXmlText(chartBytes));

      charts.push({
        sheetName,
        anchor: `${colToLetters(col)}${row}`,
        chartPath,
        title: parseChartTitle(chartDoc),
        chartType: parseChartType(chartDoc),
        series: parseChartSeries(chartDoc)
      });
    }
  }

  return charts;
}
```

`parseDrawingCharts(...)` は、anchor から `graphicFrame` 内の `chart` 参照を取り、drawing rels を通して chart XML を読む。画像抽出と似た枠組みだが、最終的に返すのはバイナリではなく `title / chartType / series` の意味情報である。

```ts
function parseDrawingShapes(
  files: Map<string, Uint8Array>,
  sheetName: string,
  sheetPath: string
): ParsedShapeAsset[] {
  const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
  const shapes: ParsedShapeAsset[] = [];
  let shapeCounter = 1;

  for (const drawingPath of sheetRels.values()) {
    if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
    const drawingBytes = files.get(drawingPath);
    if (!drawingBytes) continue;
    const drawingDoc = xmlToDocument(decodeXmlText(drawingBytes));
    const anchors = getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(getElementsByLocalName(drawingDoc, "twoCellAnchor"));

    for (const anchor of anchors) {
      const from = getFirstChildByLocalName(anchor, "from");
      const colNode = getFirstChildByLocalName(from || anchor, "col");
      const rowNode = getFirstChildByLocalName(from || anchor, "row");
      const col = Number(getTextContent(colNode)) + 1;
      const row = Number(getTextContent(rowNode)) + 1;
      if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
        continue;
      }

      if (getElementsByLocalName(anchor, "blip").length > 0) continue;
      if (getElementsByLocalName(anchor, "chart").length > 0) continue;

      const shapeNode = getFirstChildByLocalName(anchor, "sp") || getFirstChildByLocalName(anchor, "cxnSp");
      if (!shapeNode) continue;
      const cNvPr = getFirstChildByLocalName(getFirstChildByLocalName(shapeNode, shapeNode.localName === "sp" ? "nvSpPr" : "nvCxnSpPr") || shapeNode, "cNvPr");
      const { widthEmu, heightEmu } = parseShapeExt(anchor, shapeNode);
      const svgAsset = drawingHelper?.renderShapeSvg?.(shapeNode, anchor, sheetName, shapeCounter) || null;
      shapes.push({
        sheetName,
        anchor: `${colToLetters(col)}${row}`,
        name: String(cNvPr?.getAttribute("name") || "").trim() || "図形",
        kind: parseShapeKind(shapeNode),
        text: parseShapeText(shapeNode),
        widthEmu,
        heightEmu,
        elementName: `xdr:${shapeNode.localName}`,
        anchorElementName: anchor.tagName || anchor.nodeName || anchor.localName || "anchor",
        rawEntries: parseShapeRawEntries(anchor),
        bbox: parseShapeBoundingBox(anchor, shapeNode, widthEmu, heightEmu),
        svgFilename: svgAsset?.filename || null,
        svgPath: svgAsset?.path || null,
        svgData: svgAsset?.data || null
      });
      shapeCounter += 1;
    }
  }

  return shapes;
}
```

`parseDrawingShapes(...)` は、画像とグラフを除いた anchor を対象に shape を拾う。`rawEntries`、bounding box、必要に応じた `svgAsset` をまとめて保持するため、3 つの drawing 抽出関数の中では最も内部メタが多い。

### 22.11 再実装可能性の残課題

本章までで、再実装時の骨格はかなり追いやすくなった。なお、同等実装をさらに安定して再現するには、次がまだあると強い。

- `resolveSimpleFormulaReferences(...)` の全体断片
- `formatCellDisplayValue(...)` の主要分岐
- `parseRelationships(...)` と path 正規化
- `parseDefinedNames(...)` と structured reference 解決
- `parseShapeRawEntries(...)` と SVG 化 helper の境界

### 22.12 表示形式の主要分岐

`formatCellDisplayValue(...)` は `display` モードの基礎であり、日付・ゼロ値・パーセント・指数・分数・通貨をここで分岐する。

- 役割: `rawValue` と `CellStyleInfo` から表示値寄り文字列を作る
- 入力: `rawValue`, `cellStyle`
- 出力: 表示形式適用後の文字列、または `null`
- 前後関係: 通常セル読込と `applyResolvedFormulaValue(...)` の両方から使われる

```ts
function formatCellDisplayValue(rawValue: string, cellStyle: CellStyleInfo): string | null {
  const numericValue = Number(rawValue);
  const formatCode = String(cellStyle.formatCode || "");
  if (!formatCode || formatCode === "General") {
    return null;
  }
  const normalized = formatCode.toLowerCase();
  const formatSections = splitFormatSections(formatCode);

  if (!Number.isNaN(numericValue) && isDateFormatCode(formatCode)) {
    const parts = excelSerialToDateParts(numericValue);
    if (!parts) return null;
    const directFormatted = formatDateByPattern(parts, formatCode);
    if (directFormatted !== null) {
      return directFormatted;
    }
    const hasDate = /y/.test(normalized)
      || /d/.test(normalized)
      || /(^|[^a-z])m(?:\/|-)/.test(normalized)
      || /(?:\/|-)m(?:[^a-z]|$)/.test(normalized);
    const hasTime = /h/.test(normalized) || /s/.test(normalized) || normalized.includes(":") || normalized.includes("am/pm");
    if (hasDate && hasTime) {
      return `${parts.yyyy}-${parts.mm}-${parts.dd} ${parts.hh}:${parts.mi}:${parts.ss}`;
    }
    if (hasTime && !hasDate) {
      return `${parts.hh}:${parts.mi}:${parts.ss}`;
    }
    return `${parts.yyyy}-${parts.mm}-${parts.dd}`;
  }

  if (Number.isNaN(numericValue)) {
    return null;
  }

  if (numericValue === 0 && formatSections[2]) {
    const zeroText = formatZeroSection(formatSections[2]);
    if (zeroText) {
      return zeroText;
    }
  }

  if (normalized.includes("%")) {
    const percentPattern = normalized.split(";")[0] || normalized;
    const decimalPlaces = (percentPattern.split(".")[1] || "").replace(/[^0#]/g, "").length;
    return `${(numericValue * 100).toFixed(decimalPlaces)}%`;
  }

  if (cellStyle.numFmtId === 186 || /dbnum3/i.test(formatCode)) {
    return formatDbNum3Pattern(rawValue);
  }

  if (cellStyle.numFmtId === 42) {
    return `¥ ${formatNumberByPattern(numericValue, "#,##0").replace(/^-/, "")}`;
  }

  if (/[#0][^;]*e\+0+/i.test(formatCode)) {
    // ... scientific notation 分岐
  }
  if (normalized.includes("?/?")) {
    return formatFractionPattern(numericValue);
  }
  if (/^[^;]*[#0,]+(?:\.[#0]+)?/.test(formatCode)) {
    // ... 数値 / 通貨パターン
  }
  return null;
}
```

### 22.13 rels 解決と名前定義

Workbook / worksheet / drawing 間の辿りは `parseRelationships(...)` が共通化している。名前定義は `parseDefinedNames(...)` で workbook 読み込み時に一括収集する。

- 役割: XML パーツ間の参照解決と workbook scope / sheet scope name の初期収集を行う
- 入力: `files`, `relsPath`, `sourcePath` または `workbookDoc`, `sheetNames`
- 出力: `Map<string, string>`、defined name 一覧
- 前後関係: `parseWorkbook(...)`、`parseWorksheetTables(...)`、drawing 抽出、formula resolver の基礎データになる

```ts
function parseRelationships(files: Map<string, Uint8Array>, relsPath: string, sourcePath: string): Map<string, string> {
  const relBytes = files.get(relsPath);
  const relations = new Map<string, string>();
  if (!relBytes) {
    return relations;
  }
  const doc = xmlToDocument(decodeXmlText(relBytes));
  const nodes = Array.from(doc.getElementsByTagName("Relationship"));
  for (const node of nodes) {
    const id = node.getAttribute("Id") || "";
    const target = node.getAttribute("Target") || "";
    if (!id || !target) continue;
    relations.set(id, normalizeZipPath(sourcePath, target));
  }
  return relations;
}

function buildRelsPath(sourcePath: string): string {
  const parts = sourcePath.split("/");
  const fileName = parts.pop() || "";
  const dir = parts.join("/");
  return `${dir}/_rels/${fileName}.rels`;
}
```

```ts
function parseDefinedNames(workbookDoc: Document, sheetNames: string[]): {
  name: string;
  formulaText: string;
  localSheetName: string | null;
}[] {
  const result: {
    name: string;
    formulaText: string;
    localSheetName: string | null;
  }[] = [];
  const definedNameElements = Array.from(workbookDoc.getElementsByTagName("definedName"));
  for (const element of definedNameElements) {
    const name = element.getAttribute("name") || "";
    if (!name || name.startsWith("_xlnm.")) continue;
    const formulaText = getTextContent(element).trim();
    if (!formulaText) continue;
    const localSheetIdText = element.getAttribute("localSheetId");
    const localSheetId = localSheetIdText == null || localSheetIdText === "" ? Number.NaN : Number(localSheetIdText);
    result.push({
      name,
      formulaText: formulaText.startsWith("=") ? formulaText : `=${formulaText}`,
      localSheetName: Number.isNaN(localSheetId) ? null : (sheetNames[localSheetId] || null)
    });
  }
  return result;
}
```

### 22.14 数式再解決の入口

Workbook 全体の再解決は `resolveSimpleFormulaReferences(...)` が担う。実際の関数本体は長いが、入口の役割は次の通りである。

- 役割: workbook 全体の数式セルに対して再解決を反復適用する
- 入力: `ParsedWorkbook`
- 出力: 返り値は `void`。各 `ParsedCell` の `rawValue` / `outputValue` / `resolution*` を更新する
- 前後関係: `parseWorkbook(...)` の最後に複数回呼ばれ、依存解決を収束させる

```ts
function resolveSimpleFormulaReferences(workbook: ParsedWorkbook): void {
  const resolver = buildFormulaResolver(workbook);
  for (const sheet of workbook.sheets) {
    for (const cell of sheet.cells) {
      if (!cell.formulaText) continue;
      if (cell.resolutionStatus === "resolved") continue;
      if (cell.resolutionStatus === "unsupported_external") continue;

      // 1. AST evaluator を試す
      // 2. 未解決なら既存 resolver 群を試す
      // 3. 解けたら applyResolvedFormulaValue(...)
      // 4. 解けなければ formula_text / fallback_formula を維持する
    }
  }
}
```

再実装時は、この入口と `buildFormulaResolver(...)`、`tryResolveFormulaExpressionWithAst(...)`、個別 resolver 群をセットで写す必要がある。

### 22.15 図形 raw flattening

図形セクションの追跡可能性は `parseShapeRawEntries(...)` に依存する。これは anchor 以下の XML を `key / value` の平坦配列へ変換する処理である。

- 役割: shape / anchor XML を Markdown に出しやすい `key / value` 列へ平坦化する
- 入力: drawing anchor 要素
- 出力: `{ key, value }[]`
- 前後関係: `parseDrawingShapes(...)` が保持し、後段で階層 Markdown へ再構成される

```ts
function parseShapeRawEntries(anchor: Element): { key: string; value: string }[] {
  const entries: { key: string; value: string }[] = [];

  function visit(node: Element, path: string): void {
    const basePath = path ? `${path}/${node.tagName || node.nodeName}` : String(node.tagName || node.nodeName);
    for (const attr of Array.from(node.attributes)) {
      entries.push({
        key: `${basePath}/@${attr.name}`,
        value: attr.value
      });
    }
    const text = getDirectTextContent(node).trim();
    if (text) {
      entries.push({
        key: `${basePath}/#text`,
        value: text
      });
    }
    for (const child of Array.from(node.children)) {
      visit(child, basePath);
    }
  }

  visit(anchor, "");
  return entries;
}
```

この `rawEntries` は後段で階層 Markdown に整形され、必要なら SVG アセット参照も併記される。

### 22.16 Formula Resolver の骨格

`buildFormulaResolver(...)` は、Workbook 全体の sheet / cell / table / defined name を索引化し、数式 evaluator と legacy resolver が使う参照 API をまとめて提供する。

- 役割: 数式解決用の共通 resolver 群を組み立てる
- 入力: `ParsedWorkbook`
- 出力: `resolveCellValue`、`resolveRangeEntries`、`resolveDefinedNameRange` などを持つ resolver object
- 前後関係: `resolveSimpleFormulaReferences(...)` から 1 回構築され、AST / legacy の両経路で共有される

```ts
function buildFormulaResolver(workbook: ParsedWorkbook) {
  const sheetMap = new Map<string, ParsedSheet>();
  const cellMaps = new Map<string, Map<string, ParsedCell>>();
  const tableMap = new Map<string, ParsedTable>();
  for (const sheet of workbook.sheets) {
    sheetMap.set(sheet.name, sheet);
    const cellMap = new Map<string, ParsedCell>();
    for (const cell of sheet.cells) {
      cellMap.set(cell.address.toUpperCase(), cell);
    }
    cellMaps.set(sheet.name, cellMap);
    for (const table of sheet.tables) {
      if (table.name) {
        tableMap.set(normalizeStructuredTableKey(table.name), table);
      }
      if (table.displayName) {
        tableMap.set(normalizeStructuredTableKey(table.displayName), table);
      }
    }
  }

  const resolvingKeys = new Set<string>();
  const definedNameMap = new Map<string, string>();
  for (const entry of workbook.definedNames) {
    const key = entry.localSheetName
      ? `${normalizeFormulaSheetName(entry.localSheetName)}::${normalizeDefinedNameKey(entry.name)}`
      : `::${normalizeDefinedNameKey(entry.name)}`;
    definedNameMap.set(key, entry.formulaText);
  }

  function lookupDefinedNameFormula(sheetName: string, name: string): string | null {
    const normalizedName = normalizeDefinedNameKey(name);
    return definedNameMap.get(`${normalizeFormulaSheetName(sheetName)}::${normalizedName}`)
      || definedNameMap.get(`::${normalizedName}`)
      || null;
  }

  function resolveCellValue(sheetName: string, address: string): string {
    const sheet = sheetMap.get(sheetName);
    if (!sheet) return "#REF!";
    const cell = cellMaps.get(sheetName)?.get(address.toUpperCase()) || null;
    if (!cell) return "";
    const key = `${sheetName}!${address.toUpperCase()}`;
    if (resolvingKeys.has(key)) {
      return "";
    }
    if (cell.formulaText && (!cell.outputValue || cell.resolutionStatus !== "resolved")) {
      resolvingKeys.add(key);
      try {
        const result = tryResolveFormulaExpressionDetailed(
          cell.formulaText,
          sheetName,
          resolveCellValue,
          undefined,
          undefined,
          cell.address
        );
        if (result?.value != null) {
          applyResolvedFormulaValue(cell, result.value, result.source || "legacy_resolver");
        }
      } finally {
        resolvingKeys.delete(key);
      }
    }
    // ... formula cell / literal cell の返し分け
  }

  function resolveDefinedNameValue(sheetName: string, name: string): string | null { /* ... */ }
  function resolveDefinedNameRange(sheetName: string, name: string): { sheetName: string; start: string; end: string } | null { /* ... */ }
  function resolveStructuredRange(sheetName: string, text: string): { sheetName: string; start: string; end: string } | null { /* ... */ }
  function resolveSpillRange(sheetName: string, address: string): { sheetName: string; start: string; end: string } | null { /* ... */ }
  function resolveRangeEntries(sheetName: string, rangeText: string): { rawValues: string[]; numericValues: number[] } { /* ... */ }

  return {
    resolveCellValue,
    resolveRangeValues: (sheetName: string, rangeText: string) => resolveRangeEntries(sheetName, rangeText).numericValues,
    resolveRangeEntries,
    resolveDefinedNameValue,
    resolveDefinedNameRange,
    resolveStructuredRange
  };
}
```

### 22.17 Structured Reference / Defined Name / Spill

legacy / AST の両経路で重要なのは、参照テキストを実セル範囲へ落とす helper 群である。

- 役割: 抽象的な参照記法を具体的なセル範囲へ正規化する
- 入力: `sheetName`, 参照テキスト、table / defined name / spill 情報
- 出力: `{ sheetName, start, end }` 形式の range
- 前後関係: AST evaluator の `resolveName` / `resolveStructuredRef` と legacy resolver の range 解決で共通利用される

```ts
function resolveDefinedNameRange(sheetName: string, name: string): { sheetName: string; start: string; end: string } | null {
  const formulaText = lookupDefinedNameFormula(sheetName, name);
  if (!formulaText) return null;
  const normalized = formulaText.replace(/^=/, "").trim();
  const directRange = parseQualifiedRangeReference(normalized, sheetName);
  if (directRange) {
    return directRange;
  }
  // INDEX を終点に持つ range などの限定対応
  // ...
}

function resolveStructuredRange(sheetName: string, text: string): { sheetName: string; start: string; end: string } | null {
  const match = String(text || "").trim().match(/^(.+?)\[([^\]]+)\]$/);
  if (!match) return null;
  const tableKey = normalizeStructuredTableKey(match[1].replace(/^'(.*)'$/, "$1"));
  const columnKey = normalizeStructuredTableKey(match[2]);
  if (!tableKey || !columnKey || columnKey.startsWith("#") || columnKey.startsWith("@")) {
    return null;
  }
  const table = tableMap.get(tableKey);
  if (!table) return null;
  const columnIndex = table.columns.findIndex((columnName) => normalizeStructuredTableKey(columnName) === columnKey);
  if (columnIndex < 0) return null;
  const startAddress = parseCellAddress(table.start);
  const endAddress = parseCellAddress(table.end);
  if (!startAddress.row || !startAddress.col || !endAddress.row || !endAddress.col) return null;
  const firstDataRow = Math.min(startAddress.row, endAddress.row) + Math.max(0, table.headerRowCount);
  const lastDataRow = Math.max(startAddress.row, endAddress.row) - Math.max(0, table.totalsRowCount);
  if (firstDataRow > lastDataRow) return null;
  const col = Math.min(startAddress.col, endAddress.col) + columnIndex;
  const colLetters = colToLetters(col);
  return {
    sheetName: table.sheetName || sheetName,
    start: `${colLetters}${firstDataRow}`,
    end: `${colLetters}${lastDataRow}`
  };
}

function resolveSpillRange(sheetName: string, address: string): { sheetName: string; start: string; end: string } | null {
  const normalizedAddress = normalizeFormulaAddress(address);
  const cell = cellMaps.get(sheetName)?.get(normalizedAddress) || null;
  if (!cell) {
    return null;
  }
  if (cell.formulaType === "array") {
    return { sheetName, start: normalizedAddress, end: normalizedAddress };
  }
  const spillRef = String(cell.spillRef || "").trim();
  if (!spillRef) {
    return { sheetName, start: normalizedAddress, end: normalizedAddress };
  }
  const directRange = parseQualifiedRangeReference(spillRef, sheetName);
  if (directRange) {
    return directRange;
  }
  // ... A1:B3 / A1 単独表現も吸収
}
```

### 22.18 AST と Legacy の接続

数式再解決の優先順は `tryResolveFormulaExpressionDetailed(...)` に集約される。ここが `cached` 不在時の実際の順序である。

- 役割: AST evaluator と legacy resolver の優先順を 1 箇所で管理する
- 入力: `formulaText`, `currentSheetName`, 各種 resolver、`currentAddress`
- 出力: 解決値と `source`、または `null`
- 前後関係: `resolveSimpleFormulaReferences(...)` とセル単位再帰解決の中心に位置する

```ts
function tryResolveFormulaExpressionDetailed(
  formulaText: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string,
  resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
  resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
  currentAddress?: string
): { value: string; source: FormulaResolutionSource } | null {
  const normalized = String(formulaText || "").trim().replace(/^=/, "");
  if (!normalized) return null;
  const directDefinedNameValue = resolveDefinedNameScalarValue?.(currentSheetName, normalized) || null;
  if (directDefinedNameValue != null) {
    return {
      value: directDefinedNameValue,
      source: "legacy_resolver"
    };
  }
  const astResolved = tryResolveFormulaExpressionWithAst(
    normalized,
    currentSheetName,
    resolveCellValue,
    resolveRangeEntries,
    currentAddress
  );
  if (astResolved != null) {
    return {
      value: astResolved,
      source: "ast_evaluator"
    };
  }
  const legacyResolved = tryResolveFormulaExpressionLegacy(
    normalized,
    currentSheetName,
    resolveCellValue,
    resolveRangeValues,
    resolveRangeEntries
  );
  if (legacyResolved == null) {
    return null;
  }
  return {
    value: legacyResolved,
    source: "legacy_resolver"
  };
}
```

legacy 側は関数群を直列に試す dispatcher になっている。

- 役割: 既存 resolver 群を順番に試し、最初に解けた結果を返す
- 入力: 正規化済み数式文字列と各種 resolver
- 出力: 解決値、または `null`
- 前後関係: AST evaluator 失敗時のみ呼ばれる

```ts
function tryResolveFormulaExpressionLegacy(
  normalized: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string,
  resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
  resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
): string | null {
  const ifResult = tryResolveIfFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
  if (ifResult != null) return ifResult;
  const ifErrorResult = tryResolveIfErrorFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
  if (ifErrorResult != null) return ifErrorResult;
  const logicalResult = tryResolveLogicalFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
  if (logicalResult != null) return logicalResult;
  const concatResult = tryResolveConcatenationExpression(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
  if (concatResult != null) return concatResult;
  const numericFunctionResult = tryResolveNumericFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
  if (numericFunctionResult != null) return numericFunctionResult;
  // ... datePart / predicate / choose / text / lookup / string / conditionalAggregate / aggregate / comparison
  return null;
}
```

### 22.19 表示形式の補助関数

`formatCellDisplayValue(...)` を同等実装へ寄せるには、分岐本体だけでなく補助関数も必要である。特にゼロ値セクション、書式セクション分割、分数・指数処理は揃えて持つ必要がある。

- 役割: 表示形式本体が依存する小さな helper を揃える
- 入力: `formatCode` や zero section
- 出力: セクション分割結果やゼロ値表示文字列
- 前後関係: `formatCellDisplayValue(...)` の補助であり、単独では使わない

```ts
function splitFormatSections(formatCode: string): string[] {
  const sections: string[] = [];
  let current = "";
  let inQuotes = false;
  for (let index = 0; index < formatCode.length; index += 1) {
    const char = formatCode[index];
    if (char === "\"") {
      inQuotes = !inQuotes;
      current += char;
      continue;
    }
    if (char === ";" && !inQuotes) {
      sections.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  sections.push(current);
  return sections;
}

function formatZeroSection(section: string): string | null {
  const normalizedSection = String(section || "");
  if (!normalizedSection) return null;
  const compact = normalizedSection.replace(/_.|\\.|[*?]/g, "").trim();
  const hasDashLiteral = /"-"|(^|[^a-z0-9])-($|[^a-z0-9])/i.test(compact);
  if (!hasDashLiteral) return null;
  if (compact.includes("¥")) return "¥ -";
  if (compact.includes("$")) return "$ -";
  return "-";
}
```

### 22.20 なお不足しているもの

この章までで `impl-spec` 単体の再実装可能性はかなり上がったが、なお同等実装へ寄せるには次を追加するとよい。

- chart / shape helper 周辺の補助関数完全版
- office drawing helper の公開境界と SVG 生成本体
- UI event binding の完全版
- 一部 helper の完全実装はなおコード参照が必要

### 22.21 Path 正規化と chart helper

`parseRelationships(...)` が参照する `Target` は相対パスを含みうるため、`normalizeZipPath(...)` を揃えないと workbook / worksheet / drawing の rels 解決が再現できない。

- 役割: rels の `Target` を source XML 基準の ZIP 内絶対パスへ直す
- 入力: `baseFilePath`, `targetPath`
- 出力: 正規化済み ZIP path
- 前後関係: `parseRelationships(...)` から必ず呼ばれ、worksheet / drawing / chart / media の解決に共通利用される

```ts
function normalizeZipPath(baseFilePath: string, targetPath: string): string {
  const baseDirParts = baseFilePath.split("/").slice(0, -1);
  const inputParts = targetPath.split("/");
  const parts = targetPath.startsWith("/") ? [] : baseDirParts;
  for (const part of inputParts) {
    if (!part || part === ".") continue;
    if (part === "..") {
      parts.pop();
      continue;
    }
    parts.push(part);
  }
  return parts.join("/");
}
```

`parseDrawingCharts(...)` を同等実装へ寄せるには、`parseChartType(...)`、`parseChartTitle(...)`、`parseChartSeries(...)` も併せて持つ必要がある。

- `parseChartType(...)`
  - 役割: chart XML からグラフ種別ラベルを決める
  - 入力: `chartDoc`
  - 出力: `"棒グラフ"` や `"棒グラフ + 折れ線グラフ (複合)"` のような文字列
  - 前後関係: `parseDrawingCharts(...)` の `chartType` を構成する

- `parseChartTitle(...)`
  - 役割: chart XML の rich text からタイトル文字列を抽出する
  - 入力: `chartDoc`
  - 出力: タイトル文字列
  - 前後関係: `parseDrawingCharts(...)` の `title` を構成する

- `parseChartSeries(...)`
  - 役割: 系列名、categories 参照、values 参照、副軸判定を抽出する
  - 入力: `chartDoc`
  - 出力: `ParsedChartAsset["series"]`
  - 前後関係: `parseDrawingCharts(...)` の `series` を構成し、`## グラフ` セクション描画へ流れる

```ts
function parseChartType(chartDoc: Document): string {
  const typeMap: Array<{ localName: string; label: string }> = [
    { localName: "barChart", label: "棒グラフ" },
    { localName: "lineChart", label: "折れ線グラフ" },
    { localName: "pieChart", label: "円グラフ" },
    { localName: "doughnutChart", label: "ドーナツグラフ" },
    { localName: "areaChart", label: "面グラフ" },
    { localName: "scatterChart", label: "散布図" },
    { localName: "radarChart", label: "レーダーチャート" },
    { localName: "bubbleChart", label: "バブルチャート" }
  ];
  const matched = typeMap
    .filter((entry) => getElementsByLocalName(chartDoc, entry.localName).length > 0)
    .map((entry) => entry.label);
  if (matched.length === 0) return "グラフ";
  if (matched.length === 1) return matched[0];
  return `${matched.join(" + ")} (複合)`;
}

function parseChartTitle(chartDoc: Document): string {
  const richText = getElementsByLocalName(chartDoc, "t")
    .map((node) => getTextContent(node))
    .filter(Boolean);
  if (richText.length > 0) {
    return richText.join("").trim();
  }
  return "";
}

function parseChartSeries(chartDoc: Document): ParsedChartAsset["series"] {
  const plotArea = getFirstChildByLocalName(chartDoc, "plotArea") || chartDoc.documentElement;
  const axisPositionById = new Map<string, string>();
  for (const axisNode of getElementsByLocalName(plotArea, "valAx")) {
    const axisIdNode = getFirstChildByLocalName(axisNode, "axId");
    const axisPosNode = getFirstChildByLocalName(axisNode, "axPos");
    const axisId = axisIdNode?.getAttribute("val") || getTextContent(axisIdNode);
    const axisPos = axisPosNode?.getAttribute("val") || getTextContent(axisPosNode);
    if (axisId) {
      axisPositionById.set(axisId, axisPos || "");
    }
  }
  // ... chart container ごとに ser を列挙して name / categoriesRef / valuesRef / axis を構成する
}
```

### 22.22 UI 集計と診断表示

`main.ts` 側の実装は変換ロジック本体ではないが、formula diagnostics や table score の算出規則を再現したい場合はここも必要になる。

- 役割: `MarkdownFile.summary` を UI 向けの集計表示へ変換する
- 入力: `WorkbookFile[]`
- 出力: summary / table score / formula diagnostics の HTML
- 前後関係: `convertWorkbookToMarkdownFiles(...)` の結果を受けてプレビュー下部へ表示する

table score 集計は `score -> strong / candidate / unknown` のラベル変換を持つ。

```ts
function getTableScoreLabel(score: number): string {
  if (score >= 7) return "strong";
  if (score >= 4) return "candidate";
  return "unknown";
}

function renderTableScoreCounts(file: WorkbookFile): string {
  const counts = {
    strong: 0,
    candidate: 0,
    unknown: 0
  };
  file.summary.tableScores.forEach((detail) => {
    counts[getTableScoreLabel(detail.score) as keyof typeof counts] += 1;
  });
  return [
    counts.strong > 0 ? `strong ${counts.strong}` : "",
    counts.candidate > 0 ? `candidate ${counts.candidate}` : "",
    counts.unknown > 0 ? `unknown ${counts.unknown}` : ""
  ].filter(Boolean).join(" / ");
}
```

formula diagnostics 集計は `status` と `source` を別軸で数える。

```ts
function getFormulaStatusLabel(status: "resolved" | "fallback_formula" | "unsupported_external" | null): string {
  if (status === "resolved") return "resolved";
  if (status === "fallback_formula") return "fallback";
  if (status === "unsupported_external") return "unsupported";
  return "unknown";
}

function getFormulaSourceLabel(source: "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null): string {
  if (source === "cached_value") return "cached";
  if (source === "ast_evaluator") return "ast";
  if (source === "legacy_resolver") return "legacy";
  if (source === "formula_text") return "formula";
  if (source === "external_unsupported") return "external";
  return "unknown";
}

function renderFormulaSourceCounts(file: WorkbookFile): string {
  const counts = {
    cached: 0,
    ast: 0,
    legacy: 0,
    formula: 0,
    external: 0,
    unknown: 0
  };
  file.summary.formulaDiagnostics.forEach((diagnostic) => {
    counts[getFormulaSourceLabel(diagnostic.source) as keyof typeof counts] += 1;
  });
  return [
    counts.cached > 0 ? `cached ${counts.cached}` : "",
    counts.ast > 0 ? `ast ${counts.ast}` : "",
    counts.legacy > 0 ? `legacy ${counts.legacy}` : "",
    counts.formula > 0 ? `formula ${counts.formula}` : "",
    counts.external > 0 ? `external ${counts.external}` : "",
    counts.unknown > 0 ? `unknown ${counts.unknown}` : ""
  ].filter(Boolean).join(" / ");
}
```

全体 summary は workbook 名、sheet 数、mode、表数、地の文数、結合数、画像数、数式数、解析セル数を集計する。

```ts
function renderAnalysisSummary(files: WorkbookFile[], workbookName: string): string {
  if (files.length === 0) {
    return '<div class="md-summary-empty">まだ変換していません。</div>';
  }
  const totalTables = files.reduce((sum, file) => sum + file.summary.tables, 0);
  const totalNarratives = files.reduce((sum, file) => sum + file.summary.narrativeBlocks, 0);
  const totalMerges = files.reduce((sum, file) => sum + file.summary.merges, 0);
  const totalImages = files.reduce((sum, file) => sum + file.summary.images, 0);
  const totalCells = files.reduce((sum, file) => sum + file.summary.cells, 0);
  const totalFormulas = files.reduce((sum, file) => sum + file.summary.formulaDiagnostics.length, 0);
  const outputMode = files[0]?.summary.outputMode || "display";
  // ... overview / totals / per-sheet items を HTML 文字列として組み立てる
}
```

### 22.23 legacy resolver 個別関数の代表断片

`tryResolveFormulaExpressionLegacy(...)` の dispatcher だけでは、どの程度まで legacy resolver が式を解けるかが見えにくい。再実装時は、少なくとも代表的な関数群の実体も持っておくと挙動の境界が分かりやすい。

- 役割: AST evaluator で解けなかった式に対して、文字列ベースの追加解決を行う
- 入力: 正規化済み数式文字列、現在 sheet 名、セル解決関数群
- 出力: 解決できた場合は文字列値、できない場合は `null`
- 前後関係: `tryResolveFormulaExpressionDetailed(...)` の後段で呼ばれ、解けなければ `formula_text` へ落ちる

`IF` は条件式を評価し、真偽に応じて第 2 / 第 3 引数を `resolveScalarFormulaValue(...)` で再帰評価する。

```ts
function tryResolveIfFunction(
  normalizedFormula: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string,
  resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
  resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
): string | null {
  const call = parseWholeFunctionCall(normalizedFormula, ["IF"]);
  if (!call) return null;
  const args = splitFormulaArguments(call.argsText.trim());
  if (args.length !== 3) return null;
  const condition = evaluateFormulaCondition(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
  if (condition == null) return null;
  return resolveScalarFormulaValue(
    condition ? args[1] : args[2],
    currentSheetName,
    resolveCellValue,
    resolveRangeValues,
    resolveRangeEntries
  );
}
```

`TEXT` は第 1 引数の値を解決し、第 2 引数の書式文字列を `formatTextFunctionValue(...)` へ渡す。

```ts
function tryResolveTextFunction(
  normalizedFormula: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string,
  resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
  resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
): string | null {
  const call = parseWholeFunctionCall(normalizedFormula, ["TEXT"]);
  if (!call) return null;
  const args = splitFormulaArguments(call.argsText.trim());
  if (args.length !== 2) return null;
  const value = resolveScalarFormulaValue(
    args[0],
    currentSheetName,
    resolveCellValue,
    resolveRangeValues,
    resolveRangeEntries
  );
  const formatText = resolveScalarFormulaValue(
    args[1],
    currentSheetName,
    resolveCellValue,
    resolveRangeValues,
    resolveRangeEntries
  );
  if (value == null || formatText == null) return null;
  return formatTextFunctionValue(value, formatText);
}
```

`XLOOKUP` は完全一致の縦方向探索を優先し、検索 range と返却 range の長さ一致を前提に 1 件目を返す。`match_mode` と `search_mode` は限定値のみ受け付ける。

```ts
function tryResolveLookupFunction(
  normalizedFormula: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string
): string | null {
  const xlookupCall = parseWholeFunctionCall(normalizedFormula, ["XLOOKUP"]);
  if (xlookupCall) {
    const args = splitFormulaArguments(xlookupCall.argsText.trim());
    if (args.length < 3 || args.length > 6) return null;
    const lookupValue = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue);
    const lookupRange = parseQualifiedRangeReference(args[1], currentSheetName);
    const returnRange = parseQualifiedRangeReference(args[2], currentSheetName);
    if (lookupValue == null || !lookupRange || !returnRange) return null;
    const lookupCells = collectRangeCells(lookupRange, resolveCellValue);
    const returnCells = collectRangeCells(returnRange, resolveCellValue);
    if (lookupCells.length === 0 || lookupCells.length !== returnCells.length) return null;
    if (args.length >= 5) {
      const matchMode = resolveScalarFormulaValue(args[4], currentSheetName, resolveCellValue);
      if (matchMode == null || !["0", ""].includes(matchMode.trim())) {
        return null;
      }
    }
    if (args.length >= 6) {
      const searchMode = resolveScalarFormulaValue(args[5], currentSheetName, resolveCellValue);
      if (searchMode == null || !["1", ""].includes(searchMode.trim())) {
        return null;
      }
    }
    for (let index = 0; index < lookupCells.length; index += 1) {
      const value = lookupCells[index];
      if (value === lookupValue || (!Number.isNaN(Number(value)) && !Number.isNaN(Number(lookupValue)) && Number(value) === Number(lookupValue))) {
        return returnCells[index] ?? "";
      }
    }
    if (args.length >= 4) {
      return resolveScalarFormulaValue(args[3], currentSheetName, resolveCellValue);
    }
  }
  return null;
}
```

`AVERAGEIFS` / `SUMIFS` は range-criteria pair を並列に評価し、`matchesCountIfCriteria(...)` を満たす行だけを集計対象とする。

```ts
function tryResolveConditionalAggregateFunction(
  normalizedFormula: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string
): string | null {
  const averageifsCall = parseWholeFunctionCall(normalizedFormula, ["AVERAGEIFS"]);
  if (averageifsCall) {
    const args = splitFormulaArguments(averageifsCall.argsText.trim());
    if (args.length < 3 || args.length % 2 === 0) return null;
    const averageRange = parseQualifiedRangeReference(args[0], currentSheetName);
    if (!averageRange) return null;
    const averageCells = collectRangeCells(averageRange, resolveCellValue);
    if (averageCells.length === 0) return null;
    const rangeCriteriaPairs = [];
    for (let index = 1; index < args.length; index += 2) {
      const rangeRef = parseQualifiedRangeReference(args[index], currentSheetName);
      const criteria = resolveScalarFormulaValue(args[index + 1], currentSheetName, resolveCellValue);
      if (!rangeRef || criteria == null) return null;
      const cells = collectRangeCells(rangeRef, resolveCellValue);
      if (cells.length !== averageCells.length) return null;
      rangeCriteriaPairs.push({ cells, criteria });
    }
    let sum = 0;
    let count = 0;
    for (let i = 0; i < averageCells.length; i += 1) {
      if (!rangeCriteriaPairs.every((entry) => matchesCountIfCriteria(entry.cells[i], entry.criteria))) {
        continue;
      }
      const numeric = Number(averageCells[i]);
      if (!Number.isNaN(numeric)) {
        sum += numeric;
        count += 1;
      }
    }
    return count > 0 ? String(sum / count) : null;
  }
  return null;
}
```

### 22.24 プレビューとダウンロード導線

`main.ts` の UI 部分は変換本体ではないが、生成結果の見せ方と保存形式を同等再現したい場合はこの導線も必要になる。

- 役割: `currentFiles` と `currentWorkbook` を preview / Markdown 保存 / ZIP 保存へ接続する
- 入力: 変換済み `WorkbookFile[]` と workbook 名
- 出力: preview DOM 更新、Markdown Blob、ZIP Blob
- 前後関係: `convertWorkbookToMarkdownFiles(...)` 実行後に `renderCurrentSelection()` が呼ばれ、その後ユーザー操作で保存処理が起動される

preview 更新は 1 つの sheet を選ばせるのではなく、現行実装では全 sheet Markdown を連結して表示する。

```ts
function renderCurrentSelection(): void {
  if (!currentFiles.length) {
    setSummaryText("まだ変換していません。");
    setScoreSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
    setFormulaSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
    setPreviewMarkdown("");
    updatePreviewModeBanner(getSelectedOutputMode());
    return;
  }
  const combinedMarkdown = currentFiles
    .map((file) => `<!-- ${createMarkdownChunkLabel(file.fileName)} -->\n${file.markdown}`)
    .join("\n\n");
  const outputMode = currentFiles[0]?.summary.outputMode || "display";
  updatePreviewModeBanner(outputMode);
  setSummaryHtml(renderAnalysisSummary(currentFiles, currentWorkbook?.name || "workbook.xlsx"));
  setScoreSummaryHtml(renderScoreSummary(currentFiles));
  setFormulaSummaryHtml(renderFormulaSummary(currentFiles));
  setPreviewMarkdown(combinedMarkdown);
  getElement<HTMLButtonElement>("downloadBtn").disabled = false;
  getElement<HTMLButtonElement>("exportZipBtn").disabled = false;
}
```

Markdown 保存は `createCombinedMarkdownExportFile(...)` が返す `{ fileName, content }` をそのまま Blob 化して保存する。

```ts
function downloadCurrentMarkdown(): void {
  const payload = getSelectedFileForDownload();
  if (!payload) {
    showError("保存対象の Markdown がありません");
    return;
  }
  const blob = new Blob([`${payload.content}\n`], { type: "text/markdown;charset=utf-8" });
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = payload.fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
  showToast("Markdown を保存しました");
}
```

ZIP 保存は `createWorkbookExportArchive(...)` を直接呼び、保存名だけ UI 側で `_xlsx2md_export` を付加する。

```ts
function downloadExportZip(): void {
  if (!currentWorkbook || currentFiles.length === 0) {
    showError("先に Markdown を生成してください");
    return;
  }
  const zipBytes = xlsx2md.createWorkbookExportArchive(currentWorkbook, currentFiles);
  const outputMode = currentFiles[0]?.summary.outputMode || "display";
  const suffix = outputMode === "display" ? "" : `_${outputMode}`;
  const blob = new Blob([zipBytes], { type: "application/zip" });
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = `${currentWorkbook.name.replace(/\\.xlsx$/i, "")}_xlsx2md_export${suffix}.zip`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
  showToast("ZIP を保存しました");
}
```

### 22.25 shape SVG helper と `drawingHelper` 境界

図形抽出の再実装では、shape メタを読む処理と SVG を描画する処理が分離されている点を明示しておく必要がある。これがないと `parseDrawingShapes(...)` だけ見て SVG 生成の責務まで内包しているように誤読しやすい。

- 役割: `parseDrawingShapes(...)` は shape XML の走査とメタ抽出を担当し、SVG の具体描画は `drawingHelper.renderShapeSvg(...)` に委譲する
- 入力: `shapeNode`, `anchor`, `sheetName`, `shapeCounter`
- 出力: `svgPath` / `svgData` を持つ補助 asset、または `null`
- 前後関係: `drawingHelper` が存在しない場合でも shape 本体の抽出は継続され、SVG だけが省略される

`core.ts` では helper の存在を `globalThis` から読み、optional chaining で呼び出す。

```ts
const drawingHelper = (globalThis as typeof globalThis & {
  __xlsx2mdDrawingHelper?: {
    renderShapeSvg?: (
      shapeNode: Element,
      anchor: string,
      sheetName: string,
      shapeIndex: number
    ) => { svgPath: string; svgData: Uint8Array } | null;
  };
}).__xlsx2mdDrawingHelper;
```

`parseDrawingShapes(...)` の責務境界は次の 1 行に集約される。

```ts
const svgAsset = drawingHelper?.renderShapeSvg?.(shapeNode, anchor, sheetName, shapeCounter) || null;
```

したがって同等実装を狙う場合は、最低でも次の 2 層を分けて再実装する。

- shape metadata layer
  - anchor / kind / text / ext / rawEntries / elementName の抽出
- shape rendering layer
  - `spPr` や geometry を SVG へ変換し、ZIP 出力用 `svgPath` / `svgData` を返す helper

### 22.26 legacy resolver の補助関数

legacy resolver のうち `IF` だけでは、条件分岐まわりの解決限界がまだ見えにくい。`IFERROR` と論理関数を加えると、どこまで文字列ベースで拾うかがかなり具体化される。

- 役割: エラー代替と論理関数の簡易評価を行う
- 入力: 正規化済み数式文字列、現在 sheet 名、セル/範囲解決関数
- 出力: 解決文字列、または `null`
- 前後関係: `tryResolveFormulaExpressionLegacy(...)` から `IF` の近辺で順に呼ばれ、複雑な式は AST evaluator 側へ委ねる

`IFERROR` は第 1 引数が Excel エラー文字列でない限りそれを返し、エラーなら第 2 引数を返す。

```ts
function tryResolveIfErrorFunction(
  normalizedFormula: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string,
  resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
  resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
): string | null {
  const call = parseWholeFunctionCall(normalizedFormula, ["IFERROR"]);
  if (!call) return null;
  const args = splitFormulaArguments(call.argsText.trim());
  if (args.length !== 2) return null;
  const primary = resolveScalarFormulaValue(
    args[0],
    currentSheetName,
    resolveCellValue,
    resolveRangeValues,
    resolveRangeEntries
  );
  if (primary != null && !/^#(?:[A-Z]+\\/[A-Z]+|[A-Z]+[!?]?)/i.test(primary.trim())) {
    return primary;
  }
  return resolveScalarFormulaValue(
    args[1],
    currentSheetName,
    resolveCellValue,
    resolveRangeValues,
    resolveRangeEntries
  );
}
```

`AND` / `OR` / `NOT` は `evaluateFormulaCondition(...)` を使い、`null` を含む場合は保守的に `null` を返す。

```ts
function tryResolveLogicalFunction(
  normalizedFormula: string,
  currentSheetName: string,
  resolveCellValue: (sheetName: string, address: string) => string,
  resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
  resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
): string | null {
  const call = parseWholeFunctionCall(normalizedFormula, ["AND", "OR", "NOT"]);
  if (!call) return null;
  const functionName = call.name;
  const args = splitFormulaArguments(call.argsText.trim());
  if (functionName === "NOT") {
    if (args.length !== 1) return null;
    const value = evaluateFormulaCondition(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
    if (value == null) return null;
    return value ? "FALSE" : "TRUE";
  }
  if (args.length === 0) return null;
  const evaluations = args.map((arg) =>
    evaluateFormulaCondition(arg, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries)
  );
  if (functionName === "AND") {
    if (evaluations.some((value) => value === false)) return "FALSE";
    if (evaluations.some((value) => value == null)) return null;
    return (evaluations as boolean[]).every(Boolean) ? "TRUE" : "FALSE";
  }
  if (functionName === "OR") {
    if (evaluations.some((value) => value === true)) return "TRUE";
    if (evaluations.some((value) => value == null)) return null;
    return (evaluations as boolean[]).some(Boolean) ? "TRUE" : "FALSE";
  }
  return null;
}
```

### 22.27 output mode notice と preview banner

`main.ts` では output mode の違いを preview 本文だけでなく、notice と banner の文言でも明示している。ここを落とすと UI 再現時に mode 差分が分かりにくくなる。

- 役割: `display / raw / both` の意味を UI 文言で補足する
- 入力: 現在の output mode
- 出力: notice / banner の DOM 更新
- 前後関係: 変換前の mode 切替時と、変換後の preview 更新時の双方で呼ばれる

変換前の notice は選択中 mode の説明だけを出す。

```ts
function updateOutputModeNotice(mode: "display" | "raw" | "both"): void {
  const notice = getElement<HTMLElement>("outputModeNotice");
  if (mode === "raw") {
    notice.textContent = "`raw` は Excel の表示値ではなく、内部値を優先して Markdown に出力します。";
    return;
  }
  if (mode === "both") {
    notice.textContent = "`both` は表示値に加えて `[raw=...]` 形式の補助情報を出力します。";
    return;
  }
  notice.textContent = "`display` は Excel の表示値寄りで出力します。";
}
```

preview 上部の banner は `display` のとき非表示で、`raw` / `both` のときだけ表示する。

```ts
function updatePreviewModeBanner(mode: "display" | "raw" | "both"): void {
  const banner = getElement<HTMLElement>("previewModeBanner");
  if (mode === "raw") {
    banner.hidden = false;
    banner.textContent = "`raw` モードです。Markdown には Excel の表示値ではなく内部値が出ます。";
    return;
  }
  if (mode === "both") {
    banner.hidden = false;
    banner.textContent = "`both` モードです。Markdown には表示値に加えて `[raw=...]` が出ます。";
    return;
  }
  banner.hidden = true;
  banner.textContent = "";
}
```

### 22.28 error / loading 制御

preview と保存導線に加えて、UI の状態遷移を再現するには error alert と loading overlay の扱いも必要になる。ここが抜けると、失敗時の見え方と処理中表示が実装とずれる。

- 役割: エラー表示と処理中表示を DOM コンポーネントへ接続する
- 入力: message、active フラグ
- 出力: alert / overlay の更新
- 前後関係: workbook 読込、変換、保存の前後で呼ばれ、`renderCurrentSelection()` や `download...()` と同じ UI 層に属する

エラー表示は custom element の `show` / `clear` を優先し、なければ素の属性と textContent を操作する。

```ts
function clearError(): void {
  const errorAlert = getElement<HTMLElement>("errorAlert") as HTMLElement & { clear?: () => void };
  if (typeof errorAlert.clear === "function") {
    errorAlert.clear();
  } else {
    errorAlert.removeAttribute("active");
    errorAlert.textContent = "";
  }
}

function showError(message: string): void {
  const errorAlert = getElement<HTMLElement>("errorAlert") as HTMLElement & { show?: (text: string) => void };
  if (typeof errorAlert.show === "function") {
    errorAlert.show(message);
  } else {
    errorAlert.textContent = message;
    errorAlert.setAttribute("active", "");
  }
}
```

loading overlay も同様に `show` / `hide` を優先し、fallback として属性を直接切り替える。

```ts
function setLoading(active: boolean, message?: string): void {
  const overlay = getElement<HTMLElement>("loadingOverlay") as HTMLElement & { show?: (text?: string) => void; hide?: () => void };
  if (active) {
    if (message) {
      overlay.setAttribute("text", message);
    }
    if (typeof overlay.show === "function") {
      overlay.show(message || "処理中です");
    } else {
      overlay.setAttribute("active", "");
    }
    return;
  }
  if (typeof overlay.hide === "function") {
    overlay.hide();
  } else {
    overlay.removeAttribute("active");
  }
}
```

### 22.29 `parseChartSeries(...)` の完全版

chart helper の中では `parseChartSeries(...)` が最も情報量が多く、ここを省略したままだと `series` 配列の作り方が十分に追えない。副軸判定と `tx / cat / val` 解決まで含めると、現行実装の series 抽出をほぼ再現できる。

- 役割: chart XML から系列名、category 参照、value 参照、primary/secondary 軸を抽出する
- 入力: `chartDoc`
- 出力: `ParsedChartAsset["series"]`
- 前後関係: `parseDrawingCharts(...)` 内で `parseChartType(...)` / `parseChartTitle(...)` と並んで使われ、Markdown の `## グラフ` セクションへ流れる

```ts
function parseChartSeries(chartDoc: Document): ParsedChartAsset["series"] {
  const plotArea = getFirstChildByLocalName(chartDoc, "plotArea") || chartDoc.documentElement;
  const axisPositionById = new Map<string, string>();
  for (const axisNode of getElementsByLocalName(plotArea, "valAx")) {
    const axisIdNode = getFirstChildByLocalName(axisNode, "axId");
    const axisPosNode = getFirstChildByLocalName(axisNode, "axPos");
    const axisId = axisIdNode?.getAttribute("val") || getTextContent(axisIdNode);
    const axisPos = axisPosNode?.getAttribute("val") || getTextContent(axisPosNode);
    if (axisId) {
      axisPositionById.set(axisId, axisPos || "");
    }
  }

  const chartContainerNames = [
    "barChart",
    "lineChart",
    "pieChart",
    "doughnutChart",
    "areaChart",
    "scatterChart",
    "radarChart",
    "bubbleChart"
  ];
  const series: ParsedChartAsset["series"] = [];

  for (const localName of chartContainerNames) {
    for (const chartNode of getElementsByLocalName(plotArea, localName)) {
      const axisIds = getElementsByLocalName(chartNode, "axId")
        .map((node) => node.getAttribute("val") || getTextContent(node))
        .filter(Boolean);
      const isSecondary = axisIds.some((axisId) => axisPositionById.get(axisId) === "r");

      for (const seriesNode of getElementsByLocalName(chartNode, "ser")) {
        const txNode = getFirstChildByLocalName(seriesNode, "tx") || seriesNode;
        const nameRef = getFirstChildByLocalName(txNode, "f");
        const nameValue = getFirstChildByLocalName(txNode, "v");
        const nameText = getElementsByLocalName(txNode, "t")
          .map((node) => getTextContent(node))
          .join("")
          .trim();
        const catRef = getFirstChildByLocalName(getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "cat") || seriesNode, "strRef") || seriesNode, "f")
          || getFirstChildByLocalName(getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "cat") || seriesNode, "numRef") || seriesNode, "f");
        const valRef = getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "val") || seriesNode, "f")
          || getFirstChildByLocalName(getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "val") || seriesNode, "numRef") || seriesNode, "f");
        series.push({
          name: nameText || getTextContent(nameValue) || getTextContent(nameRef) || "系列",
          categoriesRef: getTextContent(catRef),
          valuesRef: getTextContent(valRef),
          axis: isSecondary ? "secondary" : "primary"
        });
      }
    }
  }

  return series;
}
```

### 22.30 toast と UI 状態遷移

保存や変換完了時の UX を同等再現したい場合は、toast と初期化/読込時の状態遷移も持っておく必要がある。ここがないと「いつ何をユーザーへ通知するか」が実装とずれる。

- 役割: 完了通知と初期化時の UI リセットを担当する
- 入力: message、workbook 読込結果、変換結果
- 出力: toast 表示、summary/preview/download ボタン状態の更新
- 前後関係: `loadWorkbookFromFile(...)`、`convertCurrentWorkbook(...)`、`initialize()` の中から呼ばれる

toast は custom element の `show` があればそれを使い、なければ `active` 属性と `textContent` を直接操作する。

```ts
function showToast(message: string): void {
  const toast = getElement<HTMLElement>("toast") as HTMLElement & { show?: (text: string) => void };
  if (typeof toast.show === "function") {
    toast.show(message);
  } else {
    toast.textContent = message;
    toast.setAttribute("active", "");
    window.setTimeout(() => toast.removeAttribute("active"), 1800);
  }
}
```

変換は `clearError()` から始まり、成功時だけ toast を出し、失敗時は `showError(...)` へ落とす。

```ts
function convertCurrentWorkbook(showSuccessToast = true): void {
  clearError();
  if (!currentWorkbook) {
    showError("先に xlsx ファイルを読み込んでください");
    return;
  }
  try {
    currentFiles = xlsx2md.convertWorkbookToMarkdownFiles(currentWorkbook, getOptions());
    renderCurrentSelection();
    if (showSuccessToast) {
      showToast("Markdown を生成しました");
    }
  } catch (error) {
    showError(error instanceof Error ? error.message : "Markdown 生成に失敗しました");
  }
}
```

workbook 読込は loading overlay を出し、成功時は parse 後に即 convert し、失敗時は preview と download 状態をまとめて初期状態へ戻す。

```ts
async function loadWorkbookFromFile(file: File): Promise<void> {
  clearError();
  setLoading(true, "xlsx を読み込んでいます");
  try {
    const arrayBuffer = await file.arrayBuffer();
    currentWorkbook = await xlsx2md.parseWorkbook(arrayBuffer, file.name);
    currentFiles = [];
    convertCurrentWorkbook(false);
    showToast("xlsx を読み込み、Markdown を生成しました");
  } catch (error) {
    currentWorkbook = null;
    currentFiles = [];
    setSummaryText("Workbook の読込に失敗しました。");
    setScoreSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
    setFormulaSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
    setPreviewMarkdown("");
    getElement<HTMLButtonElement>("downloadBtn").disabled = true;
    getElement<HTMLButtonElement>("exportZipBtn").disabled = true;
    showError(error instanceof Error ? error.message : "xlsx の読込に失敗しました");
  } finally {
    setLoading(false);
  }
}
```

初期化時も同じ空状態へ寄せる。

```ts
function initialize(): void {
  clearError();
  setSummaryText("まだ変換していません。");
  setScoreSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
  setFormulaSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
  setPreviewMarkdown("");
  updateOutputModeNotice(getSelectedOutputMode());
  updatePreviewModeBanner(getSelectedOutputMode());
  getElement<HTMLButtonElement>("downloadBtn").disabled = true;
  getElement<HTMLButtonElement>("exportZipBtn").disabled = true;
  bindFileInput();
  bindActions();
}
```

### 22.31 shape helper 周辺の補助関数

`parseDrawingShapes(...)` だけだと、`kind`、`text`、`ext`、`bbox` の求め方が分からず再実装しづらい。現行実装では shape helper 群を小さく分けている。

- 役割: shape XML から種別、テキスト、サイズ、bounding box を取り出す
- 入力: `shapeNode`、`anchor`
- 出力: `kind`、`text`、`widthEmu` / `heightEmu`、`bbox`
- 前後関係: `parseDrawingShapes(...)` から順に呼ばれ、最後に `extractShapeBlocks(...)` で近接図形を block 化する

shape 種別は connector / textbox / preset geometry 名の順で判定する。

```ts
function parseShapeKind(shapeNode: Element | null): string {
  if (!shapeNode) return "shape";
  if (shapeNode.localName === "cxnSp") return "connector";
  if (getElementsByLocalName(shapeNode, "txBody").length > 0) {
    const geomNode = getFirstChildByLocalName(getFirstChildByLocalName(shapeNode, "spPr") || shapeNode, "prstGeom");
    const geom = String(geomNode?.getAttribute("prst") || "").trim();
    if (geom === "rect" || geom === "") return "textbox";
  }
  const geomNode = getFirstChildByLocalName(getFirstChildByLocalName(shapeNode, "spPr") || shapeNode, "prstGeom");
  const geom = String(geomNode?.getAttribute("prst") || "").trim();
  return geom || "shape";
}
```

shape テキストは `txBody` 配下の `t` を連結して返す。

```ts
function parseShapeText(shapeNode: Element | null): string {
  if (!shapeNode) return "";
  return getElementsByLocalName(shapeNode, "t")
    .map((node) => getTextContent(node))
    .filter(Boolean)
    .join("")
    .trim();
}
```

サイズは `xfrm/ext` か anchor 直下の `ext` から取得し、bounding box は anchor の from/to と EMU 定数を使って計算する。

```ts
function parseShapeExt(anchor: Element, shapeNode: Element | null): { widthEmu: number | null; heightEmu: number | null } {
  const extNode =
    getDirectChildByLocalName(
      getDirectChildByLocalName(getDirectChildByLocalName(shapeNode || anchor, "spPr") || shapeNode || anchor, "xfrm"),
      "ext"
    )
    || getDirectChildByLocalName(anchor, "ext");
  const widthEmu = Number(extNode?.getAttribute("cx") || "");
  const heightEmu = Number(extNode?.getAttribute("cy") || "");
  return {
    widthEmu: Number.isFinite(widthEmu) ? widthEmu : null,
    heightEmu: Number.isFinite(heightEmu) ? heightEmu : null
  };
}

function parseShapeBoundingBox(anchor: Element, shapeNode: Element | null, widthEmu: number | null, heightEmu: number | null): {
  left: number;
  top: number;
  right: number;
  bottom: number;
} {
  const fromCol = parseAnchorInt(anchor, "from", "col") || 0;
  const fromRow = parseAnchorInt(anchor, "from", "row") || 0;
  const fromColOff = parseAnchorInt(anchor, "from", "colOff") || 0;
  const fromRowOff = parseAnchorInt(anchor, "from", "rowOff") || 0;
  const toCol = parseAnchorInt(anchor, "to", "col");
  const toRow = parseAnchorInt(anchor, "to", "row");
  const toColOff = parseAnchorInt(anchor, "to", "colOff") || 0;
  const toRowOff = parseAnchorInt(anchor, "to", "rowOff") || 0;

  const left = fromCol * DEFAULT_CELL_WIDTH_EMU + fromColOff;
  const top = fromRow * DEFAULT_CELL_HEIGHT_EMU + fromRowOff;

  if (toCol !== null && toRow !== null) {
    return {
      left,
      top,
      right: toCol * DEFAULT_CELL_WIDTH_EMU + toColOff,
      bottom: toRow * DEFAULT_CELL_HEIGHT_EMU + toRowOff
    };
  }

  const ext = parseShapeExt(anchor, shapeNode);
  return {
    left,
    top,
    right: left + Math.max(1, ext.widthEmu || widthEmu || DEFAULT_CELL_WIDTH_EMU),
    bottom: top + Math.max(1, ext.heightEmu || heightEmu || DEFAULT_CELL_HEIGHT_EMU)
  };
}
```

### 22.32 office drawing helper と SVG 生成本体

`core.ts` から見えるのは `drawingHelper.renderShapeSvg(...)` の呼び出しだけだが、同等実装を狙うなら `office-drawing.ts` 側の公開境界と SVG 生成本体も必要になる。現行実装は汎用ベクタ描画エンジンではなく、textbox / rect / connector を対象にした軽量 renderer である。

- 役割: shape XML から簡易 SVG を生成し、`assets/<sheet>/shape_XXX.svg` として返す
- 入力: `shapeNode`, `anchor`, `sheetName`, `shapeIndex`
- 出力: `{ filename, path, data } | null`
- 前後関係: `core.ts` では `globalThis.__xlsx2mdOfficeDrawing.renderShapeSvg` を `drawingHelper` として参照する

まず helper 側では、色・寸法・文字列・種別のための小さな補助関数を持つ。

```ts
function parseHexColor(root: Element | null): string | null {
  const srgb = getElementsByLocalName(root, "srgbClr")[0] || null;
  if (srgb?.getAttribute("val")) {
    return `#${String(srgb.getAttribute("val")).trim()}`;
  }
  const scheme = getElementsByLocalName(root, "schemeClr")[0] || null;
  const schemeVal = String(scheme?.getAttribute("val") || "").trim();
  const schemeMap: Record<string, string> = {
    accent1: "#4472C4",
    accent2: "#ED7D31",
    accent3: "#A5A5A5",
    accent4: "#FFC000",
    accent5: "#5B9BD5",
    accent6: "#70AD47",
    tx1: "#000000",
    tx2: "#44546A",
    lt1: "#FFFFFF",
    lt2: "#E7E6E6"
  };
  return schemeMap[schemeVal] || null;
}

function parseShapeDimensions(anchor: Element, shapeNode: Element | null): { widthPx: number; heightPx: number } {
  const extNode = getDirectChildByLocalName(anchor, "ext")
    || getDirectChildByLocalName(getDirectChildByLocalName(getDirectChildByLocalName(shapeNode || anchor, "spPr"), "xfrm"), "ext");
  const widthEmu = Number(extNode?.getAttribute("cx") || "");
  const heightEmu = Number(extNode?.getAttribute("cy") || "");
  return {
    widthPx: emuToPx(widthEmu, 160),
    heightPx: emuToPx(heightEmu, 48)
  };
}
```

rect / textbox 系は塗り色、線色、線幅を取り、中央寄せテキスト付きの `<rect>` SVG を返す。

```ts
function renderRectLikeSvg(shapeNode: Element, anchor: Element, text: string, treatAsTextbox: boolean): string {
  const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
  const spPr = getDirectChildByLocalName(shapeNode, "spPr");
  const fillColor = parseHexColor(getDirectChildByLocalName(spPr, "solidFill")) || (treatAsTextbox ? "#FFFFFF" : "#F3F3F3");
  const lineNode = getDirectChildByLocalName(spPr, "ln");
  const strokeColor = parseHexColor(lineNode) || "#333333";
  const strokeWidth = Math.max(1, Math.round(Number(lineNode?.getAttribute("w") || "") / 9525) || 1);
  const safeText = escapeXml(text);
  const textMarkup = safeText
    ? `<text x="${Math.round(widthPx / 2)}" y="${Math.round(heightPx / 2)}" text-anchor="middle" dominant-baseline="middle" font-size="14" font-family="sans-serif" fill="#000000">${safeText}</text>`
    : "";
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}">`,
    `  <rect x="1" y="1" width="${Math.max(1, widthPx - 2)}" height="${Math.max(1, heightPx - 2)}" fill="${fillColor}" stroke="${strokeColor}" stroke-width="${strokeWidth}"/>`,
    textMarkup ? `  ${textMarkup}` : "",
    `</svg>`
  ].filter(Boolean).join("\\n");
}
```

connector は中央水平線と矢印 marker のみを描画する。

```ts
function renderConnectorSvg(shapeNode: Element, anchor: Element): string {
  const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
  const spPr = getDirectChildByLocalName(shapeNode, "spPr");
  const lineNode = getDirectChildByLocalName(spPr, "ln");
  const strokeColor = parseHexColor(lineNode) || "#333333";
  const strokeWidth = Math.max(1, Math.round(Number(lineNode?.getAttribute("w") || "") / 9525) || 1);
  const effectiveHeight = Math.max(heightPx, 24);
  const y = Math.round(effectiveHeight / 2);
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${effectiveHeight}" viewBox="0 0 ${widthPx} ${effectiveHeight}">`,
    `  <defs>`,
    `    <marker id="arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto" markerUnits="strokeWidth">`,
    `      <path d="M0,0 L0,6 L9,3 z" fill="${strokeColor}"/>`,
    `    </marker>`,
    `  </defs>`,
    `  <line x1="2" y1="${y}" x2="${Math.max(2, widthPx - 4)}" y2="${y}" stroke="${strokeColor}" stroke-width="${strokeWidth}" marker-end="url(#arrow)"/>`,
    `</svg>`
  ].join("\\n");
}
```

公開関数 `renderShapeSvg(...)` は kind に応じて renderer を切り替え、最終的に UTF-8 の `Uint8Array` を返す。

```ts
function renderShapeSvg(shapeNode: Element, anchor: Element, sheetName: string, shapeIndex: number): SvgRenderResult {
  const kind = parseShapeKind(shapeNode);
  if (!kind) return null;
  let svg = "";
  if (kind === "connector") {
    svg = renderConnectorSvg(shapeNode, anchor);
  } else {
    svg = renderRectLikeSvg(shapeNode, anchor, parseShapeText(shapeNode), kind === "textbox");
  }
  const safeDir = createSafeSheetAssetDir(sheetName);
  const filename = `shape_${String(shapeIndex).padStart(3, "0")}.svg`;
  return {
    filename,
    path: `assets/${safeDir}/${filename}`,
    data: textEncoder.encode(`${svg}\\n`)
  };
}
```

helper の公開境界は `globalThis.__xlsx2mdOfficeDrawing` で、`core.ts` 側はこれを読み込んで使う。

```ts
(globalThis as typeof globalThis & {
  __xlsx2mdOfficeDrawing?: {
    renderShapeSvg: typeof renderShapeSvg;
  };
}).__xlsx2mdOfficeDrawing = {
  renderShapeSvg
};
```

### 22.33 UI event binding

`main.ts` の状態遷移は関数単体だけでは閉じず、file input と各 button の event binding まで見て初めて流れがつながる。ここを持っておくと、初期化から読込、変換、保存までの UI フローを文書だけで追いやすい。

- 役割: DOM 要素と `loadWorkbookFromFile(...)` / `convertCurrentWorkbook(...)` / 保存処理を結びつける
- 入力: file input change、button click、output mode change
- 出力: workbook 読込、再変換、保存、preview banner 更新
- 前後関係: `initialize()` の最後で `bindFileInput()` と `bindActions()` が呼ばれ、`DOMContentLoaded` 時に起動される

file input は最初の 1 ファイルだけを取り、`loadWorkbookFromFile(...)` に渡す。

```ts
function bindFileInput(): void {
  const fileInput = getElement<HTMLInputElement>("xlsxFileInput");
  fileInput.addEventListener("change", async () => {
    const file = fileInput.files?.[0];
    if (!file) return;
    await loadWorkbookFromFile(file);
  });
}
```

action binding は convert / markdown download / zip download / output mode change をそれぞれ独立に結びつける。

```ts
function bindActions(): void {
  getElement<HTMLButtonElement>("convertBtn").addEventListener("click", () => {
    convertCurrentWorkbook(true);
  });
  getElement<HTMLButtonElement>("downloadBtn").addEventListener("click", () => {
    downloadCurrentMarkdown();
  });
  getElement<HTMLButtonElement>("exportZipBtn").addEventListener("click", () => {
    downloadExportZip();
  });
  getElement<HTMLElement>("outputModeSelect").addEventListener("change", () => {
    const mode = getSelectedOutputMode();
    updateOutputModeNotice(mode);
    if (!currentFiles.length) {
      updatePreviewModeBanner(mode);
    }
  });
}
```

最終的な起動点は `DOMContentLoaded` で、初期空状態を作ったうえで binding を有効化する。

```ts
document.addEventListener("DOMContentLoaded", initialize);
```

### 22.34 なおコード参照が必要な helper 一覧

ここまでで主要経路は文書側へ移せたが、同等挙動を完全再現するには、なおいくつかの補助関数はコード本体を見た方が安全である。再実装時は次を「要参照 helper」として扱う。

- 役割: `impl-spec` 単体で追える範囲と、なおコード参照が必要な範囲の境界を明示する
- 入力: `core.ts` / `main.ts` の補助関数群
- 出力: 再実装時の参照優先リスト
- 前後関係: 実装開始時のチェックリストとして使う

数式 resolver では、次の helper が細かな互換性を左右する。

- `formatTextFunctionValue(...)`
  - `TEXT(...)` の書式文字列解釈を担う
- `evaluateFormulaCondition(...)`
  - `IF` / `AND` / `OR` / `NOT` の条件評価を担う
- `parseWholeFunctionCall(...)`
  - function name と引数文字列の安全な切り出しを担う
- `splitFormulaArguments(...)`
  - 入れ子括弧や文字列リテラルを考慮した引数分割を担う
- `resolveScalarFormulaValue(...)`
  - legacy resolver の多くが最終的に依存するスカラー値解決の中核
- `matchesCountIfCriteria(...)`
  - `SUMIFS` / `AVERAGEIFS` 系の criteria 判定を担う

UI 層では、次の helper が実装量は小さいが再現性に効く。

- `getElement(...)`
  - 例外化を含む DOM 取得の共通入口
- `getOptions(...)`
  - checkbox / select から `MarkdownOptions` を構成する
- `getSelectedOutputMode(...)`
  - `display / raw / both` の選択値を型付きで返す

現時点の `impl-spec` は、少なくとも次のレベルまでは単独で追える。

- workbook parse から markdown export までの主経路
- formula diagnostics と table detection の主要ロジック
- chart / image / shape metadata 抽出
- shape SVG の軽量 renderer の骨格
- UI の preview / save / error / loading / toast / binding の流れ

一方で、次のレベルはなお「コードを見て詰める」前提である。

- 書式文字列の細かな互換性
- legacy resolver 補助関数群の完全互換
- 引数分割や criteria 判定の細部
- 一部 helper の edge case 処理

### 22.35 再実装時の推奨手順

`impl-spec` を読みながら別実装を起こす場合、章順どおりに作るよりも依存順に組み立てた方が失敗しにくい。現行実装を踏まえると、次の順で作るのが最も安定する。

- 役割: `impl-spec` を使った再実装の着手順を示す
- 入力: 本文 1 章から 22 章までの仕様と実装参考コード
- 出力: 再実装時の作業順序
- 前後関係: 新規実装、移植実装、他言語化の開始時に参照する

1. Workbook / XML 基盤を作る  
   対象章: 4, 5, 22.1, 22.8, 22.13, 22.21  
   まず ZIP 展開、XML 読込、rels 解決、sharedStrings、styles、definedNames、sheet モデルを固める。

2. セル値と表示形式を作る  
   対象章: 6, 7.1-7.3, 22.2, 22.12, 22.19  
   通常セル、formatted display、`display / raw / both`、数式セルの cached 判定を作る。

3. 数式 evaluator と fallback を作る  
   対象章: 7.4-7.9, 8, 22.3, 22.14, 22.16, 22.17, 22.18, 22.23, 22.26  
   `cached -> AST -> legacy -> formula_text` の流れと diagnostics を作る。

4. 表検出と narrative 抽出を作る  
   対象章: 9, 10, 11, 22.4, 22.5, 22.9  
   table 候補、score、merge token、地の文、リスト化を作る。

5. drawing 系 metadata を作る  
   対象章: 12, 13, 14, 22.10, 22.15, 22.21, 22.29, 22.31  
   image / chart / shape の anchor と metadata を作る。

6. Markdown 組み立てと export を作る  
   対象章: 15, 16, 22.6, 22.24  
   sheet markdown、combined markdown、ZIP export を作る。

7. UI と diagnostics 表示を作る  
   対象章: 17, 18, 22.22, 22.27, 22.28, 22.30, 22.33  
   summary、score、formula diagnostics、preview、download、loading を作る。

8. shape SVG renderer を後付けする  
   対象章: 22.25, 22.32  
   core から独立度が高いので最後でよい。metadata 抽出と分離して実装する。

この順にすると、少なくとも次の中間成果物を段階的に確認できる。

- 3 まで: 数式込みの sheet model が得られる
- 4 まで: 表と地の文を含む markdown が出せる
- 6 まで: 実用的な export が完成する
- 8 まで: 現行 UI / asset 出力へかなり近づく

### 22.36 再実装時の確認観点

再実装では、コードを書いた直後に「何を見て合否判断するか」がないとズレが残りやすい。現行 `xlsx2md` に寄せるなら、少なくとも次の観点で段階確認するとよい。

- 役割: 再実装の途中成果物を評価する確認観点を示す
- 入力: 再実装した parser / converter / UI
- 出力: 実装差分の早期発見
- 前後関係: 22.35 の各段階の後で使う

Workbook / sheet 基盤の確認:

- workbook 名、sheet 名、sheet 数が安定して取れているか
- `sharedStrings` と `styles` の解決で通常セルの display 値が崩れていないか
- `definedNames` と table 定義が sheet モデルへ入っているか

数式系の確認:

- `cachedValueState` が `present_nonempty / present_empty / absent` を区別できるか
- `resolutionStatus` と `resolutionSource` が分かれているか
- `cached -> AST -> legacy -> formula_text` の順で落ちるか
- 外部参照が `unsupported_external` になるか

表検出と narrative の確認:

- 表候補が `seed cell` の連結成分から作られているか
- score が `strong / candidate / unknown` のどれに入るか追えるか
- 表に採用したセルが narrative 側へ二重出力されないか
- merge が `[MERGED←]` / `[MERGED↑]` で出るか

drawing 系の確認:

- image / chart / shape が anchor を持って抽出されるか
- chart に title / type / series が入るか
- shape に `kind / text / rawEntries / bbox` が入るか
- SVG を有効にした場合だけ `shape_XXX.svg` が出るか

Markdown / export の確認:

- sheet 単位 Markdown と combined Markdown の順序が安定しているか
- output mode ごとに file 名 suffix が一致するか
- ZIP に combined Markdown と assets が入るか

UI の確認:

- preview が全 sheet 連結表示になるか
- summary / table score / formula diagnostics が表示されるか
- `raw / both` のときだけ preview banner が出るか
- 読込失敗時に preview と download ボタンが初期状態へ戻るか

fixture や実データで検証する場合は、`tests/fixtures/README.md` と `local-data-review.md` を併読すると、どの入力がどの論点に効くか追いやすい。

### 22.37 本文書の保守ルール

`impl-spec` は現行実装仕様として使う以上、実装変更時にどこを更新すべきかが明示されていた方が保守しやすい。今後の更新は少なくとも次の単位で同期するとよい。

- 役割: 実装変更時の文書同期ポイントを明示する
- 入力: `core.ts` / `main.ts` / `office-drawing.ts` の変更
- 出力: `impl-spec` と関連 md の整合維持
- 前後関係: 実装変更後のレビュー、PR 作成、文書更新時に参照する

数式処理を変えた場合:

- `7. 数式セル処理仕様`
- `8. 数式診断仕様`
- `22.2`, `22.3`, `22.14`, `22.16`, `22.17`, `22.18`, `22.23`, `22.26`
- 必要に応じて `xlsx-formula-subset.md`

表検出や narrative を変えた場合:

- `9. 表検出仕様`
- `10. 地の文抽出仕様`
- `11. 結合セル仕様`
- `22.4`, `22.5`, `22.9`
- 必要に応じて `local-data-review.md`

drawing 系を変えた場合:

- `12. 画像仕様`
- `13. グラフ仕様`
- `14. 図形仕様`
- `22.10`, `22.15`, `22.21`, `22.25`, `22.29`, `22.31`, `22.32`
- 必要に応じて `tests/fixtures/README.md`

Markdown / export / UI を変えた場合:

- `15. Markdown 組み立て仕様`
- `16. 出力ファイル仕様`
- `17. UI 上の表示仕様`
- `22.6`, `22.22`, `22.24`, `22.27`, `22.28`, `22.30`, `22.33`
- 必要に応じて `README.md`

`xlsx2md-spec.md` との関係では、次のルールで切り分ける。

- 現行実装の挙動変更
  - まず `xlsx2md-impl-spec.md` を更新する
- 上位方針や将来構想の変更
  - `xlsx2md-spec.md` を更新する
- 入口説明や利用者向けの見え方変更
  - `README.md` を更新する

この切り分けを保つと、`README`、`spec`、`impl-spec` の役割が混ざりにくい。

### 22.38 用語と状態値の早見表

本文後半では `status`、`source`、`cachedValueState`、table score label など、似た値が繰り返し出てくる。実装や文書を読み返すときの参照先として、ここにまとめておく。

- 役割: 頻出する enum 相当値やラベルの意味を一覧化する
- 入力: `core.ts` と `main.ts` で使う状態値
- 出力: 読み返し用の早見表
- 前後関係: 数式診断、table detection、UI summary を読むときに参照する

`resolutionStatus`

| 値 | 意味 |
| --- | --- |
| `resolved` | 最終的に値を確定できた |
| `fallback_formula` | 値解決できず、式文字列をそのまま出力する |
| `unsupported_external` | 外部参照を含むため未対応として扱う |

`resolutionSource`

| 値 | 意味 |
| --- | --- |
| `cached_value` | `<v>` の cached value を採用した |
| `ast_evaluator` | AST evaluator で解決した |
| `legacy_resolver` | 文字列ベースの legacy resolver で解決した |
| `formula_text` | 解決できず式文字列へ落とした |
| `external_unsupported` | 外部参照として未対応扱いにした |

`cachedValueState`

| 値 | 意味 |
| --- | --- |
| `present_nonempty` | `<v>` 要素があり、中身も空ではない |
| `present_empty` | `<v>` 要素はあるが、中身は空文字 |
| `absent` | `<v>` 要素が存在しない |
| `null` | 数式セルではない、または該当なし |

table score label

| 値 | 意味 |
| --- | --- |
| `strong` | score 7 以上 |
| `candidate` | score 4 以上 7 未満 |
| `unknown` | score 4 未満 |

output mode

| 値 | 意味 |
| --- | --- |
| `display` | Excel の表示値寄りで出力する |
| `raw` | 内部値を優先して出力する |
| `both` | 表示値に加えて `[raw=...]` を補助出力する |

shape kind の代表値

| 値 | 意味 |
| --- | --- |
| `textbox` | text box として扱う矩形系図形 |
| `rect` | 長方形図形 |
| `connector` | 線・コネクタ図形 |
| `shape` | 上記に当てはまらない一般図形 |

### 22.39 最低限の確認入力

再実装や大きな改修のあと、全 fixture を毎回精査するのは重い。最初の段階では、少数の入力で主要経路が通るかを見るだけでも差分検知の効率がよい。

- 役割: 最低限確認すべき入力タイプを整理する
- 入力: fixture、実データ、または同等の `.xlsx`
- 出力: 初期スモークチェックの観点
- 前後関係: 22.35 の実装順、22.36 の確認観点と併用する

最小セットとしては、少なくとも次を用意するとよい。

1. 単純表 + 通常セル中心の workbook  
   目的: workbook parse、sheet model、基本 Markdown 出力の確認  
   観点: 表検出、display 値、combined Markdown、ZIP 出力

2. 数式中心の workbook  
   目的: `cachedValueState`、`resolutionStatus`、`resolutionSource` の確認  
   観点: `cached -> AST -> legacy -> formula_text`、外部参照の扱い、数式診断

3. narrative / リスト混在 workbook  
   目的: 表と地の文の分離、list 化条件の確認  
   観点: narrative block、箇条書き化、表との二重出力防止

4. 画像を含む workbook  
   目的: drawing rels 解決と assets 出力の確認  
   観点: `assets/<sheet>/image_XXX.ext`、anchor、Markdown の `## 画像`

5. chart を含む workbook  
   目的: chart metadata 抽出の確認  
   観点: chart type、title、series、`## グラフ`

6. shape を含む workbook  
   目的: shape metadata と SVG 出力の確認  
   観点: `kind / text / rawEntries / bbox`、`shape_XXX.svg`

7. merge 多用 workbook  
   目的: merge 展開と表出力の確認  
   観点: `[MERGED←]` / `[MERGED↑]`、代表セルの採用、表レイアウト維持

実リポジトリでは、詳細な fixture 対応は `tests/fixtures/README.md` にまとまっている。最初はそこで各 fixture の `主目的` と `対応章` を見て、上の 7 類型へ割り当てるとよい。

実データで確認する場合は、`local-data-review.md` にある `効く仕様論点` を参照すると、どの workbook が

- 表検出
- レイアウト分解
- 数式 evaluator 優先度
- 表示差分

のどれに効くかをすぐ判断できる。

### 22.40 典型的な差分症状と疑う箇所

再実装や改修では、出力差分が出ても原因の当たりを付けられないと調査が長引きやすい。現行実装に寄せる観点では、症状ごとに最初に疑う場所を持っておくと効率がよい。

- 役割: 差分症状から調査開始点を引けるようにする
- 入力: Markdown 差分、diagnostics 差分、asset 差分、UI 差分
- 出力: 最初に見るべき章・関数群
- 前後関係: 22.36 の確認観点で差分を見つけた後に使う

`cached` のはずが `ast` や `legacy` になる:

- まず疑う箇所
  - `7.2 cached value の扱い`
  - `8.3 source`
  - `22.2 数式セル読込と cached 判定`
- 典型原因
  - `<v>` 要素の有無と空文字を区別できていない
  - `cachedValueState` が `present_empty` になっていない

数式診断の `status` と `source` が噛み合わない:

- まず疑う箇所
  - `8.2 status`
  - `8.3 source`
  - `22.14`, `22.18`, `22.22`
- 典型原因
  - `resolutionStatus` 更新と `resolutionSource` 更新の順がずれている
  - fallback 時に `formula_text` へ落とす処理が漏れている

表が過剰に分割される、または narrative に吸われる:

- まず疑う箇所
  - `9. 表検出仕様`
  - `10. 地の文抽出仕様`
  - `22.4`, `22.5`, `22.9`
- 典型原因
  - seed cell 判定が弱い
  - 連結成分の近傍条件が変わっている
  - 採用済みセルの narrative 除外が漏れている

`[MERGED←]` / `[MERGED↑]` が出ない、または位置がずれる:

- まず疑う箇所
  - `11. 結合セル仕様`
  - `15.3 テーブル描画`
  - `22.6 Markdown 組み立て`
- 典型原因
  - merge range 展開の方向判定がずれている
  - 代表セル以外を空文字のままにしている

画像や chart が出ない:

- まず疑う箇所
  - `12. 画像仕様`
  - `13. グラフ仕様`
  - `22.10 drawing 抽出`
  - `22.21 Path 正規化と chart helper`
- 典型原因
  - rels の `Target` 正規化が誤っている
  - drawing rels と sheet rels の参照経路が混ざっている

chart の title / series が空になる:

- まず疑う箇所
  - `13. グラフ仕様`
  - `22.21`
  - `22.29`
- 典型原因
  - `t` / `f` / `v` の拾い分けが足りない
  - `plotArea` 以下の chart container 列挙が不足している

shape は出るが SVG が出ない:

- まず疑う箇所
  - `14. 図形仕様`
  - `22.25`
  - `22.32`
- 典型原因
  - `drawingHelper` / `__xlsx2mdOfficeDrawing` の公開境界がつながっていない
  - `renderShapeSvg(...)` が `kind` 判定で `null` を返している

`raw / both` の見え方が違う:

- まず疑う箇所
  - `6.3 outputMode ごとの出力値`
  - `15. Markdown 組み立て仕様`
  - `22.24`, `22.27`
- 典型原因
  - 出力値選択と UI notice/banner の両方を直していない
  - filename suffix と本文の mode 表示がずれている

ZIP の中身や保存名が違う:

- まず疑う箇所
  - `16. 出力ファイル仕様`
  - `22.6`
  - `22.24`
- 典型原因
  - combined Markdown ではなく sheet 単位で出している
  - UI 側の保存名規則と core 側の export 規則が一致していない

UI だけ古い状態に見える:

- まず疑う箇所
  - `17. UI 上の表示仕様`
  - `22.22`, `22.28`, `22.30`, `22.33`
- 典型原因
  - `renderCurrentSelection()` を呼ぶ前提が崩れている
  - 読込失敗時の reset や初期化処理が不足している

### 22.41 変更種別ごとの確認入力の当て方

差分症状の当たりを付けられても、次にどの fixture や実データを当てるか迷うと確認が遅くなる。変更種別ごとに優先して見る入力類型を持っておくと、確認コストをかなり下げられる。

- 役割: 変更内容から、最初に当てるべき入力を引けるようにする
- 入力: 実装変更の種類
- 出力: 優先して確認する fixture / 実データ類型
- 前後関係: 22.39 の「最低限の確認入力」と、22.40 の「典型的な差分症状」の間をつなぐ

数式 evaluator / diagnostics を変えた場合:

- 優先して当てる入力
  - 数式中心の workbook
  - 外部参照や未対応関数を含む workbook
  - `cachedValueState` の違いが出る workbook
- まず見る観点
  - `status`
  - `source`
  - `cachedValueState`
  - `raw / both` 出力差

表検出や narrative 抽出を変えた場合:

- 優先して当てる入力
  - 単純表 + 通常セル中心の workbook
  - narrative / リスト混在 workbook
  - merge 多用 workbook
  - レイアウト中心の実データ
- まず見る観点
  - table score
  - narrative block
  - 二重出力
  - merge token

表示形式や output mode を変えた場合:

- 優先して当てる入力
  - 日付 / 時刻 / 通貨 / パーセントを含む workbook
  - `TEXT(...)` を含む workbook
  - `display / raw / both` 差が見えやすい workbook
- まず見る観点
  - 表示値
  - raw 値
  - filename suffix
  - preview banner / notice

image / chart / shape を変えた場合:

- 優先して当てる入力
  - 画像あり workbook
  - chart あり workbook
  - shape あり workbook
  - image + chart + shape が混在する実データ
- まず見る観点
  - anchor
  - asset path
  - `## 画像` / `## グラフ` / `## 図形`
  - chart series
  - shape SVG

export / ZIP / UI を変えた場合:

- 優先して当てる入力
  - sheet 数が複数ある workbook
  - assets を含む workbook
  - 読込失敗も再現できる入力
- まず見る観点
  - combined Markdown
  - ZIP 中身
  - download file 名
  - disabled button / loading / toast / error の状態遷移

fixture を選ぶときは `tests/fixtures/README.md` の `主目的` と `対応章` を先に見る。実データを選ぶときは `local-data-review.md` の `効く仕様論点` を先に見る。この順にすると、変更箇所と確認入力の対応が取りやすい。

### 22.42 実装変更レビュー時の観点

`impl-spec` は再実装だけでなく、既存実装への変更レビューにも使える。レビュー時は「コードが正しそうか」だけでなく、「既存仕様との関係が明確か」を見ると差分の質が上がる。

- 役割: 実装変更のレビュー時に見る観点を整理する
- 入力: `core.ts` / `main.ts` / `office-drawing.ts` の差分
- 出力: 仕様差分、回帰リスク、文書更新漏れの発見
- 前後関係: 実装変更後、PR 前、または自己レビュー時に参照する

数式系の変更では、次を確認する。

- `resolutionStatus` と `resolutionSource` を同時に更新しているか
- `cachedValueState` の扱いが空文字と要素なしを区別しているか
- `cached -> AST -> legacy -> formula_text` の順序を崩していないか
- `xlsx-formula-subset.md` へ反映すべき変更がないか

表検出と narrative の変更では、次を確認する。

- table 採用済みセルが narrative 側へ漏れていないか
- score や label の閾値が UI 表示と一致しているか
- merge token の方向と代表セルの扱いが崩れていないか
- `xlsx2md-spec.md` の上位方針と矛盾していないか

drawing 系の変更では、次を確認する。

- rels 解決と asset path 規則が一致しているか
- chart title / series / type のどれかが欠けていないか
- shape metadata と shape SVG の責務境界が崩れていないか
- `README.md` の説明や fixture 対応表の更新が必要でないか

Markdown / export / UI の変更では、次を確認する。

- combined Markdown と ZIP 出力の規則が揃っているか
- file 名 suffix と output mode 表示が揃っているか
- preview / summary / diagnostics / download button の状態遷移が崩れていないか
- `README.md` と `impl-spec` の両方を更新すべき変更でないか

文書更新の観点では、少なくとも次を確認する。

- 現行実装の挙動変更なら `xlsx2md-impl-spec.md` を更新したか
- 上位方針や将来構想の変更なら `xlsx2md-spec.md` を更新したか
- 入口説明やユーザー向け見え方の変更なら `README.md` を更新したか
- fixture の主目的や対応章が変わるなら `tests/fixtures/README.md` を更新したか

### 22.43 変更影響の見取り図

`xlsx2md` は `core.ts` の中で多くの処理がつながっているため、1 箇所の変更がどこへ波及するかを把握していないとレビューや検証が漏れやすい。ここでは、代表的な変更点ごとの影響先をまとめる。

- 役割: 変更箇所から影響範囲を素早く見積もれるようにする
- 入力: 実装上の変更ポイント
- 出力: 影響しやすい章、出力、UI、確認入力
- 前後関係: 22.37 の保守ルール、22.40 の差分症状、22.42 のレビュー観点と併用する

数式セル読込まわりを変えた場合:

- 主な変更点
  - `<v>` / `<f>` 読込
  - `cachedValueState`
  - `resolutionStatus`
  - `resolutionSource`
- 影響先
  - Markdown の数式出力
  - 数式診断一覧
  - summary の formula 件数
  - `raw / both` の見え方
- 主に見る章
  - 7, 8, 22.2, 22.22, 22.38

AST / legacy resolver を変えた場合:

- 主な変更点
  - evaluator 呼び出し順
  - resolver helper
  - fallback 条件
- 影響先
  - `source` の内訳
  - `fallback_formula` 件数
  - display 値と diagnostics の整合
- 主に見る章
  - 7.4-7.9, 22.3, 22.16-22.18, 22.23, 22.26

表検出を変えた場合:

- 主な変更点
  - seed cell 判定
  - 連結成分
  - score 算出
- 影響先
  - 表数
  - narrative block 数
  - table score summary
  - sheet Markdown の構造
- 主に見る章
  - 9, 10, 17.2, 22.4, 22.5, 22.9, 22.22

drawing rels / path 解決を変えた場合:

- 主な変更点
  - `parseRelationships(...)`
  - `buildRelsPath(...)`
  - `normalizeZipPath(...)`
- 影響先
  - image / chart / shape 抽出全体
  - asset path
  - ZIP の中身
- 主に見る章
  - 12, 13, 14, 16.3, 22.10, 22.13, 22.21

chart helper を変えた場合:

- 主な変更点
  - chart type
  - title
  - series
- 影響先
  - `## グラフ` セクション
  - chart metadata
  - 実データ review の差分
- 主に見る章
  - 13, 22.21, 22.29

shape helper / SVG を変えた場合:

- 主な変更点
  - `kind`
  - `text`
  - `bbox`
  - SVG renderer
- 影響先
  - `## 図形` セクション
  - `shape_XXX.svg`
  - shape block 化や asset 出力
- 主に見る章
  - 14, 22.15, 22.25, 22.31, 22.32

Markdown export / UI を変えた場合:

- 主な変更点
  - combined Markdown
  - ZIP naming
  - summary / diagnostics HTML
  - preview banner / notice
- 影響先
  - ダウンロードファイル名
  - preview 本文
  - summary pill / diagnostics 表示
  - README の説明
- 主に見る章
  - 15, 16, 17, 22.6, 22.22, 22.24, 22.27-22.30, 22.33
