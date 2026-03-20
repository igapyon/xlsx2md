# xlsx2md Excel Formula Subset

## 概要

この文書は、`xlsx2md` が扱う Excel 数式のサブセットを整理するための設計メモです。

目的は次の通りです。

- 既存の場当たり的な式対応を、対象範囲が明確な仕様へ寄せる
- 将来の `AST` ベース実装に向けて、最小限の構文を整理する
- `ClosedXML.Parser` の ABNF を参照しつつ、`xlsx2md` に必要な部分だけを切り出す

参照元:

- [MS-XLSX-parser-grammar.abnf](./references/MS-XLSX-parser-grammar.abnf)
- 出典プロジェクト: <https://github.com/ClosedXML/ClosedXML.Parser>

## 位置づけ

この文書は、Excel 数式の完全仕様ではありません。

## 対応状況サマリ

| 区分 | 内容 |
| --- | --- |
| 実装済み | `cached value` 優先、`AST evaluator`、既存 resolver、`definedNames`、sheet scope name、structured reference の一部、shared formula 展開 |
| 部分対応 | dynamic array / spill、配列定数、`space intersection`、lookup 系や集計系の一部、表示形式再適用との組み合わせ |
| 未対応 | Excel 数式の完全互換、lambda 系、future function 群、外部 Workbook 参照の完全解決、dynamic array の全面対応 |

- 対象は `xlsx2md` が実務上扱いたいサブセットに限る
- `cached value` を優先する現在方針は維持する
- 自前評価は、`cached value` が無い場合や比較・補完が必要な場合の補助手段とする

### cached value と自前評価の関係

`xlsx2md` における数式 evaluator の位置づけは、Excel 数式の完全再実装ではない。

- 通常の `.xlsx` では、多くの数式セルに `cached value` が保存されている
- そのため、通常利用では `cached value` が主役であり、自前 evaluator は補助機能として扱う
- 自前 evaluator の主目的は、未計算保存ブックや一部欠損ブックで情報欠落を減らすことにある
- したがって、重要なのは Excel 数式全体への完全対応ではなく、`cached value` 優先を崩さずに不足分を補うことである

要するに、`xlsx2md` の数式 evaluator は次の役割を持つ。

- `cached value` がある場合はそれを採用する
- `cached value` が無い場合は、対応済みサブセットの範囲で自前評価を試みる
- それでも解けない場合は、式文字列保持や `fallback_formula` により情報欠落を避ける

この方針により、`xlsx2md` は Excel 数式の完全互換ではなく、Markdown 変換に必要な範囲での実務的な補完を目指す。

### 数式解決の優先順

`xlsx2md` における数式解決は、現在は次の優先順を前提とする。

1. `cached value`
2. `AST evaluator`
3. 従来の文字列ベース resolver
4. `fallback_formula` または式文字列保持

補足:

- `cached value` が存在する場合は、それを最優先で採用する
- `cached value` が欠損している場合のみ、`AST evaluator` を先に試す
- `AST evaluator` で未対応、または評価不能な場合は、従来の文字列ベース resolver へフォールバックする
- それでも解けない場合に限り、`fallback_formula` として保持する

この順序により、既存の実装資産を活かしながら、対応済みサブセットを段階的に AST 側へ寄せる。

### 既存 resolver の位置づけ

既存の文字列ベース resolver は、当面は削除対象ではない。

- 現時点では、`AST evaluator` が未対応の式を救うための互換 fallback として必要である
- `local-data` の実戦ブックで、想定外の式や Excel 特有の揺れに対する安全装置として機能する
- したがって、短期的には「安全のために残す」方針を採る

一方で、中長期的な方向性は次の通りである。

- 主要ケースは `AST evaluator` 側へ段階的に寄せる
- `existing resolver` の担当範囲は観測しながら縮小する
- 実データ上で依存が薄くなった段階で、互換 fallback として最小限へ絞る

要するに、`existing resolver` は今は必要だが、将来的には主役ではなく後方互換の安全装置へ寄せていく。

## 対応対象

### 1. 基本トークン

- 数値リテラル
- 文字列リテラル
- 論理値リテラル `TRUE` / `FALSE`
- エラー定数の一部
  - `#N/A`
  - `#REF!`
  - `#VALUE!`

### 2. 参照

- A1 形式セル参照
  - `A1`
  - `$A$1`
  - `A$1`
  - `$A1`
- シート付き参照
  - `Sheet1!A1`
  - `'日本語シート'!C4`
- 範囲参照
  - `A1:B5`
  - `Sheet1!A1:B5`
- 名前定義参照
  - workbook scope
  - sheet scope
- 構造化参照の一部
  - `課題[期日]`

備考:

- parser 上は sheet scope name を独立した `scoped_name` として扱う

### 3. 演算子

- 単項演算
  - `+`
  - `-`
- 後置演算
  - `%`
  - `#`
- 二項演算
  - `+`
  - `-`
  - `*`
  - `/`
  - `&`
- 比較演算
  - `=`
  - `<>`
  - `<`
  - `<=`
  - `>`
  - `>=`

### 4. 優先対応関数

現在実装済み、または `xlsx2md` で重要度が高いものを優先関数とする。

- 条件
  - `IF`
  - `IFERROR`
- 論理
  - `AND`
  - `OR`
  - `NOT`
- 参照
  - `INDEX`
  - `MATCH`
  - `VLOOKUP`
  - `HLOOKUP`
  - `XLOOKUP`
- 集計
  - `SUM`
  - `SUMPRODUCT`
  - `AVERAGE`
  - `MIN`
  - `MAX`
  - `COUNT`
  - `COUNTA`
  - `COUNTIF`
  - `SUMIF`
  - `AVERAGEIF`
  - `COUNTIFS`
  - `SUMIFS`
  - `AVERAGEIFS`
- 文字列
  - `TEXT`
  - `LEFT`
  - `RIGHT`
  - `MID`
  - `LEN`
  - `LOWER`
  - `UPPER`
  - `TRIM`
  - `SUBSTITUTE`
  - `REPLACE`
  - `REPT`
  - `CONCATENATE`
  - `FIND`
  - `SEARCH`
- 日付・数値
  - `DATE`
  - `VALUE`
  - `DATEVALUE`
  - `TODAY`
  - `WEEKDAY`
  - `YEAR`
  - `MONTH`
  - `DAY`
  - `ROW`
  - `COLUMN`
  - `EDATE`
  - `EOMONTH`
  - `ROUND`
  - `ROUNDUP`
  - `ROUNDDOWN`
  - `INT`
  - `ABS`
- 判定
  - `ISBLANK`
  - `ISNUMBER`
  - `ISTEXT`
  - `ISERROR`
  - `ISNA`
  - `NA`

## 当面の非対象

初期サブセットでは次を対象外とする。

- lambda 系
  - `LAMBDA`
  - `LET`
  - `MAP`
  - `REDUCE`
  - `SCAN`
- 動的配列の完全対応
  - ただし最小対応として、`A1#` のような spill 演算子は parser / evaluator で扱える
- 配列定数の完全対応
  - ただし最小対応として、`{1,2;3,4}` や `{1+2,A1;DATE(2024,3,17),4}` のような配列定数は parser / evaluator で扱える
- 名前管理や参照スタイルの完全互換
- volatile 関数の完全再計算
  - `NOW`
  - `TODAY`
  - `RAND`
  - `RANDBETWEEN`
- Excel のすべての future function
- 完全な `R1C1` 文法

## 最小構文モデル

`xlsx2md` の数式サブセットは、概念上次のような構文を持つ。

```text
Formula
  = Expression

Expression
  = Comparison

Comparison
  = Concat (comp_op Concat)?

Concat
  = Additive ("&" Additive)*

Additive
  = Multiplicative (("+" | "-") Multiplicative)*

Multiplicative
  = Unary (("*" | "/") Unary)*

Unary
  = ("+" | "-")? Primary

Primary
  = Number
  | String
  | Boolean
  | CellRef
  | RangeRef
  | NameRef
  | StructuredRef
  | FunctionCall
  | "(" Expression ")"
```

補足:

- `RangeRef` は関数引数文脈で主に使う
- `NameRef` は workbook scope / sheet scope を含む
- `StructuredRef` は必要最小限のみ扱う
- `FunctionCall` は上記の優先対応関数に限定する

## 参照 ABNF との対応イメージ

`ClosedXML.Parser` の ABNF と、`xlsx2md` サブセットの関係は次のように考える。

- `expression`
  - そのまま採る
- `constant`
  - 一部のみ採る
- `cell-reference`
  - A1 形式中心で採る
- `name-reference`
  - 採る
- `structure-reference`
  - 一部のみ採る
- `function-call`
  - 関数一覧全体は採らず、必要関数だけに限定する

つまり、ABNF 全体を実装対象にするのではなく、`xlsx2md` に必要な productions をサブセット化して用いる。

## 実装方針

### 現行方針

- `cached value` を最優先
- `cached value` 欠損時は `AST evaluator` を先に試す
- `AST evaluator` が未対応または失敗時のみ、従来の文字列ベース resolver を使う
- 解けない式は `fallback_formula`

### 次段方針

今後は、文字列ベースのパターン追加から次の段階へ寄せる。

1. 式文字列をサブセット parser で `AST` 化する
2. `AST evaluator` で評価する
3. 既存の文字列ベース evaluator は互換用 fallback とする

補足:

- `dynamic array / spill` については、`A1#` の parser / evaluator / core 側の入口までは追加済み
- ただし runtime では `f@ref` を使う都合上、現時点では `t="array"` を spill と見なさない保守的な扱いにしている
- 実ブックでの厳密な判定は、dynamic array 実例 fixture を追加してから詰める

## 導入ステップ案

### Step 1

- 現在の evaluator を維持
- この文書で対象サブセットを固定

### Step 2

- `Primary / Unary / Additive / Comparison` を対象に小さな parser を導入
- `CellRef / NameRef / FunctionCall` の基本形を `AST` 化

現状:

- `src/xlsx2md/ts/formula/tokenizer.ts`
- `src/xlsx2md/ts/formula/parser.ts`
- parser で対応済み:
  - 絶対参照
    - `$A$1`
    - `A$2`
    - `$H3`
  - sheet scope name
    - `Other!LocalCross`
  - error constant
    - `#N/A`
  - row qualifier 付き structured reference
    - `チェックリスト[[#This Row],[数量]]`
    - `タスク[[#This Row],[完了?]]`
  - 配列定数の最小対応
    - `{1,2;3,4}`
    - `{1+2,A1;DATE(2024,3,17),4}`
  - spill 演算子の最小対応
    - `A1#`
  - `space intersection` の最小対応
    - `A1:C3 B2:D4`

### Step 3

- `AST evaluator` を追加し、既存の文字列ベース resolver と並走させる
- 解ける式は AST 側を先に試し、未対応や失敗時は従来 resolver へフォールバックする

現状:

- `src/xlsx2md/ts/formula/evaluator.ts`
- `src/xlsx2md/ts/core.ts` から限定的に AST フック済み
- evaluator で対応済みの主な関数:
  - 条件・論理
    - `IF`
    - `IFERROR`
    - `AND`
    - `OR`
    - `NOT`
  - 参照
    - `INDEX`
    - `MATCH`
    - `VLOOKUP`
    - `HLOOKUP`
    - `XLOOKUP`
  - 集計
    - `SUM`
    - `SUMPRODUCT`
    - `AVERAGE`
    - `MIN`
    - `MAX`
    - `COUNT`
    - `COUNTA`
    - `COUNTIF`
    - `COUNTIFS`
    - `SUMIF`
    - `SUMIFS`
    - `AVERAGEIF`
    - `AVERAGEIFS`
    - `SUBTOTAL` の一部
      - `1 / 4 / 5 / 9`
      - `101 / 104 / 105 / 109`
  - 文字列
    - `TEXT`
    - `LEFT`
    - `RIGHT`
    - `MID`
    - `LEN`
    - `LOWER`
    - `TRIM`
    - `SUBSTITUTE`
    - `REPLACE`
    - `REPT`
    - `UPPER`
    - `CONCATENATE`
    - `FIND`
    - `SEARCH`
  - 日付・数値
    - `DATE`
    - `VALUE`
    - `DATEVALUE`
    - `TODAY`
    - `WEEKDAY`
    - `YEAR`
    - `MONTH`
    - `DAY`
    - `ROW`
    - `COLUMN`
    - `EDATE`
    - `EOMONTH`
    - `ROUND`
    - `ROUNDUP`
    - `ROUNDDOWN`
    - `INT`
    - `ABS`
  - 判定
    - `ISBLANK`
    - `ISNUMBER`
    - `ISTEXT`
    - `ISERROR`
    - `ISNA`
    - `NA`

### 観測状況

- `scripts/observe-xlsx2md-formulas.mjs` で `local-data` を観測可能
- 現在は parser 観点では `local-data/` の全対象で `ast_ng 0`
- 次段の焦点は parser ではなく、AST evaluator をどこまで優先し、どの関数群をさらに寄せるか

## 未決事項

- `ROW / COLUMN` の引数なし形は、現在セル参照が分かる文脈でのみ AST evaluator が対応する
- 文脈なしでの `ROW() / COLUMN()` をどう扱うか
- `VLOOKUP / HLOOKUP` の近似一致は AST evaluator で最小対応済み
- `XLOOKUP` の `match_mode / search_mode` は最小対応済み
- `XLOOKUP` の wildcard `match_mode=2` は最小対応済み
- `XLOOKUP` の binary search `search_mode=2/-2` は最小対応済み
- 近似一致の境界条件や未ソート範囲をどこまで扱うか
- 既存の文字列ベース resolver と AST evaluator の優先順をどこまで入れ替えるか
- `space` intersection の完全対応をどこまで扱うか
- 配列定数の完全対応をどの段階で扱うか
- `TODAY` / `NOW` を cached value 専用に留めるか
- `ClosedXML.Parser` の ABNF をどこまで文書上へ写経するか

## メモ

- `xlsx2md` の目的は Excel 数式の完全互換ではない
- ただし、実データで頻出する実務関数は優先的に吸収する
- そのため、完全 parser よりも「対象を絞ったサブセット parser」の方が方針として適している
