# xlsx2md local-data review

`docs/xlsx2md/local-data/` に置いた実データについて、`xlsx2md` の重点確認対象を整理するメモ。

関連文書:

- 概要と使い方: [README.md](./README.md)
- 上位仕様と設計方針: [xlsx2md-spec.md](./xlsx2md-spec.md)
- 現行実装の詳細仕様: [xlsx2md-impl-spec.md](./xlsx2md-impl-spec.md)
- 数式サブセットの検討メモ: [xlsx-formula-subset.md](./xlsx-formula-subset.md)

## 現状サマリ

2026-03-18 時点で、`local-data` の Excel 系ファイルは 10 件。

機械集計では、ほとんどの Workbook が parse 可能で、数式も概ね解決できている。

加えて、`scripts/observe-xlsx2md-formulas.mjs` で AST 観測を継続している。

- parse / fallback は既存 resolver で安定している
- 2026-03-18 時点では、観測対象の `local-data` について AST parse は全件 `ast_ng 0`
- そのため、次段は parser 未対応の穴埋めではなく、AST evaluator の適用範囲整理と表示差分確認が中心になる
- `dynamic array / spill` については、`local-data` 直確認の範囲では明確な `A1#` 利用例は未確認
- worksheet XML の `f@ref` は確認できるが、見えているものは主に `t="array"` であり、spill 実例の確認には別サンプルが必要

重点確認対象:

- `TF0ffdef6d-9a19-4593-bdde-924d9e0aba2d7153cdcc_wac-2a2872261254.xlsx`
  - `計算`: formulas `2317`, resolved `2317`, fallback `0`, ast ok `2317`, ast ng `0`
  - parse / fallback は解消済み。AST parse 観点でも解消済み
  - 効く仕様論点: 数式 evaluator 優先度、表示差分
- `TFc2b640a6-8ee1-4258-9669-a8ab0b41240fb1a9c9ca_wac-2fedb8d0a784.xlsx`
  - `イベント プランナー`: images `18`, merges `174`
  - 画像と結合セルが多く、レイアウト由来の意味落ち確認に向く
  - 効く仕様論点: レイアウト分解、表検出、画像アンカー保持
- `TF97739ac3-cc1c-40fb-8682-f809e067145e8f18ec64_wac-912bbff00931.xlsx`
  - `月間プランナー`: merges `71`, formulas `128`, resolved `128`
  - 数式は解けているが、見た目の意味再現を確認したい
  - 効く仕様論点: 表検出、レイアウト分解、表示差分
- `TFe6ae6c3f-7542-4b5a-80ce-d131a04ae548d10af21f_wac-3845acf76347.xlsx`
  - `To Do リスト`: リストブロック化確認に使用
  - `買い物リスト`: 会計書式ゼロ値 `¥ -` の確認に使用
  - 効く仕様論点: リスト化、表示形式

## Workbook 別メモ

### TF0ffdef6d-9a19-4593-bdde-924d9e0aba2d7153cdcc_wac-2a2872261254.xlsx

効く仕様論点: 数式 evaluator 優先度、表示差分、レイアウト分解

観測結果:

- `老後資金プランナー`
  - formulas `801`, resolved `801`, fallback `0`
  - images `1`, merges `25`
  - ast parse ok `801`, ng `0`
- `計算`
  - formulas `2317`, resolved `2317`, fallback `0`
  - images `0`, merges `0`
  - ast parse ok `2317`, ng `0`

結論:

- `計算` シートは parse / fallback / AST parse ともに解消済み
- 次は `老後資金プランナー` の巨大表分割方針を詰める

目視差分:

- `老後資金プランナー` シートは、Excel 上では「グラフ + 入力パネル + 詳細表」の複合レイアウトだが、現状 Markdown では `B1-J59` 全体が 1 つの巨大表として出る
- 画面上部のグラフ説明や入力セクション見出しまで同一表へ吸われており、人間にとっては読みにくい
- この種のシートは、巨大表 1 個よりも「導入文 / 入力ブロック / 詳細表 / 画像」のような分割が望ましい

### TF2a72be1c-7be5-413d-b345-417c06878d3ab665d7ad_wac-acd2741d3bcc.xlsx

効く仕様論点: structured reference / defined name、数式 evaluator 優先度、表示差分

観測結果:

- `課題`
  - formulas `12`, resolved `12`, fallback `0`
  - ast parse ok `12`, ng `0`
- `月単位のビュー`
  - formulas `86`, resolved `86`, fallback `0`
  - ast parse ok `86`, ng `0`
- `週単位のビュー`
  - formulas `85`, resolved `85`, fallback `0`
  - ast parse ok `85`, ng `0`

結論:

- structured reference, defined name, `EOMONTH`, 反復再解決により fallback `0` まで到達済み
- 表示上の差分が残るかどうかの目視比較
- AST parse は解消済みなので、以後は表示差分や runtime 適用順の観測が中心

### TF97739ac3-cc1c-40fb-8682-f809e067145e8f18ec64_wac-912bbff00931.xlsx

効く仕様論点: 表検出、レイアウト分解、画像参照、表示差分

観測結果:

- `月間プランナー`
  - formulas `128`, resolved `128`, fallback `0`
  - images `1`, merges `71`
  - ast parse ok `128`, ng `0`
- `リスト`
  - formulas `9`, resolved `9`, fallback `0`
  - ast parse ok `9`, ng `0`

結論:

- 大量 merge を含むシートで、表検出・地の文・保存画像参照が自然か

目視差分:

- `月間プランナー` は Excel 上ではカレンダー/ボード系レイアウトだが、現状 Markdown では曜日列ごとに多数の小表へ分解される
- `目標と優先事項`、前月・翌月ミニカレンダー、各曜日の予定欄が個別表になっており、元の月間カレンダーとしてのまとまりは失われる
- この種のシートは通常の表検出だけでは不十分で、`カレンダー/ボード系` という別カテゴリの検討が必要

### TFc2b640a6-8ee1-4258-9669-a8ab0b41240fb1a9c9ca_wac-2fedb8d0a784.xlsx

効く仕様論点: レイアウト分解、表検出、画像アンカー保持、数式 evaluator 優先度

観測結果:

- `イベント プランナー`
  - formulas `11`, resolved `11`, fallback `0`
  - images `18`, merges `174`
  - ast parse ok `11`, ng `0`
- `支出`
  - formulas `23`, resolved `23`, fallback `0`
  - ast parse ok `23`, ng `0`
- `収入`
  - formulas `36`, resolved `36`, fallback `0`
  - ast parse ok `36`, ng `0`
- `概要`
  - formulas `6`, resolved `6`, fallback `0`
  - ast parse ok `6`, ng `0`

結論:

- 多画像シートでのアンカー位置と Markdown 出力の妥当性
- merge 多用シートでの表/地の文/画像の切り分け
- `SUBTOTAL`, `UPPER` は AST parse 観点では解消済み。以後は evaluator 優先度と表示差分を確認する

目視差分:

- `イベント プランナー` は Excel 上では装飾・画像・複数セクションが大きく効くレイアウト文書
- 現状 Markdown では `C10-AM14` などの広い merge 領域がそのまま巨大表として出ており、`[MERGED←]` が大量に並ぶ
- `議題`、`イベント チェックリスト`、`イベント カテゴリ`、`主な連絡先` といったセクションの存在は取れているが、視覚レイアウト依存の意味は落ちる
- この種のシートは、表の完全再現ではなく「セクション分割 + 表抽出 + 画像位置保持」に寄せるのが妥当
- なお、上部の入力フォームのような空欄の多い罫線領域については、現時点では保守的に扱う
- 表候補スコアが高い場合は表として残してよいが、横に広く疎で merge が多い領域は表候補から外して narrative / section として扱うこともある
- 将来的には `フォームブロック` や `入力パネル` としての別扱いを検討する余地がある

### TFe6ae6c3f-7542-4b5a-80ce-d131a04ae548d10af21f_wac-3845acf76347.xlsx

効く仕様論点: リスト化、表示形式、軽量レイアウト確認

観測結果:

- `買い物リスト`
  - formulas `20`, resolved `20`, fallback `0`
  - images `1`, merges `2`
  - ast parse ok `20`, ng `0`
- `予算の内訳`
  - formulas `3`, resolved `3`, fallback `0`
  - ast parse ok `3`, ng `0`
- `To Do リスト`
  - formulas `0`
- `共有リスト`
  - formulas `0`, images `1`

結論:

- `To Do リスト` は 1 パラグラフではなく、リストブロックとして扱う実装を追加済み
- `買い物リスト` の会計書式ゼロ値は `¥0.00` ではなく `¥ -` を優先する実装を追加済み

## 次の優先順

1. `TFc2b640.../イベント プランナー` の多画像・多結合差分を確認
2. `TF97739.../月間プランナー` の merge 多用シートを確認
3. `TF0ffdef.../老後資金プランナー` の巨大表分割方針を詰める
4. `dynamic array / spill` の実例 fixture を追加して runtime を検証する

## dynamic array / spill メモ

- `A1#` を parser / evaluator / core の最小入口までは追加済み
- ただし、2026-03-18 時点の `local-data` 直確認では、明示的な spill 利用例は未確認
- 現在 `core.ts` は worksheet XML の `f@ref` を `spillRef` として保持している
- ただし `f@ref` は `t="array"` にも現れるため、実運用上の挙動差分確認には dynamic array 実例 workbook が必要

## 人手確認があると助かるもの

- 上記 1-3 の代表シートについて、Excel 画面のスクリーンショット
- 特に「これを Markdown でどう見せたいか」がある場合は、その期待イメージ
- 条件付き書式やアイコンが意味を持つ列について、意味説明
