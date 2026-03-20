# TODO xlsx2md

## 実装タスク

- fixture 用 Excel ブックを追加する
  - `tests/fixtures/formula/formula-spill-sample01.xlsx`
  - `tests/fixtures/merge/merge-multiline-sample01.xlsx`
    - 結合セル内の改行付きテキストを確認する fixture
    - 追加時は fixture だけでなく Markdown 正規化ポリシー変更もセットで見直す
    - 現状は `markdown-normalize.ts` と `sheet-markdown.ts` で改行を空白化し、`xlsx2md-sheet-markdown.test.js` もその前提
- formula 次段タスク
  - `scripts/observe-xlsx2md-formulas.mjs` による観測を継続し、AST evaluator 側へ寄せる関数群を整理する
  - 優先順は `cached value -> AST evaluator -> 既存 resolver -> fallback_formula` で固定
  - 既存の文字列ベース evaluator は互換 fallback として維持しつつ、AST evaluator 側へ段階移行する
  - 既存 resolver は現時点では安全装置として必要であり、短期的には削除しない
  - 中長期的には、実データ観測を踏まえて担当範囲を縮小し、後方互換 fallback へ寄せる
  - 次候補:
    - `XLOOKUP` の binary search `search_mode=2/-2` の境界条件を必要に応じて詰める
    - 実データ頻出関数を AST evaluator 側へさらに寄せる
- `セクション分割ブロック` の実装導入順を決める
- UI の `formulaDiagnostics` / `tableScores` 表示を見直す
- Markdown エスケープを統一する
  - 表セル、narrative、見出し、箇条書きで共通方針を持つ
  - 少なくとも `改行 / | / \`` を安全に扱う
  - 必要に応じて行頭の Markdown 記号 (`#`, `-`, `*`, `>`) も整理する
  - 結合セル内の改行を `<br>` として許容するか、別の表現にするかを決める

## 未対応事項

- 数式未対応の整理
  - `space intersection` の完全対応
  - 配列定数の完全対応
  - dynamic array / spill の完全対応
  - `LAMBDA / LET / MAP / REDUCE / SCAN`
  - 完全な `R1C1` 文法
  - Excel の future function 全般
  - `NOW` など volatile 関数の完全再計算
- レイアウト未対応の整理
  - `セクション分割ブロック` の導入
  - `カレンダー / ボード / ダッシュボード系` シートの専用扱い
  - レイアウト中心シートの完全再現は対象外であり、`セクション / 表 / リスト / 画像` 分解で扱う
  - `イベント プランナー` のようなフォーム風罫線領域は、現時点では保守的に扱う
  - DrawingML の図形 (`xdr:sp` / `xdr:cxnSp` など) は、現時点では安全に無視またはメタデータ抽出に留める
  - `DrawingML -> SVG` は将来候補
  - グラフは当面、意味情報のテキスト化で固定し、`Chart -> SVG` は保留とする
  - SmartArt は現時点では fallback とし、意味解釈や SVG 化の対象外とする

## 方針未確定

- `XLOOKUP` 近似一致や binary search を未ソート範囲でどう扱うか
- `ROW / COLUMN` の文脈なし引数なし形をどう扱うか
- 配列定数をどこまで Excel 互換で広げるか
- `A1#` のような spill 演算子を、runtime でどこまで実解決するか
- `f@ref` を spill と array formula でどう見分けるか
- `TODAY / NOW` を cached value 専用に留めるか
- `existing resolver` から AST evaluator へ、どこまで段階移行したら縮小判断するか

## レイアウト系の整理

- `local-data` の実データ差分レビュー継続
  - 重点対象は `docs/local-data-review.md` を参照
  - 優先順:
    - `TFc2b640.../イベント プランナー` の多画像・多結合差分確認
    - `TF97739.../月間プランナー` の merge 多用差分確認
  - 人手確認があると助かるもの:
    - 代表シートの Excel スクリーンショット
    - 条件付き書式やアイコン列の意味説明
- レイアウト中心シート方針の維持と具体化
  - 見た目再現ではなく、`セクション / 表 / リスト / 画像` への分解を優先する
- `セクション分割ブロック` 導入検討
  - 対象候補: 入力パネル、概要カード、見出し付きの広い merge 領域
- `カレンダー / ボード / ダッシュボード系` シートの別カテゴリ化検討
  - 対象候補: `TF97739.../月間プランナー`
  - 対象候補: `TFc2b640.../イベント プランナー`
  - 対象候補: `TF0ffdef.../老後資金プランナー`

## 進捗メモ

- fixture ベースの実ファイル調整は一段落
- formula Step 2 の最小 parser 土台は追加済み
- formula Step 3 の最小 evaluator 土台は追加済み
- `core.ts` に no-op の `extractSectionBlocks(...)` は追加済み

## 参照

- 正本: `docs/TODO.md`
- 関連仕様:
  - `docs/xlsx2md-spec.md`
  - `docs/xlsx2md-impl-spec.md`
  - `docs/xlsx-formula-subset.md`
