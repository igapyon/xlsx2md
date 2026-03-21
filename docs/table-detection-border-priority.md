# Table Detection Border Priority

## 背景

現状の `detectTableCandidates(...)` は、次の 2 系統を使って表候補を作る。

- 罫線 seed cell の 4 近傍連結成分
- 値ありまたは罫線あり seed cell の 4 近傍連結成分

後者の fallback は、罫線が弱い表を拾える一方で、地の文や入力フォーム風レイアウトを表として誤検知しやすい。

## 目的

- 罫線を主手掛かりに表を検出する明示モードを追加する
- 非罫線ベース fallback による誤検知を減らす
- 既存既定値は急に壊さず、選択可能な検出モードとして導入する

## モード案

- `balanced`
  - 現状相当
  - border seed と all seed の両方を使う
- `border-priority`
  - 罫線 seed を優先
  - 基本は border seed 由来候補のみ採用
  - all seed fallback は無効、またはかなり限定して使う

`balanced` を既定値として維持し、誤検知が辛い workbook / sheet では `border-priority` を選べるようにする案が第一候補。

## 最小仕様

`border-priority` では、少なくとも次を満たしたい。

- border seed component から得た候補は現状どおり scoring する
- all seed component 由来の fallback 候補は作らない
- `trimTableCandidateBounds(...)` はそのまま使う
- `tableScores` には mode 差が見える理由を残せるとよい

最小実装では、`detectTableCandidates(...)` の後半にある all seed fallback を mode 条件で無効化するだけでも効果がある見込み。

## 差し込み位置

現状の差し込み候補は次。

1. `table-detector.ts`
   - `detectTableCandidates(...)` に mode 引数を追加
2. `core.ts`
   - `sheetMarkdown` へ渡す依存の mode 配線
3. `sheet-markdown.ts`
   - options から table detection mode を受け取る
4. CLI / GUI
   - ユーザーが mode を切り替えられるようにする

## オプション案

- 内部名: `tableDetectionMode`
- 値:
  - `balanced`
  - `border-priority`

既存の `outputMode` / `formattingMode` と同じく、CLI / GUI / summary に出せる形が望ましい。

## テスト方針

Step 2 では、まず誤検知しやすい最小 fixture または最小 unit test を用意する。

- 罫線なしで値が密集しているが、表として扱いたくないケース
- 罫線ありの 2x2 以上で、表として扱いたいケース
- 同一データで `balanced` は候補あり、`border-priority` は候補なしになるケース

## 非目標

- 表検出アルゴリズム全体の刷新
- 罫線以外のスコアリング重み最適化を同時に全部やること
- レイアウト系シートの最終解決

## 次の一歩

1. 誤検知する最小 fixture または unit test 用シートを決める
2. `border-priority` 用の小テストを追加する
3. その後に最小実装を入れる
