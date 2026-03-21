# rich text / Markdown rendering メモ

## 1. 文書の位置づけ

この文書は、`xlsx2md` における rich text、セル内改行、Markdown 記号の扱いを見える化するための設計メモである。

目的は次の通りである。

- 現在の実装がどこで何をしているかを整理する
- `plain` / `github` の責務差を明確にする
- 将来の Markdown escape 対応や renderer 分離の判断材料にする

本書は理想設計を先に固定するための文書ではなく、現行実装と近い将来の拡張点を切り分けるための補助文書である。

## 2. 背景

`xlsx2md` は、もともと Excel の表示値や構造を Markdown に落とすことを主目的としていた。

今回追加された rich text 関連の要件では、少なくとも次を扱う必要がある。

- Excel のセル全体装飾
  - `bold`
  - `italic`
  - `strike`
  - `underline`
- shared string / inline string の部分装飾
- セル内改行
- Markdown 記号を含む生文字

これらはすべて「文字列をどう Markdown にレンダリングするか」という共通の問題に属する。

## 3. 現在の実装位置

現行実装は、完全な Markdown parser / AST / renderer を新設したものではない。

実際には、既存の `xlsx -> internal model -> markdown` の流れに対して、rich text 用の中間情報を追加し、出力時に分岐している。

主な責務は次の通りである。

- `src/xlsx2md/ts/shared-strings.ts`
  - `sharedStrings.xml` の `si / r / rPr / t` を読む
  - run 単位の `bold / italic / strike / underline` を抽出する
- `src/xlsx2md/ts/styles-parser.ts`
  - `styles.xml` の `font / cellXfs` を読む
  - セル全体 style の `bold / italic / strike / underline` を抽出する
- `src/xlsx2md/ts/worksheet-parser.ts`
  - shared string / inline string / cell style を統合し、セルごとの `richTextRuns` と `textStyle` を持たせる
- `src/xlsx2md/ts/markdown-escape.ts`
  - 生文字としての Markdown 記号、`<`, `>`, `&` を escape する
- `src/xlsx2md/ts/rich-text-parser.ts`
  - `outputValue / textStyle / richTextRuns` をもとに `text / lineBreak / styledText` token 列へ変換する
- `src/xlsx2md/ts/rich-text-plain-formatter.ts`
  - `plain` モード固有の text 化を担当する
- `src/xlsx2md/ts/rich-text-github-formatter.ts`
  - `github` モード固有の wrapper 適用と `<br>` 化を担当する
- `src/xlsx2md/ts/rich-text-renderer.ts`
  - token 列を `plain` / `github` の文字列へ描画する
  - `plain` は素朴な text 化、`github` は formatter 呼び出しを担当する
- `src/xlsx2md/ts/markdown-table-escape.ts`
  - Markdown table 専用のセル escape を担当する
- `src/xlsx2md/ts/sheet-markdown.ts`
  - シート構造の都合と Markdown section/table 組み立てを担当する

したがって、現状は「小さな parser / renderer パイプラインを持つ lightweight 実装」であり、本格的な Markdown AST renderer にはまだ至っていない。

## 4. 現在の中間表現

現行セルモデルは、少なくとも次の rich text 関連情報を持つ。

- `textStyle`
  - セル全体 style としての `bold / italic / strike / underline`
- `richTextRuns`
  - 部分装飾を持つ run 配列
  - 各 run は少なくとも次を持つ
    - `text`
    - `bold`
    - `italic`
    - `strike`
    - `underline`

考え方としては次の通りである。

- セル全体 style は fallback として使う
- run 情報がある場合は run 単位のレンダリングを優先する
- `rich-text-parser.ts` 内では、次の token を使う
  - `text`
  - `lineBreak`
  - `styledText`
- `styledText` は将来の拡張余地のため、単一文字列ではなく `parts` を持つ
- `styledText.parts` の各 part は次を持つ
  - `kind`
    - `text`
    - `escaped`
  - `text`
    - Markdown へ実際に出力する文字列
  - `rawText`
    - escape 前の元文字

## 5. モードごとの責務

### 5.1 `plain`

`plain` は、見た目依存の装飾を落とし、素朴なテキストへ寄せるモードである。

現時点では次の方針を取る。

- `bold / italic / strike / underline` を出力しない
- セル内改行は空白へ正規化する
- 生文字はできるだけそのまま出す

`plain` は「最小限の情報共有用」であり、GitHub 固有の表現を持ち込まない。

### 5.2 `github`

`github` は、GitHub 上で比較的安定して解釈できる表現へ寄せるモードである。

現時点では次の方針を取る。

- `bold` -> `**...**`
- `italic` -> `*...*`
- `strike` -> `~~...~~`
- `underline` -> `<ins>...</ins>`
- セル内改行 -> `<br>`

`github` は「Markdown 純正だけ」ではなく、「GitHub 上で安定する Markdown + 一部 HTML」のモードとして扱う。

## 6. 現在のレンダリング順序

現行の考え方は概ね次の順序である。

1. Excel 由来の文字列を受け取る
2. `markdown-escape.ts` で生文字を escape する
3. `rich-text-parser.ts` で `plain` または `github` の token 列を作る
4. `rich-text-renderer.ts` が `plain` / `github` の描画経路を選ぶ
5. `plain` の場合は `rich-text-plain-formatter.ts` が text 化する
6. `github` の場合は `rich-text-github-formatter.ts` が装飾や `<br>` を反映する
7. 表セルとして出力する場合は `markdown-table-escape.ts` で表用 escape を別途適用する

この構成により、表セルでは少なくとも `|` の崩れは抑えられる。

一方で、Markdown 記号そのものの escape は、現時点では限定的である。

## 7. いま問題になっていること

rich text と `<br>` を扱い始めると、次の問題が表面化する。

### 7.1 生文字と生成記号の混在

Excel のセルに次のような文字が入る場合がある。

- `*`
- `_`
- `~~`
- `#`
- `-`
- `1.`
- `[label](url)`
- `` `code` ``
- `<tag>`
- `|`

ここで難しいのは、`xlsx2md` 自身も `**`, `*`, `~~`, `<ins>`, `<br>` を生成していることである。

つまり、次の 2 種類を分離して扱う必要がある。

- ユーザーが元から入力した文字
- renderer が意味を持たせるために生成した記号

### 7.2 表セル内の安全性

表セルでは、通常の段落よりもさらに安全性が必要になる。

特に重要なのは次である。

- `|` の列崩れ
- 改行の扱い
- 装飾記号と生文字の干渉

### 7.3 現状は renderer が少し太っている

現在の `sheet-markdown.ts` は、次を同時に扱っている。

- 値モードの切替
- formatting mode の切替
- rich text run のレンダリング
- 改行の変換
- 表セル入力に向けた文字列整形の前処理

これは短期的には動くが、Markdown escape を本格対応するには責務分離が必要になる。

## 8. 将来の分離方針

将来的には、少なくとも次の 3 層へ寄せるのが自然である。

### 8.1 生文字 normalizer / escape

責務:

- Excel 由来の生文字を受け取り、unsafe 制御文字や改行表現を整える
- Markdown 上で危険な生文字を escape する
- ただし装飾記号はまだ付けない

### 8.2 rich text parser

責務:

- `textStyle` または `richTextRuns` を受け取り、`plain` / `github` の方針に応じて token 列へ変換する
- `lineBreak` と `styledText` の境界をここで確定する

### 8.3 rich text renderer

責務:

- parser が返した token 列を、どの描画経路へ流すか決める
- `plain` と `github` の責務境界を保つ

### 8.4 plain formatter

責務:

- `plain` モード固有の text 化
- `styledText.parts` を style を落として束ねる

### 8.5 github formatter

責務:

- `github` モード固有の wrapper 適用
- `<br>` 化
- `styledText.parts` を文字列へ束ねて style をかける

### 8.6 Markdown table escape

責務:

- 表セル専用 escape
- renderer が生成した構文は壊さず、表セルとして危険な部分だけを保護する

## 9. 将来的な中間表現の候補

今後 escape を強化する場合、単なる文字列よりも token 列に寄せた方が安全である。

例えば次のような表現である。

```text
TextToken("abc")
StyledToken(style=bold, parts=[Part(kind=text, rawText="d", text="d"), ...])
LineBreakToken()
TextToken("[label](x)")
```

このような中間表現を持てば、

- 生文字 escape
- 装飾付与
- `<br>` 変換
- 表セル用 escape

を分離しやすくなる。

ただし現時点では、そこまで進めなくても `richTextRuns` を正本として扱うだけで多くのケースは整理できる。

## 10. 当面の実務方針

現時点では、次の方針で十分である。

- `plain` は素朴な文字列共有用として維持する
- `github` は GitHub 上の可読性を優先する
- rich text fixture と markdown escape fixture の両方で `plain / github` を回帰テストする
- Markdown escape を本格対応する前に、renderer の責務境界を整理する

## 11. テスト観点

rich text / Markdown rendering で今後も固定したい観点は次の通りである。

- セル全体 `bold / italic / strike / underline`
- 部分装飾 run
- 複合装飾
- セル内改行の `<br>` 化
- 表セル内での崩れ
- Markdown 記号を含む生文字
- `plain` と `github` の差分

これらは、少なくとも次の fixture で回帰テストを持つ。

- `tests/fixtures/rich/rich-text-github-sample01.xlsx`
- `tests/fixtures/rich/rich-markdown-escape-sample01.xlsx`

## 12. 結論

今回の実装は「腕力だけの場当たり対応」ではなく、小さい escape / parser / plain-formatter / github-formatter / renderer / table-escape を段階的に分離した lightweight pipeline である。

一方で、Markdown escape まで本格的に扱うなら、次の段階では parser / token / renderer に近い整理が必要になる。

したがって、当面の判断は次の通りである。

- いまは `richTextRuns + textStyle` を軸に進める
- `plain` と `github` の責務を混ぜない
- escape 本格対応の前に、rendering pipeline を見える化しておく

この文書は、その見える化のための起点である。
