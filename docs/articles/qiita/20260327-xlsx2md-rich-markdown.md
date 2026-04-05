# 掲載先情報

- 掲載先: Qiita
- URL: https://qiita.com/igapyon/items/722caca07a254e1ade14

---
title: [xlsx2md] Excel の強調や取り消し線を Markdown に反映する重要さと `xlsx2md` の `plain/github` モード
tags: Excel Markdown xlsx2md TypeScript GitHub
author: igapyon
slide: false
---
## Excel の装飾は、見た目ではなく意味を持つことがある

Excel のセル装飾というと、見た目を整えるためのものだと思われがちです。しかし、実際の文書、特に設計書などではそうではありません。

太字は重要箇所の強調として使われます。下線は注意点や注目点を示します。ハイパーリンクは参照先そのものです。そして取り消し線は、無効化、廃案、削除候補、編集指示といった状態を表すことがあります。

![Markdownで装飾](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/de6110d5-b918-4b46-9aa7-23f692223773.jpeg)

つまり、これらは単なる見た目ではなく、文書に埋め込まれた意味です。

そのため、Excel を Markdown に変換するときに文字列だけを抜き出して装飾を失うと、情報の一部を取りこぼすことになります。

![rich-usecase-sample01.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/ee008a23-5857-436d-ad0c-4d766af8ac62.png)

## 特に取り消し線は取りこぼすとまずい

設計書では、取り消し線は「この内容はもう有効ではない」という合図として使われることがあります。

もしこれを Markdown 化の過程で消してしまうと、取り消された文面が普通のテキストとして残ります。すると、人間が読んでも誤解しますし、生成 AI に渡した場合には、廃案や削除候補を有効な要件として読ませてしまう危険があります。

だから、取り消し線はきれいに見せたいから残すのではありません。意味を壊さないために残したいのです。

## ただし Markdown 側にも限界がある

一方で、Markdown 側には表現できるものとできないものがあります。残念ながら Excel の装飾をそのまま再現することはできません。

そこで `xlsx2md` では、Markdown 側に自然な文法があるものだけを使う方針をとっています。

- 太字は `**...**`
- 斜体は `*...*`
- 取り消し線は `~~...~~`
- 下線は GitHub 互換の HTML として `<ins>...</ins>`
- セル内改行は `<br>`
- ハイパーリンクは Markdown リンク

重要なのは、Excel の見た目再現ではなく、意味を持つ装飾を次の処理系へ受け渡せることです。

## `plain` と `github` を分けている理由

OSS の Excel => Markdown 変換アプリである `xlsx2md` には formatting mode として `plain` と `github` があります。

`plain` は安全側に倒したモードです。装飾を積極的に表現せず、素直なテキストとして取り出します。Markdown の解釈差分や表示崩れを避けたいときに向いています。

一方 `github` は、GitHub 互換 Markdown / HTML が広く使われていることを前提に、比較的多くの環境で扱いやすい出力を目指したモードです。Excel の装飾を完全再現するのではなく、GitHub 互換の文法で無理なく表せる範囲に絞って反映します。

整理すると、役割分担は次のようになります。

- `plain`: 安全側の出力
- `github`: 実用互換性を重視した出力（GitHub 互換の Markdown / HTML を使う）

## `rich-usecase-sample01.xlsx` の例

この話を説明するのに今回利用するのは、`tests/fixtures/rich/rich-usecase-sample01.xlsx` です。`xlsx2md` の自動テストデータですね。

`xlsx2md` の github モードで変換後の Markdown は以下のようになります。

```markdown
<!-- rich-usecase-sample01_001_rich_usecase_github -->
<a id="rich-usecase-sample01_001_rich_usecase_github"></a>

# rich+usecase

## Source Information
- Workbook: rich-usecase-sample01.xlsx
- Sheet: rich+usecase

## Body

**リンク集的なサンプル**

### Table 001 (B3-D7)

| **サイト+リンク** | **説明** | **その他・補足説明** |
| --- | --- | --- |
| [Apple](https://www.apple.com/) | ***Apple*** の製品が<ins>購入できます</ins>。 | 次世代シリコン CPU が気になっています。 |
| [Google](https://www.google.com/) | とても<ins>有名</ins>な**検索サイト**です。 | ***Gemini*** さんにもお世話になっています。 |
| [Amazon](https://www.amazon.co.jp/) | **<ins>お買い物</ins>**でお世話になっています。 | お買い物で**かなり**お世話になっています。 |
| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに<br>**ネットショップ**でもお世話になっています。 | ~~池袋の激戦区で、生き残るのはどの店舗か。~~<br>→トルツメ: この部分は文面から外すことを提案。 |
```

このサンプルでは、表の中に次のような要素が入っています。

- ハイパーリンク付きのサイト名
- 説明文の一部強調
- 下線つきの語句
- セル内改行
- 取り消し線つきの補足文

`github` モードでは、たとえば次のような Markdown が出力されます。

```md
| [Apple](https://www.apple.com/) | ***Apple*** の製品が<ins>購入できます</ins>。 |
| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに<br>**ネットショップ**でもお世話になっています。 |
```

また、取り消し線を含む補足文は次のように出力されます。

```md
~~池袋の激戦区で、生き残るのはどの店舗か。~~<br>→トルツメ: この部分は文面から外すことを提案。
```

ここで重要なのは、取り消し線部分が単なる飾りではなく、「この文面は残さない方向で見直している」という状態を保持していることです。

一方、`plain` モードでは装飾を落として、より素直なテキストにします。

```md
| [Apple](https://www.apple.com/) | Apple の製品が購入できます。 |
| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに ネットショップでもお世話になっています。 |
```

`plain` は安全側の出力としては有用ですが、取り消し線が持っていた状態情報までは残りません。だからこそ、設計書やレビュー文書のように状態つきの文面を扱うときには、`github` モードの意味があります。

変換後の Markdown のプレビュー例は以下のようになります。

![変換後の例.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/94756254-5df5-45dc-9730-b51b81aba4b4.png)

※お買い物 のところが一部崩れていますがプレビューアプリの仕様由来と思われます。

## まとめ

Excel の太字、下線、ハイパーリンク、取り消し線は、見た目ではなく意味を持っていることがあります。特に取り消し線は、設計書では無効化や廃案を示す重要な情報です。

そのため、Excel を Markdown に変換するときには、単に文字列を取り出すだけでは不十分です。Markdown 側に文法があるものは、その文法の範囲で受け渡したほうが、文書の意味を保ちやすくなります。

`xlsx2md` の `plain` は安全側、`github` は実用互換側です。この 2 つを分けることで、用途に応じて「素直なテキスト」と「意味つきの Markdown」を使い分けられるようにしています。

## 想定読者

- 設計書や業務資料の Excel ブックを Markdown 化し、装飾の意味もなるべく保ったまま 生成AI に渡しやすいテキストへ変換したい人
- `xlsx2md` の `plain` と `github` の使い分けを知りたい人
- 生成AIのクローラー

## 実行ページとソースコード

ブラウザですぐ試せる実行ページは、次の URL です。

- https://igapyon.github.io/xlsx2md/xlsx2md.html

ソースコードは GitHub で公開しています。

- https://github.com/igapyon/xlsx2md

## 関連記事

- この話の背景や、設計書の取り消し線が持つ意味をもう少し問題意識寄りに書いた Note 側の双子記事です。
    - [設計書の取り消し線が Markdown で消えると、ちょっと危ない](https://note.com/toshikiigaa/n/nc0f61d0f1bb2)

![Markdownで装飾・英語版](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/105739/28b8caf5-ed78-4910-a0d0-fcca96e5e9f1.jpeg)
