# xlsx2md

`xlsx2md` は、Excel (`.xlsx`) をローカルで読み込み、地の文と表と画像を Markdown として抽出する Single-file Web App です。

- ブラウザ内でローカルに動作し、サーバ通信を行いません。
- 全シートをシートごとの手作業なしで自動変換します。
- 地の文と表と画像を Markdown として抽出できます。

## 目的

- Excel ブックの情報を、生成AI に渡しやすい Markdown 形式に変換
- 地の文・表・画像を、生成AIが利用しやすい形として抽出
- シートごとの手作業なしで、ブック全体を一括処理
- サーバと通信不要で、ローカル環境だけで処理
- Webブラウザだけで動作し、追加アプリのインストール不要

## 実現方法

- `.xlsx` ファイルをブラウザ内で読み込み
- ファイルの中身を展開して読み取り
- シートや画像などの情報を解析
- 数式は保存済みの値を優先し、必要に応じて数式を解析
- 地の文・表・画像を抽出
- 表を検知して Markdown の表へ変換
- グラフは設定情報のみを抽出
- 図形は元データをテキストとして抽出し、対応できるものは SVG も出力
- ブック全体を Markdown 形式にまとめる

## Screenshots

`xlsx2md` プログラム本体:

![xlsx2md screenshot 0](docs/screenshots/xlsx2md_0.png)

入力の `.xlsx` ブック:

![xlsx2md screenshot 1](docs/screenshots/xlsx2md_1.png)

変換後の Markdown テキスト:

![xlsx2md screenshot 2a](docs/screenshots/xlsx2md_2a.png)

Markdown のプレビュー画面:

![xlsx2md screenshot 2b](docs/screenshots/xlsx2md_2b.png)

## Usage

1. Webブラウザで `xlsx2md.html` を開く
2. `.xlsx` ファイルを選択する
3. 読み込み後、自動で全シートの Markdown が生成される
4. Markdown または ZIP を保存する

## 関連文書

- 上位仕様と設計方針: [docs/xlsx2md-spec.md](./docs/xlsx2md-spec.md)
- 現行実装に即した詳細仕様: [docs/xlsx2md-impl-spec.md](./docs/xlsx2md-impl-spec.md)
- fixture 用 Excel ブックの作成メモ: [tests/fixtures/README.md](./tests/fixtures/README.md)

## LICENSE

### xlsx2md

- Apache License 2.0 のもとで公開しています。
- ライセンス本文は [LICENSE](./LICENSE) を参照してください。

### ClosedXML.Parser

- `xlsx2md` で Excel 数式文法の参考資料として参照しているプロジェクトです。
- `ClosedXML.Parser` は MIT License として公開されています。
- Excel 数式文法の参照資料を公開している `ClosedXML.Parser` プロジェクトに感謝します。
