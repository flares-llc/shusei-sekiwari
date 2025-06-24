# shusei-sekiwari

## OSS としての公開について

本システムはオープンソースソフトウェア（OSS）として MIT ライセンスで公開しています。

- ライセンス: [MIT License](./LICENSE)
- 貢献: [CONTRIBUTING.md](./CONTRIBUTING.md) をご参照ください。
- 行動規範: [CODE_OF_CONDUCT.md](./CODE_OF_CONDUCT.md)

## 概要

本システムは、Google スプレッドシートと Google Apps Script（GAS）を利用して動作するシフト・集計管理システムです。

`sample/sample.xlsx` は参考用の Excel ファイルであり、そのままでは動作しません。Excel ファイルの内容や各セルの状態を Google スプレッドシートに再現し、GAS を設定することでシステムが利用可能となります。

## 初期設定手順

1. **Google スプレッドシートの作成**

   - Google ドライブ上で新しいスプレッドシートを作成します。

2. **Excel 内容の反映**

   - `sample/sample.xlsx` を参考に、シート構成や各セルの内容・数式・書式を Google スプレッドシートに手動で再現してください。

3. **GAS（Google Apps Script）の設置**

   - スプレッドシートのメニューから「拡張機能」→「Apps Script」を開きます。
   - `gas/main.js` の内容をコピーし、Apps Script エディタに貼り付けて保存します。

4. **トリガーの設定**
   - Apps Script エディタで「トリガー」メニューを開き、必要な関数に対して時間主導型や編集時など、用途に応じたトリガーを設定してください。

---

ご不明点があれば、`sample/sample.xlsx` や `gas/main.js` をご参照ください。
