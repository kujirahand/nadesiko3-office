# なでしこ3用オフィス(Excel)プラグイン

[日本語プログラミング言語「なでしこ3」](https://nadesi.com)用のプラグインです。

Excelファイルを読み書きできます。Excelファイルを直接書き換えるため、Excel自体のインストールは不要です。
Win/Mac/Linuxで使えます。

## プラグインのインストールの方法

- [Windows版](https://github.com/kujirahand/nadesiko3win32)の場合、実行ファイル「start.exe」を起動した後、「拡張プラグインのインストール」をクリックして「nadesiko3-office」と入力します。
- macOS/Linuxの場合、Node.jsをインストールした環境で、ターミナルで`npm install nadesiko3-office`とコマンドを実行します。

## 利用可能な命令

- [なでしこさんマニュアル > nadesiko3-office](https://nadesi.com/v3/doc/index.php?nadesiko3-office&show)より利用可能な命令の一覧を確認できます。

## 利用例

[リポジトリのsample](https://github.com/kujirahand/nadesiko3-office/tree/master/sample)フォルダに簡単な使い方のサンプルがあります。

```
!「nadesiko3-office」を取り込む。

「Book1.xlsx」のエクセル開く
「B3」のエクセルセル取得して表示。
「B3」に「あいうえお」をエクセル設定。
「B15」に「合計」をエクセル設定。
「C15」に「=SUM(C2:C14)」をエクセル設定。
「test-out.xlsx」へエクセル保存。
```

## 備考

なでしこv1のエクセル命令は、「エクセル起動」が必須でしたが、なでしこv3では不要です。直接『エクセル新規ブック』あるいは『エクセル開く』命令を使って、Excel操作を始めます。


