# なでしこ3用オフィス(Excel)プラグイン

[日本語プログラミング言語「なでしこ3」](https://nadesi.com)用のプラグインです。

# インストールの方法

- [Windows版](https://github.com/kujirahand/nadesiko3win32)の場合、「npm-install.vbs」というスクリプトがあるので、このスクリプトを実行して、「nadesiko3-office」と入力します。
- macOS/Linuxの場合、Node.jsをインストールした環境で、ターミナルで`npm install nadesiko3-office`とコマンドを実行します。

# 利用例

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




