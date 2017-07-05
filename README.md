# axwrpdf(experimental)
PDF rendering plugin For Susie

## 説明
Windows 8.1 以降のOSが対応しているPDFレンダリング機能を利用して、PDFファイルをアーカイブとして各ページを画像ファイル化するSusie Pluginです。

OS以外には依存ファイルがないので、外部ライブラリ等は不要です。

[こちらの64bit API](http://toro.d.dooo.jp/slplugin.html)に対応したバージョンもあります。

## 制約
- Windows 8.1でも動くはずですが、動作確認できる環境がないので、Windows 10のみに対応しています。
- Susieアプリケーションの終了時に、アプリケーションが異常終了する場合があります。
異常終了する原因がわからないため、このPluginは *実験的* という扱いで実用を推奨できません。
- 処理を高速化するために、各ページの画像ファイルのサイズは1バイトとして報告します。(展開するときに実サイズがわかります)
この仕様により、Massigra,Leeyesは表示できるようですが、NeeViewではエラーになってしまいます。

## 謝辞
こちらの情報が大変参考になりました。

- http://dev.activebasic.com/egtra/2015/12/24/853/



