# axwrpdf(experimental)
PDF rendering plugin For Susie

## 説明
Windows 8.1 以降のOSが対応しているPDFレンダリング機能を利用して、PDFファイルをアーカイブとして各ページを画像ファイル化するSusie Pluginです。

OS以外には依存ファイルがないので、外部ライブラリ等は不要です。

[こちらの64bit API](http://toro.d.dooo.jp/slplugin.html)に対応したバージョンもあります。

## 制約
- Windows 8.1でも動くはずですが、動作確認できる環境がないので、Windows 10のみに対応しています。
- Susieアプリケーションの終了時に、アプリケーションが異常終了する場合があります。
~~異常終了する原因がわからないため、このPluginは *実験的* という扱いで実用を推奨できません。~~
手元の環境では、ディスプレイドライバーを更新したら発生しなくなりました。
しかし、引き続きこのPluginは *実験的* という扱いです。
- 処理を高速化するために、各ページの画像ファイルのサイズは1バイトとして報告します。(展開するときに実サイズがわかります)
- 全体的に、処理が遅いです。
- 限られた環境でしか確認しておらず、誤動作・動作不良が発生する可能性があります。
- MITライセンスです。利用にあたっては、全て自己責任でお願いします。

## ダウンロード
[https://github.com/udaken/axwrpdf/releases](https://github.com/udaken/axwrpdf/releases)

## 謝辞
こちらの情報が大変参考になりました。

- [Direct2DでPDFを描画するAPIを使ってみた - イグトランスの頭の中](http://dev.activebasic.com/egtra/2015/12/24/853/)
