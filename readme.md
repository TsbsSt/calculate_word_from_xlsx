# calculate_word_from_xlsx
## 概要
所定書式のプロットから推定文字数を計算するプログラムです。


## config.jsonc仕様
### file
ファイルに関する設定です。

#### extension
対応拡張子一覧です。基本的にはxlsx形式を推奨しています。
将来的に他の拡張子も対応するかもしれないので、リスト形式で対応拡張子を設定可能なようにしています。

### sheet
ワークブックのシートに関する設定です。

#### target
編集対象のシートを指定します。
数値を指定すればその番号のシートを、名前を指定すればその名前のシートを編集します。
該当するシートが無い・指定が無い場合は一つ目のシートを編集します。

#### headers
ヘッダーの定義を指定します。

- text      ：プロット用文章です。この文章を元に文字数を計算します。
- details   ：プロットを補足する文章です。この文章を元に文字数を計算します。
- size      ：各プロット行のコマサイズです。文字数計算に使用します。
- words     ：各プロット行の推定文字数です。プログラムの計算結果を記入します。
- total     ：各プロット行の推定文字数の総計です。プログラムの計算結果を記入します。

### words
プロットから文字数を計算するために使用します。
数値で指定してください。

- text_line     ：プロット一行のみなし文字数です。行数とこの数値を掛けて文字数を計算します。
- detail_line   ：プロット補足一行のみなし文字数です。行数とこの数値を掛けて文字数を計算します。


## ライセンス
MIT licenseに準拠します。


# リンク
## twitter
[Twitter](https://twitter.com/2basaSato)

## HP
[甘翼](https://sweetwings.feeling.jp/kanyoku/)

