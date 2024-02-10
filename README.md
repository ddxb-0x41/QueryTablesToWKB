# QueryTablesToWKB（）

VBAで可変長テキストをQueryTablesを利用してWorkbookオブジェクトを返しす。

# Installation

fncQueryTablesToWKB.basをVBEでインポートする。

# Usage

```bash
Dim WKB As Workbook
Set WKB = QueryTablesToWKB(FilePath, CharSet:="UTF-8", isGeneralColumn:=Array(3, 4), isSkipColumn:=Array(9, 13, 14, 15)
```

※ファイルが読み込めない場合などは、WKB は Nothing が返ってきます。

### ■必須オプション
FilePath:= ファイルパス

CharSet:= 文字コード：SHIFT-JIS、UTF-8、UTF-16）

### ■追加オプション
Delimiter:= 区切り文字：デフォルトは「,」（カンマ）

LineSeparator:= 改行：デフォルトはvbCrLf、指定してもvbLfくらい

isGeneralColumn:= 標準は文字列で取り込むので、自動でカラム認識を配列で指定

isSkipColumn:= 取り込み除外するカラムを配列で指定

# License
"ddxb" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).

