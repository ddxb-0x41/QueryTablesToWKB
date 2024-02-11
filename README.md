# QueryTablesToWKB

VBAで可変長テキストをQueryTablesを利用してWorkbookオブジェクトを返す。

# Installation

fncQueryTablesToWKB.basをVBEでインポートする。

# Usage

```bash
Dim WKB As Workbook
Set WKB = QueryTablesToWKB(FilePath, CharSet:="UTF-8", _
                           isGeneralColumn:=Array(3, 4), _
                           isSkipColumn:=Array(9, 13, 14, 15)
```

※ファイルが読み込めない場合などは、WKB は Nothing が返えす。

### ■必須オプション
FilePath:= ファイルパス
### ■追加オプション
CharSet:= 文字コード：SHIFT-JIS、UTF-8、UTF-16）

Delimiter:= 区切り文字：デフォルトは「,」（カンマ）

LineSeparator:= 改行：デフォルトはvbCrLf、指定してもvbLfくらい

isGeneralColumn:= 標準は文字列で取り込むので、自動でカラム認識を配列で指定

isSkipColumn:= 取り込み除外するカラムを配列で指定

# License
"ddxb" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).

