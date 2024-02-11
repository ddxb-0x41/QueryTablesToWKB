# QueryTablesToWKB

VBAで可変長テキストをQueryTablesを利用して開いてWorkbookオブジェクトを返すFunctionです。

# Installation

fncQueryTablesToWKB.basをVBEでインポートする。

# Usage

```bash
Dim WKB As Workbook
Set WKB = QueryTablesToWKB(FilePath, CharSet:="UTF-8", _
                           Delimiter:=vbTab, _
                           LineSeparator:=vbLf, _
                           isGeneralColumn:=Array(3, 4), _
                           isSkipColumn:=Array(9, 13, 14, 15)
If WKB Is Nothing Then
  Debug.Print "ファイルが存在しないか、文字コードの指定が正しくない。"
Else
  Debug.Print "ファイルを開きました。"
End If
```

### ■必須オプション
FilePath:= ファイルパス

### ■追加オプション
CharSet:= 文字コード：SHIFT-JIS、UTF-8、UTF-16）

Delimiter:= 区切り文字：無指定ならデフォルト値で「,」（カンマ）

LineSeparator:= 改行：無指定ならデフォルト値でvbCrLfを指定。

isGeneralColumn:= 標準は文字列で取り込むので、自動でカラム認識を配列で指定（自動であれば数値や日付で自動認識します）

isSkipColumn:= 取り込み除外するカラムを配列で指定

# License
"ddxb" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).

