Attribute VB_Name = "fncQueryTablesToWKB"
Option Explicit
Private Sub Callback_Sample()
    Dim WSH As Object: Set WSH = CreateObject("WScript.Shell")
    Dim FilePath As String
    FilePath = WSH.SpecialFolders("Desktop") & "\dummy.csv"
    Dim WKB As Workbook
    Set WKB = QueryTablesToWKB(FilePath, CharSet:="UTF-8", isGeneralColumn:=Array(3, 4), isSkipColumn:=Array(9, 13, 14, 15))
End Sub
Function QueryTablesToWKB(ByVal FilePath As String, _
              Optional ByVal CharSet As String = "SHIFT-JIS", _
              Optional ByVal Delimiter As String = ",", _
              Optional ByVal LineSeparator As String = vbCrLf, _
              Optional ByVal isGeneralColumn As Variant, _
              Optional ByVal isSkipColumn As Variant) As Workbook
    Const adReadLine = -2
    Dim CharSetType As Object: Set CharSetType = CreateObject("Scripting.Dictionary")
    With CharSetType
        .Add "SHIFT-JIS", 932
        .Add "UTF-8", 65001
        .Add "UTF-16", 1200
        .Add "UNICODE", 1200
    End With
    CharSet = UCase(CharSet)
    If Not CharSetType.Exists(CharSet) Then
        QueryTablesToWKB = Nothing
        GoTo Finally
    End If
    Dim sh As Worksheet
    Dim ReadTextLine As Variant
    Dim ColumnDataTypes As Variant, i As Long
    Dim isGeneralFormat As Boolean
    Dim isSkipFormat As Boolean
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 2
        .CharSet = CharSet
        .LineSeparator = Asc(LineSeparator)
        .LoadFromFile FilePath
        Do Until .EOS
            ReadTextLine = Split(.ReadText(adReadLine), Delimiter)
            If UBound(ReadTextLine) <= 0 Then
                QueryTablesToWKB = Nothing
                GoTo Finally
            Else
                '
            End If
            Exit Do
        Loop
        .Close
    End With
    With New Collection
        For i = LBound(ReadTextLine) To UBound(ReadTextLine)
            isGeneralFormat = False
            isSkipFormat = False
            If IsArray(isGeneralColumn) Then
                isGeneralFormat = isArrayExists(isGeneralColumn, i + 1)
            End If
            If IsArray(isSkipColumn) Then
                isSkipFormat = isArrayExists(isSkipColumn, i + 1)
            End If
            If isGeneralFormat Then
                .Add xlGeneralFormat    '自動
            ElseIf isSkipFormat Then
                .Add xlSkipColumn       'スキップカラム
            Else
                .Add xlTextFormat       '文字列
            End If
        Next
        ReDim ColumnDataTypes(1 To .Count): For i = 1 To .Count: ColumnDataTypes(i) = .Item(i): Next
    End With
    Application.StatusBar = "[QueryTables読み込み]" & Dir(FilePath)
    Application.ScreenUpdating = False
    Set QueryTablesToWKB = Workbooks.Add
    Set sh = QueryTablesToWKB.ActiveSheet
    With sh.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=sh.Range("A1"))
        .TextFileColumnDataTypes = ColumnDataTypes
        .TextFilePlatform = CharSetType(CharSet)
        .AdjustColumnWidth = False
        .TextFileOtherDelimiter = Delimiter
        .Refresh BackgroundQuery:=False
        .Delete
    End With
Finally:
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Function
Private Function isArrayExists(ByVal ArrayList As Variant, ByVal CheckValue As Variant) As Boolean
    Dim s As Variant
    For Each s In ArrayList
        If s = CheckValue Then
            isArrayExists = True
            Exit For
        End If
    Next
End Function
