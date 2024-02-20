Attribute VB_Name = "fncQueryTablesToWKB"
Option Explicit
Private Const adReadLine = -2
Private Const adTypeText = 2
Private Const adCRLF = -1
Private Const adCR = 13
Private Const adLF = 10
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Callback_Sample()
    Dim WSH As Object: Set WSH = CreateObject("WScript.Shell")
    Dim FilePath As String
    FilePath = WSH.SpecialFolders("Desktop") & "\TEST.txt"
    Dim WKB As Workbook
    Set WKB = QueryTablesToWKB(FilePath, CharSet:="UTF-8", isGeneralColumn:=Array(3, 4), isSkipColumn:=Array(13, 14, 15))
    If WKB Is Nothing Then
        'NOOP
    Else
        WKB.Close SaveChanges:=False
    End If
End Sub
Function QueryTablesToWKB(ByVal FilePath As String, _
    Optional ByVal CharSet As String = "SHIFT-JIS", _
    Optional ByVal isVisibleWKB As Boolean = True, _
    Optional ByVal Delimiter As String = ",", _
    Optional ByVal LineSeparator As String = vbCrLf, _
    Optional ByVal isGeneralColumn As Variant, _
    Optional ByVal isSkipColumn As Variant) As Workbook
    Dim CharSetType As Object: Set CharSetType = CreateObject("Scripting.Dictionary")
    With CharSetType
        .Add "SHIFT-JIS", 932
        .Add "UTF-8", 65001     'UTF-8 or UTF-8BOM
        .Add "UTF-16", 1200     'UTF-16LEBOM
    End With
    Dim LineSeparatorType As Object: Set LineSeparatorType = CreateObject("Scripting.Dictionary")
    With LineSeparatorType
        .Add vbCrLf, adCRLF
        .Add vbLf, adLF
        .Add vbCr, adCR
    End With
    CharSet = UCase(CharSet)
    If Not CharSetType.Exists(CharSet) Then
        GoTo Finally '文字コード指定が対応していない
    ElseIf Not LineSeparatorType.Exists(LineSeparator) Then
        GoTo Finally '改行コード指定が対応していない
    ElseIf Dir(FilePath, vbNormal) = "" Then
        GoTo Finally 'ファイルが存在しない
    ElseIf Not (IsArray(isGeneralColumn) Or TypeName(isGeneralColumn) = "Error") Then
        GoTo Finally 'isGeneralColumnの引数がおかしい
    ElseIf Not (IsArray(isSkipColumn) Or TypeName(isSkipColumn) = "Error") Then
        GoTo Finally 'isSkipColumnの引数がおかしい
    Else
        'NOOP
    End If
    Dim sh As Worksheet
    Dim ReadTextLine As Variant
    Dim ColumnDataTypes As Variant, i As Long
    Dim isGeneralFormat As Boolean
    Dim isSkipFormat As Boolean
    With CreateObject("ADODB.Stream")
        .Open
        .Type = adTypeText
        .CharSet = CharSet
        .LineSeparator = LineSeparatorType(LineSeparator)
        .LoadFromFile FilePath
        Do Until .EOS
            ReadTextLine = Split(.ReadText(adReadLine), Delimiter)
            If UBound(ReadTextLine) <= 0 Then
                QueryTablesToWKB = Nothing
                GoTo Finally
            Else
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
                            .Add xlSkipColumn       'SKIP
                        Else
                            .Add xlTextFormat       '文字列
                        End If
                    Next
                    ReDim ColumnDataTypes(1 To .Count): For i = 1 To .Count: ColumnDataTypes(i) = .Item(i): Next
                End With
            End If
            Exit Do '１行目しかカラム数評価しないので、そもそもDo ～ Loopいらない
        Loop
        .Close
    End With
    Application.StatusBar = "[Loading...]" & Dir(FilePath)
    Set QueryTablesToWKB = Workbooks.Add
    If Not isVisibleWKB Then
        Application.Windows(QueryTablesToWKB.Name).Visible = isVisibleWKB
    End If
    Set sh = QueryTablesToWKB.ActiveSheet
    With sh.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=sh.Cells(1, 1))
        .TextFileColumnDataTypes = ColumnDataTypes
        If Not CharSetType(CharSet) = 1200 Then '1200は指定するとコケることがあるので無指定
            .TextFilePlatform = CharSetType(CharSet)
        End If
        .AdjustColumnWidth = False
        If Delimiter = "," Then
            .TextFileCommaDelimiter = True
        ElseIf Delimiter = ";" Then
            .TextFileSemicolonDelimiter = True
        Else
            .TextFileOtherDelimiter = Delimiter
        End If
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    Application.StatusBar = False
    GoTo Finally
ErrorHandler:
    
Finally:
    Application.StatusBar = False
End Function
Private Function isArrayExists(ArrayList As Variant, CheckValue As Variant) As Boolean
    Dim s As Variant
    If IsArray(ArrayList) Then
        For Each s In ArrayList
            If s = CheckValue Then
                isArrayExists = True
                Exit For
            End If
        Next
    End If
End Function


