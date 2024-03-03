Attribute VB_Name = "Mod_QueryTablesToWKB"
Option Explicit
Private Const adSaveCreateNotExist = 1
Private Const adSaveCreateOverWrite = 2
Private Const adWriteChar = 0
Private Const adWriteLine = 1
Private Const adReadLine = -2
Private Const adReadAll = -1
Private Const adTypeBinary = 1
Private Const adTypeText = 2
Private Const adCRLF = -1
Private Const adCR = 13
Private Const adLF = 10
Private Sub Callback_Sample()
    Dim ArrayList As Variant
    Dim WSH As Object: Set WSH = CreateObject("WScript.Shell")
    Dim FilePath As String
    FilePath = WSH.SpecialFolders("Desktop") & "\dummy.csv"
    Dim WKB As Workbook

    Set WKB = QueryTablesToWKB(FilePath, CharSet:="SHIFT-JIS", isGeneralColumn:=Array(3, 4), isSkipColumn:=Array(13, 14, 15))
    If WKB Is Nothing Then
        'NOOP
    Else
        ArrayList = WKB.ActiveSheet.Cells(1, 1).CurrentRegion
        WKB.Close SaveChanges:=False
        Set WKB = Nothing
    End If
    Stop
End Sub
Function QueryTablesToWKB(FilePath As String, _
    Optional CharSet As String = "SHIFT_JIS", _
    Optional isVisibleWKB As Boolean = True, _
    Optional Delimiter As String = ",", _
    Optional LineSeparator As String = vbCrLf, _
    Optional isGeneralColumn As Variant = Empty, _
    Optional isSkipColumn As Variant = Empty _
    ) As Workbook
    '�K�v���W���[��
    'GetArrayDimensionCount
    'isArrayExists
    Dim CharSetType As Object: Set CharSetType = CreateObject("Scripting.Dictionary")
    With CharSetType
        .Add "SHIFT_JIS", 932   'Shift_JIS
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
    If CharSet = "SHIFT-JIS" Then CharSet = Replace(CharSet, "-", "_")
    If Not CharSetType.Exists(CharSet) Then
        '�����R�[�h�w�肪�Ή����Ă��Ȃ�
        GoTo Finally
    ElseIf Not LineSeparatorType.Exists(LineSeparator) Then
        '���s�R�[�h�w�肪�Ή����Ă��Ȃ�
        GoTo Finally
    ElseIf Dir(FilePath, vbNormal) = "" Then
        '�t�@�C�������݂��Ȃ�
        GoTo Finally
    ElseIf Not (GetArrayDimensionCount(isGeneralColumn) = 1 Or IsEmpty(isGeneralColumn)) Then
        'isGeneralColumn�̈�������������
        GoTo Finally
    ElseIf Not (GetArrayDimensionCount(isSkipColumn) = 1 Or IsEmpty(isSkipColumn)) Then
        'isSkipColumn�̈�������������
        GoTo Finally
    Else
        '�������s
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
                        If isSkipFormat Then
                            .Add xlSkipColumn       'SKIP
                        ElseIf isGeneralFormat Then
                            .Add xlGeneralFormat    '����
                        Else
                            .Add xlTextFormat       '������
                        End If
                    Next
                    ReDim ColumnDataTypes(1 To .Count): For i = 1 To .Count: ColumnDataTypes(i) = .Item(i): Next
                End With
            End If
            Exit Do '�P�s�ڂ����J�������]�����Ȃ��̂ŁA��������Do �` Loop����Ȃ�
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
        If Not CharSetType(CharSet) = 1200 Then '1200�͎w�肷��ƃR�P��̂Ŗ��w��
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
    GoTo Finally
ErrorHandler:
    
Finally:
    Application.StatusBar = False
End Function
Private Function isArrayExists(ByVal SourceArray As Variant, _
    ByVal Value As Variant _
    ) As Boolean
    '��1�����z��ɒl�����݂��邩�m�F����
    '�K�{���W���[��
    'GetArrayDimensionCount
    Dim s As Variant
    If IsArray(SourceArray) Then
        If GetArrayDimensionCount(SourceArray) = 1 Then
            For Each s In SourceArray
                If s = Value Then
                    isArrayExists = True
                    Exit For
                End If
            Next
        End If
    End If
End Function
Private Function GetArrayDimensionCount(ByVal ArrayList As Variant) As Integer
    '�������z��̎�������Ԃ�
    Dim i As Long
    Dim DimensionCount As Long
    If IsArray(ArrayList) Then
        On Error Resume Next
        Do While Err.Number = 0
            DimensionCount = DimensionCount + 1
            i = UBound(ArrayList, DimensionCount) '�g�p���Ȃ��l
        Loop
        On Error GoTo 0
        Err.Number = 0
        GetArrayDimensionCount = DimensionCount - 1
    End If
End Function
