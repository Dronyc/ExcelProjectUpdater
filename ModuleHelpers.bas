Attribute VB_Name = "ModuleHelpers"
'===============================================================
' ВСПОМОГАТЕЛЬНЫЙ МОДУЛЬ (ModuleHelpers)
' Общие функции, используемые во всех модулях
'===============================================================
Option Explicit

' Windows API для управления окнами (TopMost)
Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST As LongPtr = -1
Private Const HWND_NOTOPMOST As LongPtr = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOACTIVATE As Long = &H10

Public mLastPct As Long
Public gStep As String

'===============================================================
' РАБОТА С ФАЙЛАМИ И ПАПКАМИ
'===============================================================
Public Function ConfirmOverwrite(ByVal filePath As String) As Boolean
    If Len(Dir(filePath)) = 0 Then
        ConfirmOverwrite = True
        Exit Function
    End If
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Файл уже существует:" & vbCrLf & filePath & vbCrLf & vbCrLf & "Перезаписать?", vbYesNo + vbQuestion)
    ConfirmOverwrite = (answer = vbYes)
End Function

Public Function DeleteFileIfExists(ByVal filePath As String) As Boolean
    DeleteFileIfExists = True
    If Len(Dir(filePath)) = 0 Then Exit Function
    On Error Resume Next
    Kill filePath
    If Err.Number <> 0 Then
        DeleteFileIfExists = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Public Function CreateTimestampFolder(ByVal baseFolder As String) As String
    Dim stamp As String, folderPath As String, suffix As Long
    stamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    folderPath = baseFolder & "\" & stamp
    Do While Len(Dir(folderPath, vbDirectory)) > 0
        suffix = suffix + 1
        folderPath = baseFolder & "\" & stamp & "_" & suffix
    Loop
    MkDir folderPath
    CreateTimestampFolder = folderPath
End Function

Public Function GetUniqueFilePath(ByVal folderPath As String, ByVal fileName As String) As String
    Dim filePath As String, baseName As String, extName As String, suffix As Long, dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left(fileName, dotPos - 1)
        extName = Mid(fileName, dotPos)
    Else
        baseName = fileName
        extName = ""
    End If
    filePath = folderPath & "\" & fileName
    Do While Len(Dir(filePath)) > 0
        suffix = suffix + 1
        filePath = folderPath & "\" & baseName & "_" & suffix & extName
    Loop
    GetUniqueFilePath = filePath
End Function

Public Function IsWorkbookOpenInApp(ByVal fileName As String) As Boolean
    Dim wb As Workbook
    For Each wb In Workbooks
        If StrComp(wb.name, fileName, vbTextCompare) = 0 Then
            IsWorkbookOpenInApp = True
            Exit Function
        End If
    Next wb
    IsWorkbookOpenInApp = False
End Function

'===============================================================
' РАБОТА С ЛИСТАМИ И ДАННЫМИ
'===============================================================
Public Function GetLastRow(ByVal ws As Worksheet) As Long
    Dim cell As Range
    Set cell = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If cell Is Nothing Then GetLastRow = 0 Else GetLastRow = cell.row
End Function

Public Function GetLastColumn(ByVal ws As Worksheet) As Long
    Dim cell As Range
    Set cell = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If cell Is Nothing Then GetLastColumn = 0 Else GetLastColumn = cell.Column
End Function

Public Function GetCleanText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        GetCleanText = ""
    Else
        GetCleanText = Trim(CStr(v))
    End If
End Function

Public Function GetCodeText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        GetCodeText = ""
        Exit Function
    End If
    If VarType(v) = vbString Then
        GetCodeText = Trim(CStr(v))
    ElseIf IsNumeric(v) Then
        GetCodeText = Format(v, "0")
    Else
        GetCodeText = Trim(CStr(v))
    End If
End Function

Public Function SafeText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then SafeText = "" Else SafeText = Trim(CStr(v))
End Function

Public Function IsEmptySourceRow(ByVal srcValues As Variant, ByVal r As Long, ByVal colIdx As Variant) As Boolean
    Dim k As Long
    For k = 1 To 12
        If Len(GetCleanText(srcValues(r, colIdx(k)))) > 0 Then
            IsEmptySourceRow = False
            Exit Function
        End If
    Next k
    IsEmptySourceRow = True
End Function

'===============================================================
' РАБОТА С ДАТАМИ
'===============================================================
Public Function ParseRuDateTime(ByVal v As Variant) As Variant
    ParseRuDateTime = Empty
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then
        Dim d As Date: d = CDate(v)
        If Year(d) = 1 And Month(d) = 1 And Day(d) = 1 And Hour(d) = 4 And Minute(d) = 0 And Second(d) = 0 Then
            ParseRuDateTime = "EXCLUDED_DATE"
        Else
            ParseRuDateTime = d
        End If
        Exit Function
    End If
    Dim s As String: s = GetCleanText(v)
    If Len(s) = 0 Then Exit Function
    Dim parts As Variant: parts = Split(s, " ")
    If UBound(parts) < 0 Then Exit Function
    Dim datePart As String, timePart As String
    datePart = parts(0)
    If UBound(parts) >= 1 Then timePart = parts(1) Else timePart = "00:00:00"
    Dim dateParts As Variant, timeParts As Variant
    dateParts = Split(datePart, ".")
    If UBound(dateParts) <> 2 Then Exit Function
    timeParts = Split(timePart, ":")
    If UBound(timeParts) < 1 Then Exit Function
    Dim dayN As Long, monthN As Long, yearN As Long
    Dim hourN As Long, minuteN As Long, secondN As Long
    On Error Resume Next
    dayN = CLng(dateParts(0)): monthN = CLng(dateParts(1)): yearN = CLng(dateParts(2))
    hourN = CLng(timeParts(0)): minuteN = CLng(timeParts(1))
    If UBound(timeParts) >= 2 Then secondN = CLng(timeParts(2)) Else secondN = 0
    If Err.Number <> 0 Then ParseRuDateTime = Empty: Exit Function
    On Error GoTo 0
    If yearN = 1 And monthN = 1 And dayN = 1 And hourN = 4 And minuteN = 0 And secondN = 0 Then
        ParseRuDateTime = "EXCLUDED_DATE"
        Exit Function
    End If
    If yearN < 1900 Then Exit Function
    Dim resultDate As Date
    On Error Resume Next
    resultDate = DateSerial(yearN, monthN, dayN) + TimeSerial(hourN, minuteN, secondN)
    If Err.Number <> 0 Then ParseRuDateTime = Empty Else ParseRuDateTime = resultDate
    On Error GoTo 0
End Function

Public Function NormalizeDateKey(ByVal v As Variant) As String
    NormalizeDateKey = ""
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Dim pv As Variant: pv = ParseRuDateTime(v)
    If Not IsEmpty(pv) And VarType(pv) <> vbString Then
        NormalizeDateKey = Format(CDate(pv), "yyyymmddhhnnss")
        Exit Function
    End If
    On Error Resume Next
    Dim d As Date: d = CDate(v)
    If Err.Number = 0 Then NormalizeDateKey = Format(d, "yyyymmddhhnnss")
    Err.Clear: On Error GoTo 0
End Function

'===============================================================
' РАБОТА С КОЛЛЕКЦИЯМИ
'===============================================================
Public Function CollectionHasKey(ByVal col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    col.item key
    CollectionHasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Sub AddErrorToCollection( _
    ByVal errorsCol As Collection, _
    ByVal errType As String, _
    ByVal key As String, _
    ByVal code As String, _
    ByVal projectName As String, _
    ByVal description As String, _
    ByVal sourceName As String)
    Dim item(1 To 7) As Variant
    item(1) = Now: item(2) = errType: item(3) = key: item(4) = code
    item(5) = projectName: item(6) = description: item(7) = sourceName
    errorsCol.Add item
End Sub

Public Sub WriteErrorsToWorkbook(ByVal wbTarget As Workbook, ByVal errorsCol As Collection)
    Dim wsErrors As Worksheet
    Set wsErrors = Nothing
    On Error Resume Next
    Set wsErrors = wbTarget.Worksheets("Ошибки")
    On Error GoTo 0
    If wsErrors Is Nothing Then Exit Sub
    Dim loErrors As ListObject
    Set loErrors = Nothing
    On Error Resume Next
    Set loErrors = wsErrors.ListObjects("тблОшибки")
    On Error GoTo 0
    If loErrors Is Nothing Then Exit Sub
    If Not loErrors.DataBodyRange Is Nothing Then loErrors.DataBodyRange.Delete
    Dim item As Variant, lr As ListRow, i As Long
    For Each item In errorsCol
        Set lr = loErrors.ListRows.Add
        For i = 1 To 7: lr.Range(i).value = item(i): Next i
    Next item
End Sub

'===============================================================
' ПОИСК ЛИСТА ПО ЗАГОЛОВКАМ
'===============================================================
Public Function FindSheetByHeaders(ByVal wb As Workbook) As Worksheet
    Dim requiredHeaders As Variant
    requiredHeaders = Split("Контрагент|Дата создания|Код в базе 1С:УПП|Наименование проекта|Состояние|Продукт|Куратор|Менеджер ОП|Руководитель проекта|Автор", "|")
    Dim minMatches As Long: minMatches = 5
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim lastCol As Long: lastCol = GetLastColumn(ws)
        If lastCol = 0 Then GoTo NextSheet
        Dim matchCount As Long: matchCount = 0
        Dim c As Long, i As Long, headerText As String
        For c = 1 To lastCol
            headerText = GetCleanText(ws.Cells(1, c).value)
            For i = 0 To UBound(requiredHeaders)
                If StrComp(headerText, requiredHeaders(i), vbTextCompare) = 0 Then
                    matchCount = matchCount + 1: Exit For
                End If
            Next i
        Next c
        If matchCount >= minMatches Then
            Set FindSheetByHeaders = ws: Exit Function
        End If
NextSheet:
    Next ws
    Set FindSheetByHeaders = Nothing
End Function

Public Function GetSourceColumnIndexes( _
    ByVal ws As Worksheet, _
    ByVal lastCol As Long, _
    ByRef colIdx As Variant, _
    ByRef missingHeaders As String) As Boolean
    Dim requiredHeaders As Variant
    requiredHeaders = Split("Контрагент|Дата создания|Код в базе 1С:УПП|Код проекта в 1С:УПП|Наименование проекта|Состояние|Тип проекта|Продукт|Куратор|Менеджер ОП|Руководитель проекта|Автор", "|")
    missingHeaders = ""
    Dim unknownHeaders As String
    Dim foundInFile() As Boolean: ReDim foundInFile(1 To lastCol)
    Dim i As Long, c As Long, header As String, found As Boolean
    For i = 0 To UBound(requiredHeaders)
        found = False
        For c = 1 To lastCol
            header = GetCleanText(ws.Cells(1, c).value)
            If StrComp(header, requiredHeaders(i), vbTextCompare) = 0 Then
                colIdx(i + 1) = c: foundInFile(c) = True: found = True: Exit For
            End If
        Next c
        If Not found Then missingHeaders = missingHeaders & "  - " & requiredHeaders(i) & vbCrLf
    Next i
    For c = 1 To lastCol
        header = GetCleanText(ws.Cells(1, c).value)
        If Len(header) > 0 Then
            If Not foundInFile(c) Then unknownHeaders = unknownHeaders & "  - Столбец " & c & ": '" & header & "'" & vbCrLf
        End If
    Next c
    GetSourceColumnIndexes = (Len(missingHeaders) = 0)
    If Not GetSourceColumnIndexes Or Len(unknownHeaders) > 0 Then
        Dim diag As String: diag = ""
        If Len(missingHeaders) > 0 Then diag = diag & "НЕ НАЙДЕНЫ обязательные столбцы:" & vbCrLf & missingHeaders & vbCrLf
        If Len(unknownHeaders) > 0 Then diag = diag & "НЕИЗВЕСТНЫЕ столбцы:" & vbCrLf & unknownHeaders
        MsgBox diag, vbExclamation, "Предупреждение при проверке заголовков"
    End If
End Function

'===============================================================
' ОФОРМЛЕНИЕ: ГРАНИЦЫ И ЦВЕТА
'===============================================================
Public Sub ApplyBorders(ByVal rng As Range)
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlThin: .color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlThick: .color = RGB(0, 0, 0)
    End With
    Dim e As Long
    For e = 7 To 10
        With rng.Borders(e)
            .LineStyle = xlContinuous: .Weight = xlThick: .color = RGB(0, 0, 0)
        End With
    Next e
    With rng.Rows(1).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThick: .color = RGB(0, 0, 0)
    End With
    With rng.Rows(1)
        .Font.Bold = True: .WrapText = True: .VerticalAlignment = xlCenter
    End With
    Dim colOffset As Long
    For colOffset = 1 To rng.Columns.count
        If rng.Cells(1, colOffset).Column > 1 Then rng.Cells(1, colOffset).HorizontalAlignment = xlCenter
    Next colOffset
End Sub

Public Function colLetter(ByVal n As Long) As String
    Dim s As String: s = ""
    Do While n > 0
        Dim rr As Long: rr = (n - 1) Mod 26
        s = Chr(65 + rr) & s
        n = (n - 1) \ 26
    Loop
    colLetter = s
End Function

Public Function HexToColor(ByVal hexText As String) As Long
    Dim s As String: s = UCase(Trim(hexText))
    If Len(s) = 7 And Left(s, 1) = "#" Then s = Mid(s, 2)
    If Len(s) <> 6 Then HexToColor = RGB(255, 255, 255): Exit Function
    Dim r As Long, g As Long, b As Long
    On Error Resume Next
    r = CLng("&H" & Mid(s, 1, 2)): g = CLng("&H" & Mid(s, 3, 2)): b = CLng("&H" & Mid(s, 5, 2))
    If Err.Number <> 0 Then HexToColor = RGB(255, 255, 255): Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    HexToColor = RGB(r, g, b)
End Function

Public Function ColorToHex(ByVal c As Long) As String
    Dim r As Long, g As Long, b As Long
    r = c Mod 256: g = (c \ 256) Mod 256: b = (c \ 65536) Mod 256
    ColorToHex = UCase(Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2))
End Function

'===============================================================
' ПРОГРЕСС-БАР
'===============================================================
Public Sub ProgressShow()
    mLastPct = -1
    ' Проверяем, не загружена ли форма уже
    On Error Resume Next
    If frmProgress Is Nothing Then
        Load frmProgress
    End If
    On Error GoTo 0
    frmProgress.Show vbModeless
    DoEvents
    ' Делаем форму TopMost (поверх всех окон)
    SetWindowPos frmProgress.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub


Public Sub ProgressSet(ByVal pct As Double, ByVal comment As String)
    On Error Resume Next
    frmProgress.SetProgress pct / 100, comment
    On Error GoTo 0
    mLastPct = Int(pct)
    DoEvents
End Sub

Public Sub ProgressSmooth(ByVal pct As Double, ByVal comment As String)
    Dim ip As Long: ip = Int(pct)
    If ip <> mLastPct Then
        mLastPct = ip
        ProgressSet pct, comment
    End If
End Sub

Public Sub ProgressHide()
    On Error Resume Next
    If Not frmProgress Is Nothing Then
        ' Сбрасываем TopMost перед выгрузкой формы
        SetWindowPos frmProgress.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        Unload frmProgress
        Set frmProgress = Nothing  ' Явное освобождение ссылки на форму
    End If
    On Error GoTo 0
End Sub

