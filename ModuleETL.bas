Attribute VB_Name = "ModuleETL"
'===============================================================
' МОДУЛЬ ETL (ModuleETL) — ЯДРО ОБРАБОТКИ ДАННЫХ
' Правила архитектуры:
'   1. Чтение: ТОЛЬКО Range.Value (весь диапазон в Variant-массив,
'      одна операция). Поэлементное чтение Cells(r,c).Value запрещено.
'   2. Обработка: ТОЛЬКО в оперативной памяти (Variant-массивы +
'      Scripting.Dictionary). Upsert-ключ — COL_KEY ("Дата создания").
'   3. Запись: ТОЛЬКО Range.Value = arr (одна пакетная операция).
'   4. Логика проверок/вычислений принадлежит формулам Excel;
'      VBA данные не вычисляет и не валидирует.
'===============================================================
Option Explicit

Private Const N_DATA_COLS As Long = 12       ' размер DATA_COLUMNS()/EXPORT_COLUMNS()
Private Const N_ERR_COLS As Long = 7         ' размер ERROR_COLUMNS()
Private Const ADODB_DATE_MIN As Date = #1/1/0100# ' служебные даты 01.01.0001 отбрасываются

'===============================================================
' ТОЧКА ВХОДА ETL: полный цикл Extract -> Transform -> Load
' для файла выгрузки FILE_EXPORT в целевую книгу решения.
'===============================================================
Public Sub UpdateProjectsFromExport(ByVal rootFolder As String)
    gStep = "Инициализация"
    LogStep "=== UpdateProjectsFromExport: НАЧАЛО ==="

    Dim sourcePath As String: sourcePath = rootFolder & FILE_EXPORT
    Dim finalPath As String: finalPath = rootFolder & FILE_FINAL
    Dim currentPath As String: currentPath = rootFolder & FILE_CURRENT

    If Len(Dir(sourcePath)) = 0 Then
        MsgBox "Поместите файл «" & FILE_EXPORT & "» в рабочую папку.", vbExclamation
        Exit Sub
    End If
    If IsWorkbookOpenInApp(FILE_EXPORT) Or IsWorkbookOpenInApp(FILE_FINAL) _
       Or IsWorkbookOpenInApp(FILE_CURRENT) Then
        MsgBox "Закройте файлы выгрузки и решения, затем повторите запуск.", vbExclamation
        Exit Sub
    End If
    EnsureFolder rootFolder & FOLDER_ARCHIVE_EXPORT
    EnsureFolder rootFolder & FOLDER_ARCHIVE_BACKUP

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wbT As Workbook
    On Error GoTo EtlFail

    ' ---------------- EXTRACT: выгрузка -> массив в памяти ----------------
    gStep = "Чтение выгрузки"
    ProgressSet 10, "Чтение выгрузки (Range.Value)..."
    Dim errorArr As Variant, nSrcRows As Long
    Dim srcArr As Variant
    srcArr = ReadExportArray(sourcePath, errorArr, nSrcRows)

    ' ---------------- Открытие целевой книги + резервная копия ----------
    gStep = "Открытие целевого файла"
    ProgressSet 25, "Открытие целевого файла..."
    Dim targetPath As String
    If Len(Dir(finalPath)) > 0 Then
        targetPath = finalPath
    ElseIf Len(Dir(currentPath)) > 0 Then
        targetPath = currentPath
    Else
        MsgBox "Файлы решения не найдены. Сначала выполните «Создать всё с нуля».", vbExclamation
        GoTo EtlRestore
    End If
    Set wbT = Workbooks.Open(fileName:=targetPath, UpdateLinks:=0, AddToMru:=False)
    If wbT.ReadOnly Then
        wbT.Close SaveChanges:=False
        Set wbT = Nothing
        MsgBox "Целевой файл открыт только для чтения.", vbExclamation
        GoTo EtlRestore
    End If
    If targetPath = finalPath Then
        wbT.SaveCopyAs GetUniqueFilePath(CreateTimestampFolder(rootFolder & FOLDER_ARCHIVE_BACKUP), FILE_FINAL)
    End If

    ' ---------------- Гарантия актуальной структуры ----------------
    ' Порядок внутри строгий: листы -> таблицы -> формулы. Формулы
    ' внедряются только когда все целевые ListObject уже существуют.
    gStep = "Проверка структуры книги"
    ProgressSet 40, "Проверка структуры книги..."
    UpgradeWorkbookStructure wbT

    ' ---------------- TRANSFORM: Upsert в оперативной памяти ----------------
    gStep = "Слияние данных (Upsert)"
    ProgressSet 55, "Слияние по ключу «" & COL_KEY & "»..."
    Dim loP As ListObject
    Set loP = wbT.Worksheets(SH_PROJECTS).ListObjects(TBL_PROJECTS)

    Dim oldArr As Variant: oldArr = ReadTableToArray(loP)   ' одно чтение .Value
    Dim merged As Variant, updatedCount As Long, addedCount As Long
    merged = MergeDataInMemory(oldArr, srcArr, loP, updatedCount, addedCount)

    ' ---------------- LOAD: пакетная запись массивов ----------------
    gStep = "Запись данных"
    ProgressSet 75, "Запись итогового массива (Range.Value = arr)..."
    WriteArrayToTable loP, merged

    Dim loE As ListObject
    Set loE = wbT.Worksheets(SH_ERRORS).ListObjects(TBL_ERRORS)
    WriteErrorsToTable loE, errorArr

    ' ---------------- Пересчёт выполняют формулы Excel ----------------
    gStep = "Пересчёт формул"
    ProgressSet 88, "Применение формул (вычисляет Excel)..."
    Application.CalculateFull

    gStep = "Сохранение"
    ProgressSet 94, "Сохранение..."
    If targetPath = finalPath Then
        wbT.Save
    Else
        wbT.SaveAs fileName:=finalPath, FileFormat:=xlOpenXMLWorkbook
    End If
    DeleteFileIfExists currentPath
    wbT.SaveCopyAs currentPath
    wbT.Close SaveChanges:=False: Set wbT = Nothing

    ' Архивация использованной выгрузки
    Dim archFolder As String
    archFolder = CreateTimestampFolder(rootFolder & FOLDER_ARCHIVE_EXPORT)
    Name sourcePath As GetUniqueFilePath(archFolder, FILE_EXPORT)

    ProgressSet 100, "Готово!"
    LogStep "UpdateProjectsFromExport: обновлено=" & updatedCount & " добавлено=" & addedCount
    GoTo EtlRestore

EtlFail:
    LogError gStep, Err.Number, Err.Description
    If Not wbT Is Nothing Then
        On Error Resume Next
        wbT.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Ошибка обновления проектов." & vbCrLf & "Шаг: " & gStep & vbCrLf & _
           Err.description, vbCritical
    Exit Sub

EtlRestore:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'===============================================================
' EXTRACT: чтение всего листа-источника ОДНОЙ операцией Range.Value,
' нормализация в массив [1..n, 1..N_DATA_COLS] + журнал ошибок.
' Ошибки здесь — только ошибки формата данных (дата/ключ/дубль);
' содержательную валидацию выполняют формулы тблПроекты.
'===============================================================
Public Function ReadExportArray(ByVal sourcePath As String, ByRef errorArr As Variant, _
                                ByRef rowCount As Long) As Variant
    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(fileName:=sourcePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    Dim wsSrc As Worksheet
    Set wsSrc = FindSheetByHeaders(wbSrc)
    If wsSrc Is Nothing Then
        wbSrc.Close SaveChanges:=False
        Err.Raise vbObjectError + 600, "ReadExportArray", "Лист с требуемыми заголовками не найден"
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSrc): lastCol = GetLastColumn(wsSrc)

    Dim colIdx As Variant: ReDim colIdx(1 To N_DATA_COLS)
    Dim missing As String
    If Not GetSourceColumnIndexes(wsSrc, lastCol, colIdx, missing) Then
        wbSrc.Close SaveChanges:=False
        Err.Raise vbObjectError + 601, "ReadExportArray", "Отсутствуют обязательные столбцы: " & missing
    End If

    ' ЕДИНСТВЕННОЕ обращение к ячейкам источника — весь диапазон сразу
    Dim raw As Variant
    If lastRow < 2 Then
        wbSrc.Close SaveChanges:=False
        rowCount = 0
        errorArr = Empty
        ReadExportArray = Empty
        Exit Function
    End If
    raw = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol)).Value
    wbSrc.Close SaveChanges:=False

    Dim n As Long: n = lastRow - 1
    rowCount = n
    If n < 1 Then
        errorArr = Empty
        ReadExportArray = Empty
        Exit Function
    End If

    Dim outArr() As Variant: ReDim outArr(1 To n, 1 To N_DATA_COLS)
    Dim keep() As Boolean: ReDim keep(1 To n)
    Dim dictKeys As Object: Set dictKeys = CreateObject("Scripting.Dictionary")
    Dim errList As Object: Set errList = CreateObject("System.Collections.ArrayList")

    Dim r As Long, c As Long, key As String, parsed As Variant

    ' ПАСС 1: нормализация ключа COL_KEY и отсев мусора (данные в память)
    For r = 1 To n
        If Not IsEmptySourceRow(raw, r, colIdx) Then
            parsed = ParseRuDateTime(raw(r, colIdx(2)))
            If IsEmpty(parsed) Then
                AppendError errList, "Ошибка даты", "", GetCodeText(raw(r, colIdx(3))), _
                            GetCleanText(raw(r, colIdx(5))), "Пустая или некорректная дата"
            Else
                If CDate(parsed) < ADODB_DATE_MIN Then
                    AppendError errList, "Исключенная дата", "", GetCodeText(raw(r, colIdx(3))), _
                                GetCleanText(raw(r, colIdx(5))), _
                                "Служебная дата ранее " & Format(ADODB_DATE_MIN, "dd.mm.yyyy")
                Else
                    key = NormalizeDateKey(CDate(parsed))
                    If Len(key) = 0 Then
                        AppendError errList, "Ошибка даты", "", GetCodeText(raw(r, colIdx(3))), _
                                    GetCleanText(raw(r, colIdx(5))), "Не удалось нормализовать ключ"
                    ElseIf dictKeys.Exists(key) Then
                        AppendError errList, "Дубль ключа", key, GetCodeText(raw(r, colIdx(3))), _
                                    GetCleanText(raw(r, colIdx(5))), _
                                    "Пропущен дубль по ключу «" & COL_KEY & "»"
                    Else
                        dictKeys.Add key, r
                        keep(r) = True
                        For c = 1 To N_DATA_COLS
                            Select Case c
                                Case 2: outArr(r, c) = CDate(parsed)
                                Case 3: outArr(r, c) = GetCodeText(raw(r, colIdx(c)))
                                Case Else: outArr(r, c) = GetCleanText(raw(r, colIdx(c)))
                            End Select
                        Next c
                    End If
                End If
            End If
        End If
    Next r

    ' ПАСС 2: сжатие до компактного массива [1..m, 1..N_DATA_COLS]
    Dim m As Long: m = 0
    For r = 1 To n
        If keep(r) Then
            m = m + 1
            If m > r Then
                For c = 1 To N_DATA_COLS
                    outArr(m, c) = outArr(r, c)
                    outArr(r, c) = Empty
                Next c
            End If
        End If
    Next r
    If m <> n Then ReDim Preserve outArr(1 To IIf(m = 0, 1, m), 1 To N_DATA_COLS)

    If errList.Count > 0 Then
        Dim eMat() As Variant: ReDim eMat(0 To errList.Count - 1, 0 To N_ERR_COLS - 1)
        Dim ei As Long, ej As Long, eItem As Variant
        For Each eItem In errList
            For ej = 0 To N_ERR_COLS - 1: eMat(ei, ej) = eItem(ej): Next ej
            ei = ei + 1
        Next eItem
        errorArr = eMat
    Else
        errorArr = Empty
    End If
    ReadExportArray = outArr
End Function

Private Sub AppendError(ByVal list As Object, ByVal eType As String, ByVal key As String, _
                        ByVal code As String, ByVal projName As String, ByVal descr As String)
    Dim item(0 To N_ERR_COLS - 1) As Variant
    item(0) = Now: item(1) = eType: item(2) = key: item(3) = code
    item(4) = projName: item(5) = descr: item(6) = FILE_EXPORT
    list.Add item
End Sub

Public Function CountErrorRows(ByRef errorArr As Variant) As Long
    If IsEmpty(errorArr) Then CountErrorRows = 0: Exit Function
    On Error Resume Next
    CountErrorRows = UBound(errorArr, 1) + 1
    If Err.Number <> 0 Then CountErrorRows = 0
    On Error GoTo 0
End Function

'===============================================================
' Чтение тела ListObject в Variant-массив ОДНОЙ операцией .Value
'===============================================================
Public Function ReadTableToArray(ByVal lo As ListObject) As Variant
    If lo.DataBodyRange Is Nothing Then
        ReadTableToArray = Empty
    Else
        ReadTableToArray = lo.DataBodyRange.Value
    End If
End Function

Private Function ArrayRowCount(ByRef arr As Variant) As Long
    On Error Resume Next
    ArrayRowCount = UBound(arr, 1)
    If Err.Number <> 0 Then ArrayRowCount = 0
    On Error GoTo 0
End Function

'===============================================================
' TRANSFORM: Upsert в оперативной памяти через Scripting.Dictionary.
'  - Ключ сопоставления строк: COL_KEY ("Дата создания"),
'    нормализованный через NormalizeDateKey (yyyymmddhhnnss).
'  - При совпадении ключа перезаписываются ТОЛЬКО столбцы из
'    DATA_COLUMNS(); значения MANUAL_COLUMNS() сохраняются из
'    существующей (старой) строки и не трогаются.
'  - Возвращает новый итоговый массив [1..m, 1..nCols].
' Вычислений логики нет — только перенос значений.
'===============================================================
Public Function MergeDataInMemory(ByRef oldArr As Variant, ByRef newRows As Variant, _
                                  ByVal loTarget As ListObject, _
                                  ByRef updatedCount As Long, ByRef addedCount As Long) As Variant
    Dim headers As Variant
    headers = TableHeadersToArray(loTarget)
    Dim nCols As Long: nCols = loTarget.ListColumns.Count

    Dim dataCols As Variant: dataCols = DATA_COLUMNS()
    Dim manualCols As Variant: manualCols = MANUAL_COLUMNS()

    ' Карта: столбец выгрузки (1..N_DATA_COLS) -> столбец тблПроекты (1..nCols)
    Dim mapDst() As Long: ReDim mapDst(1 To N_DATA_COLS)
    Dim i As Long
    For i = 1 To N_DATA_COLS
        mapDst(i) = HeaderIndex(headers, CStr(dataCols(i)))
    Next i

    ' Столбцы ручного ввода в целевой схеме: НЕ перезаписываются при upsert
    Dim isManual() As Boolean: ReDim isManual(1 To nCols)
    For i = LBound(manualCols) To UBound(manualCols)
        Dim mi As Long: mi = HeaderIndex(headers, CStr(manualCols(i)))
        If mi > 0 Then isManual(mi) = True
    Next i

    Dim keyDst As Long: keyDst = HeaderIndex(headers, COL_KEY)

    ' Индекс существующих строк по ключу COL_KEY (Scripting.Dictionary)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.compareMode = vbBinaryCompare

    Dim nOld As Long: nOld = ArrayRowCount(oldArr)
    Dim keepOld() As Boolean: ReDim keepOld(1 To IIf(nOld = 0, 1, nOld))
    Dim r As Long, c As Long, key As String

    For r = 1 To nOld
        key = NormalizeDateKey(oldArr(r, keyDst))
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                dict.Add key, r          ' позиция строки в oldArr (1-based)
                keepOld(r) = True        ' первая строка с ключом сохраняется
            End If
        Else
            keepOld(r) = True            ' строки без ключа сохраняются как есть
        End If
    Next r

    ' Индекс строк выгрузки по ключу (O(1)-поиск при слиянии)
    Dim dictNew As Object: Set dictNew = CreateObject("Scripting.Dictionary")
    dictNew.compareMode = vbBinaryCompare
    Dim nNew As Long: nNew = ArrayRowCount(newRows)
    For r = 1 To nNew
        If Not IsEmpty(newRows(r, 2)) Then
            key = NormalizeDateKey(newRows(r, 2))
            If Len(key) > 0 Then
                If Not dictNew.Exists(key) Then dictNew.Add key, r
            End If
        End If
    Next r

    ' Итоговое число строк: сохранённые старые + новые из выгрузки
    Dim m As Long: m = 0
    For r = 1 To nOld
        If keepOld(r) Then m = m + 1
    Next r

    Dim useNew() As Boolean: ReDim useNew(1 To IIf(nNew = 0, 1, nNew))
    For r = 1 To nNew
        If Not IsEmpty(newRows(r, 2)) Then
            key = NormalizeDateKey(newRows(r, 2))
            If Len(key) > 0 Then
                If dict.Exists(key) Then
                    updatedCount = updatedCount + 1   ' будет merged-строка вместо старой
                ElseIf CLng(dictNew(key)) = r Then
                    useNew(r) = True                  ' первое вхождение нового ключа
                    addedCount = addedCount + 1
                    m = m + 1
                End If
            Else
                useNew(r) = True                      ' без ключа — новая строка
                addedCount = addedCount + 1
                m = m + 1
            End If
        End If
    Next r

    If m = 0 Then MergeDataInMemory = Empty: Exit Function

    ' Сборка результата [1..m, 1..nCols] в памяти
    Dim result() As Variant: ReDim result(1 To m, 1 To nCols)
    Dim outR As Long: outR = 0

    ' Старые строки: совпавшие по ключу сливаются с данными выгрузки,
    ' значения MANUAL_COLUMNS() остаются из СТАРОЙ строки.
    For r = 1 To nOld
        If keepOld(r) Then
            outR = outR + 1
            For c = 1 To nCols: result(outR, c) = oldArr(r, c): Next c
            key = NormalizeDateKey(oldArr(r, keyDst))
            If Len(key) > 0 Then
                If dictNew.Exists(key) Then
                    Dim nr As Long: nr = CLng(dictNew(key))
                    For i = 1 To N_DATA_COLS
                        If mapDst(i) > 0 And Not isManual(mapDst(i)) Then
                            If i = 2 Then
                                result(outR, mapDst(i)) = CDate(newRows(nr, 2))
                            Else
                                result(outR, mapDst(i)) = newRows(nr, i)
                            End If
                        End If
                    Next i
                End If
            End If
        End If
    Next r

    ' Новые строки выгрузки (ручные поля пусты)
    For r = 1 To nNew
        If useNew(r) Then
            outR = outR + 1
            For i = 1 To N_DATA_COLS
                If mapDst(i) > 0 Then
                    If i = 2 Then
                        result(outR, mapDst(i)) = CDate(newRows(r, 2))
                    Else
                        result(outR, mapDst(i)) = newRows(r, i)
                    End If
                End If
            Next i
        End If
    Next r

    MergeDataInMemory = result
End Function

' Заголовки ListObject в векторный массив (без обращения к ячейкам)
Public Function TableHeadersToArray(ByVal lo As ListObject) As Variant
    Dim n As Long: n = lo.ListColumns.Count
    Dim arr() As Variant: ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n: arr(i) = lo.ListColumns(i).Name: Next i
    TableHeadersToArray = arr
End Function

Private Function HeaderIndex(ByRef headers As Variant, ByVal name As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(i)), name, vbTextCompare) = 0 Then
            HeaderIndex = i: Exit Function
        End If
    Next i
    HeaderIndex = 0
End Function

'===============================================================
' LOAD: пакетная запись массива в ListObject.
' Старые строки удаляются, тело разворачивается до нужного размера,
' запись выполняется ОДНОЙ операцией Range.Value = arr.
'===============================================================
Public Sub WriteArrayToTable(ByVal lo As ListObject, ByRef arr As Variant)
    Dim nRows As Long, nCols As Long
    nRows = ArrayRowCount(arr)
    If nRows = 0 Then
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
        Exit Sub
    End If
    nCols = UBound(arr, 2)

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete

    Do While lo.ListRows.Count < nRows
        lo.ListRows.Add
    Loop

    ' Одна пакетная операция записи всего массива
    lo.DataBodyRange.Value = arr
End Sub

' Пакетная запись журнала ошибок в тблОшибки (одна операция .Value)
Public Sub WriteErrorsToTable(ByVal lo As ListObject, ByRef errorArr As Variant)
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If IsEmpty(errorArr) Then Exit Sub

    Dim nRows As Long: nRows = UBound(errorArr, 1) + 1
    If nRows < 1 Then Exit Sub

    Dim arr() As Variant: ReDim arr(1 To nRows, 1 To N_ERR_COLS)
    Dim r As Long, c As Long
    For r = 0 To nRows - 1
        For c = 0 To N_ERR_COLS - 1
            arr(r + 1, c + 1) = errorArr(r, c)
        Next c
    Next r

    Do While lo.ListRows.Count < nRows
        lo.ListRows.Add
    Loop
    lo.DataBodyRange.Value = arr
End Sub

'===============================================================
' ЭТАЛОН: пакетное чтение эталонного файла и запись в тблЭталон.
' Сравнение "Есть в эталоне" выполняет формула Excel, не VBA.
'===============================================================
Public Sub LoadReferenceIntoTable(ByVal referencePath As String, ByVal loTarget As ListObject)
    Dim wbRef As Workbook
    Set wbRef = Workbooks.Open(fileName:=referencePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    Dim wsRef As Worksheet
    Set wsRef = FindSheetByHeaders(wbRef)
    If wsRef Is Nothing Then
        wbRef.Close SaveChanges:=False
        Err.Raise vbObjectError + 602, "LoadReferenceIntoTable", _
                  "В эталонном файле не найден лист с требуемыми заголовками"
    End If

    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsRef): lastCol = GetLastColumn(wsRef)

    Dim refArr As Variant
    If lastRow >= 2 Then
        refArr = wsRef.Range(wsRef.Cells(2, 1), wsRef.Cells(lastRow, lastCol)).Value
    End If
    wbRef.Close SaveChanges:=False

    WriteArrayToTable loTarget, refArr
End Sub
