Attribute VB_Name = "ModuleStructure"
'===============================================================
' МОДУЛЬ СТРУКТУРЫ (ModuleStructure)
' Реализует СТРОГИЙ ПОРЯДОК ЗАВИСИМОСТЕЙ:
'   ШАГ 1: создание/очистка всех листов;
'   ШАГ 2: создание ListObject с заголовками (справочники первыми,
'          тблПроекты последним из таблиц);
'   ШАГ 3: внедрение формул Calculated Columns (только после того,
'          как тблПроекты, тблСправочникПродуктов и тблЭталон
'          гарантированно существуют в книге);
'   ШАГ 4: оформление;
'   ШАГ 5: ReorderAllSheets.
' Все имена берутся исключительно из констант ModuleConfig.bas.
'===============================================================
Option Explicit

'===============================================================
' ТОЧКА ВХОДА СТРУКТУРЫ: создать файлы решения с нуля.
' Порядок внутри строго соблюдает правило зависимостей.
'===============================================================
Public Sub CreateProjectFiles(ByVal rootFolder As String)
    LogStep "CreateProjectFiles: начало"

    Dim finalPath As String: finalPath = rootFolder & FILE_FINAL
    Dim currentPath As String: currentPath = rootFolder & FILE_CURRENT

    If IsWorkbookOpenInApp(FILE_FINAL) Or IsWorkbookOpenInApp(FILE_CURRENT) Then
        MsgBox "Закройте файлы решения и повторите запуск.", vbExclamation
        Exit Sub
    End If
    If Not ConfirmOverwrite(finalPath) Then Exit Sub
    If Not ConfirmOverwrite(currentPath) Then Exit Sub

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CreateFail

    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)

    ' --- ШАГ 1: ЛИСТЫ (физическое существование до любых объектов) ---
    ProgressSet 8, "Шаг 1: Создание листов..."
    EnsureAllSheets wb

    ' --- ШАГ 2: LISTOBJECT С ЗАГОЛОВКАМИ (таблиц ещё не было!) ---
    ProgressSet 25, "Шаг 2: Создание умных таблиц..."
    CreateAllTables wb

    ' --- ШАГ 3: ФОРМУЛЫ (все целевые таблицы уже созданы) ---
    ProgressSet 45, "Шаг 3: Внедрение формул..."
    ApplyAllFormulas wb

    ' --- ШАГ 4: ОФОРМЛЕНИЕ ---
    ProgressSet 60, "Шаг 4: Оформление..."
    ApplyAllFormatting wb

    ' --- ШАГ 5: СТРОГАЯ ИЕРАРХИЯ ЛИСТОВ ---
    ProgressSet 70, "Шаг 5: Упорядочивание листов..."
    ReorderAllSheets wb

    ' --- Сохранение файлов решения ---
    ProgressSet 80, "Сохранение файлов..."
    DeleteFileIfExists finalPath
    DeleteFileIfExists currentPath
    wb.SaveAs fileName:=finalPath, FileFormat:=xlOpenXMLWorkbook
    wb.SaveCopyAs currentPath
    wb.Close SaveChanges:=False

    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    LogStep "CreateProjectFiles: завершено"
    Exit Sub

CreateFail:
    LogError "CreateProjectFiles", Err.Number, Err.Description
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'===============================================================
' ШАГ 1: СОЗДАНИЕ ВСЕХ ЛИСТОВ КНИГИ (полная очистка существующих)
'===============================================================
Public Sub EnsureAllSheets(ByVal wb As Workbook)
    LogStep "EnsureAllSheets: начало"
    Dim order As Variant: order = SHEET_ORDER()
    Dim i As Long
    For i = LBound(order) To UBound(order)
        GetOrCreateSheet wb, CStr(order(i)), True
    Next i
    LogStep "EnsureAllSheets: завершено"
End Sub

Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String, _
                                 Optional ByVal clearExisting As Boolean = True) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Visible = xlSheetVisible
    If clearExisting Then ClearSheetContents ws
    Set GetOrCreateSheet = ws
End Function

Private Sub ClearSheetContents(ByVal ws As Worksheet)
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop
    Dim i As Long
    For i = ws.ChartObjects.Count To 1 Step -1: ws.ChartObjects(i).Delete: Next i
    For i = ws.Hyperlinks.Count To 1 Step -1: ws.Hyperlinks(i).Delete: Next i
    ws.Cells.Clear
End Sub

'===============================================================
' ШАГ 2: СОЗДАНИЕ ВСЕХ ТАБЛИЦ.
' Справочники и тблЭталон создаются ДО тблПроекты, чтобы формулы
' (ШАГ 3) никогда не ссылались на несуществующие объекты.
'===============================================================
Public Sub CreateAllTables(ByVal wb As Workbook)
    LogStep "CreateAllTables: начало"

    CreateTableFromHeaders wb.Worksheets(SH_LEGEND), TBL_LEGEND, _
        Array("Статус проверки", "Описание", "Цвет (HEX)"), STYLE_MAIN
    PopulateLegend wb.Worksheets(SH_LEGEND)

    CreateTableFromHeaders wb.Worksheets(SH_STATUS_REF), TBL_MANUALSTATUS, _
        Array("Статус"), STYLE_MAIN
    PopulateSingleColumn wb.Worksheets(SH_STATUS_REF), MANUAL_STATUSES()

    CreateTableFromHeaders wb.Worksheets(SH_STATE_REF), TBL_STATES, _
        Array("Допустимое значение"), STYLE_MAIN
    PopulateSingleColumn wb.Worksheets(SH_STATE_REF), ALLOWED_STATES()

    CreateTableFromHeaders wb.Worksheets(SH_PRODUCT_REF), TBL_PRODUCTS, _
        PRODUCT_COLUMNS(), STYLE_PRODUCTS

    CreateTableFromHeaders wb.Worksheets(SH_REFERENCE), TBL_REFERENCE, _
        REFERENCE_COLUMNS(), STYLE_REFERENCE
    wb.Worksheets(SH_REFERENCE).Columns(2).NumberFormatLocal = DATE_FORMAT_LOCAL

    CreateTableFromHeaders wb.Worksheets(SH_ERRORS), TBL_ERRORS, _
        ERROR_COLUMNS(), STYLE_MAIN

    CreateTableFromHeaders wb.Worksheets(SH_EXPORT), TBL_EXPORT, _
        EXPORT_COLUMNS(), STYLE_MAIN

    UpdateReferenceNamedRanges wb

    ' Основная таблица — последней из ШАГа 2: её формулы (ШАГ 3)
    ' будут ссылаться только на уже существующие таблицы.
    Dim loProjects As ListObject
    Set loProjects = CreateTableFromHeaders(wb.Worksheets(SH_PROJECTS), TBL_PROJECTS, _
                                            PROJECT_COLUMNS(), STYLE_MAIN)
    loProjects.ListColumns("Код в базе 1С:УПП").Range.NumberFormatLocal = "@"
    loProjects.ListColumns(COL_KEY).Range.NumberFormatLocal = DATE_FORMAT_LOCAL

    LogStep "CreateAllTables: завершено"
End Sub

' Создание ListObject по вектору заголовков.
' Заголовки пишутся одной пакетной операцией Range.Value = arr.
Public Function CreateTableFromHeaders(ByVal ws As Worksheet, ByVal tableName As String, _
                                       ByRef headers As Variant, _
                                       Optional ByVal styleName As String = "") As ListObject
    Dim nCols As Long: nCols = UBound(headers) - LBound(headers) + 1

    DeleteTableIfExists ws, tableName

    Dim arr() As Variant: ReDim arr(1 To 1, 1 To nCols)
    Dim c As Long
    For c = 1 To nCols: arr(1, c) = headers(LBound(headers) + c - 1): Next c
    ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols)).Value = arr

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, nCols)), , xlYes)
    lo.Name = tableName
    If Len(styleName) > 0 Then lo.TableStyle = styleName
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    Set CreateTableFromHeaders = lo
End Function

Public Sub DeleteTableIfExists(ByVal ws As Worksheet, ByVal tableName As String)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not lo Is Nothing Then
        Application.DisplayAlerts = False
        lo.Delete
        Application.DisplayAlerts = True
    End If
End Sub

' Пакетное заполнение одноколоночного справочника
Private Sub PopulateSingleColumn(ByVal ws As Worksheet, ByRef values As Variant)
    Dim n As Long: n = UBound(values) - LBound(values) + 1
    Dim arr() As Variant: ReDim arr(1 To n, 1 To 1)
    Dim i As Long
    For i = 1 To n: arr(i, 1) = values(LBound(values) + i - 1): Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 1)).Value = arr
    ws.ListObjects(1).Resize ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 1))
End Sub

' Пакетное заполнение легенды статусов
Private Sub PopulateLegend(ByVal ws As Worksheet)
    Dim rows As Variant: rows = LEGEND_ROWS()
    Dim n As Long: n = UBound(rows) - LBound(rows) + 1
    Dim arr() As Variant: ReDim arr(1 To n, 1 To 3)
    Dim i As Long, parts As Variant
    For i = 1 To n
        parts = Split(CStr(rows(LBound(rows) + i - 1)), "|")
        arr(i, 1) = parts(0): arr(i, 2) = parts(1): arr(i, 3) = parts(2)
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 3)).Value = arr
    ws.ListObjects(TBL_LEGEND).Resize ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 3))
End Sub

' Именованные диапазоны для формул валидации и MATCH
Private Sub UpdateReferenceNamedRanges(ByVal wb As Workbook)
    Dim loStates As ListObject, loManual As ListObject
    Set loStates = wb.Worksheets(SH_STATE_REF).ListObjects(TBL_STATES)
    Set loManual = wb.Worksheets(SH_STATUS_REF).ListObjects(TBL_MANUALSTATUS)
    On Error Resume Next
    wb.Names(NAME_ALLOWEDSTATES).Delete
    wb.Names(NAME_MANUALSTATUSES).Delete
    On Error GoTo 0
    wb.Names.Add Name:=NAME_ALLOWEDSTATES, RefersTo:=loStates.DataBodyRange
    wb.Names.Add Name:=NAME_MANUALSTATUSES, RefersTo:=loManual.DataBodyRange
End Sub

'===============================================================
' ШАГ 3: ВНЕДРЕНИЕ ФОРМУЛ (Calculated Columns).
' Вызывается ТОЛЬКО после CreateAllTables. Перед записью любой
' формулы выполняется жёсткая проверка существования целевых
' таблиц (правило зависимостей).
'===============================================================
Public Sub ApplyAllFormulas(ByVal wb As Workbook)
    LogStep "ApplyAllFormulas: начало"

    ' --- Гарантии зависимостей ---
    If GetTable(wb, SH_PRODUCT_REF, TBL_PRODUCTS) Is Nothing Then
        Err.Raise vbObjectError + 513, "ApplyAllFormulas", _
                  "Отсутствует " & TBL_PRODUCTS & ": формулы внедрять нельзя."
    End If
    If GetTable(wb, SH_REFERENCE, TBL_REFERENCE) Is Nothing Then
        Err.Raise vbObjectError + 514, "ApplyAllFormulas", _
                  "Отсутствует " & TBL_REFERENCE & ": формулы внедрять нельзя."
    End If

    Dim lo As ListObject
    Set lo = GetTable(wb, SH_PROJECTS, TBL_PROJECTS)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 515, "ApplyAllFormulas", _
                  "Отсутствует " & TBL_PROJECTS & ": формулы внедрять нельзя."
    End If

    ' Временная строка: DataBodyRange должен существовать,
    ' чтобы Excel растянул формулу как Calculated Column.
    Dim tmpRow As ListRow
    If lo.DataBodyRange Is Nothing Then Set tmpRow = lo.ListRows.Add

    Dim pairs As Variant: pairs = PROJECT_FORMULAS()
    Dim i As Long, parts As Variant
    For i = LBound(pairs) To UBound(pairs)
        parts = Split(CStr(pairs(i)), "|", 2)
        SetColumnFormula lo, CStr(parts(0)), CStr(parts(1))
    Next i

    If Not tmpRow Is Nothing Then tmpRow.Delete

    ' Выпадающий список ручного статуса — ссылка на существующий name
    On Error Resume Next
    With lo.ListColumns("Статус проверки ручной").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=" & NAME_MANUALSTATUSES
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True
        .ErrorTitle = VALIDATION_STATUS_TITLE
        .ErrorMessage = VALIDATION_STATUS_TEXT
    End With
    On Error GoTo 0

    ' Формулы листа Статистика — после создания тблПроекты
    BuildStatisticsSheet wb

    LogStep "ApplyAllFormulas: завершено"
End Sub

' Запись формулы в первую ячейку DataBodyRange: Excel сам
' распространяет её на весь столбец как Calculated Column.
Public Sub SetColumnFormula(ByVal lo As ListObject, ByVal columnName As String, _
                            ByVal formulaText As String)
    On Error Resume Next
    Dim lc As ListColumn
    Set lc = lo.ListColumns(columnName)
    If lc Is Nothing Then On Error GoTo 0: Exit Sub
    If Not lc.DataBodyRange Is Nothing Then
        lc.DataBodyRange.Cells(1, 1).FormulaLocal = formulaText
    End If
    On Error GoTo 0
End Sub

Private Function GetTable(ByVal wb As Workbook, ByVal sheetName As String, _
                          ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    Set GetTable = lo
End Function

'===============================================================
' ЛИСТ СТАТИСТИКА: ТОВАРИЩЕСКИЕ ФОРМУЛЫ COUNTIFS по тблПроекты.
' VBA пишет только строки формул — вычисления выполняет Excel.
'===============================================================
Private Sub BuildStatisticsSheet(ByVal wb As Workbook)
    Dim ws As Worksheet: Set ws = wb.Worksheets(SH_STATISTICS)
    Dim types As Variant: types = PROJECT_TYPES()
    Dim nTypes As Long: nTypes = UBound(types) - LBound(types) + 1

    ws.Range("A1").Value = "Статистика качества заполнения проектов"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ' Матрица заголовков: A3..(Всего) + A4..A11 — одной пакетной операцией
    Dim totalCol As Long: totalCol = 2 + nTypes
    Dim hdr() As Variant: ReDim hdr(1 To 9, 1 To totalCol)
    Dim r As Long, c As Long
    hdr(1, 1) = "Показатель"
    For c = 1 To nTypes: hdr(1, 1 + c) = types(LBound(types) + c - 1): Next c
    hdr(1, totalCol) = "Всего"
    hdr(2, 1) = "Всего проектов"
    hdr(3, 1) = "Без ошибок"
    hdr(4, 1) = "С ошибками"
    hdr(5, 1) = "Закрытые проекты"
    hdr(6, 1) = "Закрытые с ошибками"
    hdr(7, 1) = "Группа ПГС"
    hdr(8, 1) = "Группа PLM"
    hdr(9, 1) = "ПГС и PLM"
    ws.Range(ws.Cells(3, 1), ws.Cells(11, totalCol)).Value = hdr
    ws.Range(ws.Cells(3, 1), ws.Cells(11, 1)).Font.Bold = True

    ' --- Матрица формул: шаблоны собираются из констант ModuleConfig ---
    Dim q As String: q = Chr(34)
    Dim baseCrit As String
    baseCrit = "тблПроекты[" & COL_KEY & "]," & q & "<>" & q & ","

    Dim rowTpl(1 To 8) As String
    rowTpl(1) = Replace(FRM_STAT_BY_TYPE, "{T}", "{C}$3")
    rowTpl(2) = "=COUNTIFS(" & baseCrit & "тблПроекты[Тип проекта],{C}$3,тблПроекты[Число ошибок],0)"
    rowTpl(3) = "=COUNTIFS(" & baseCrit & "тблПроекты[Тип проекта],{C}$3,тблПроекты[Число ошибок]," & q & ">0" & q & ")"
    rowTpl(4) = "=COUNTIFS(" & baseCrit & "тблПроекты[Тип проекта],{C}$3,тблПроекты[Проект закрыт],TRUE)"
    rowTpl(5) = "=COUNTIFS(" & baseCrit & "тблПроекты[Тип проекта],{C}$3,тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок]," & q & "<>0" & q & ")"
    rowTpl(6) = "=COUNTIFS(" & baseCrit & "тблПроекты[Тип проекта],{C}$3,тблПроекты[Группа ПГС]," & q & "да" & q & ")"
    rowTpl(7) = "=COUNTIFS(" & baseCrit & "тблПроекты[Тип проекта],{C}$3,тблПроекты[Группа PLM]," & q & "да" & q & ")"
    rowTpl(8) = "=COUNTIFS(" & baseCrit & "тблПроекты[Тип проекта],{C}$3,тблПроекты[Группа ПГС]," & q & "да" & q & ",тблПроекты[Группа PLM]," & q & "да" & q & ")"

    ' Вектор формул столбца "Всего" (не зависят от типа проекта)
    Dim totFrm(1 To 8) As String
    totFrm(1) = FRM_STAT_TOTAL
    totFrm(2) = FRM_STAT_NOERRORS
    totFrm(3) = FRM_STAT_WITHERRORS
    totFrm(4) = FRM_STAT_CLOSED
    totFrm(5) = FRM_STAT_CLOSED_ERRORS
    totFrm(6) = FRM_STAT_PGS
    totFrm(7) = FRM_STAT_PLM
    totFrm(8) = FRM_STAT_BOTH

    ' Сборка матрицы формул [1..8, 1..totalCol] в памяти...
    Dim fArr() As Variant: ReDim fArr(1 To 8, 1 To totalCol)
    Dim colL As String
    For c = 1 To nTypes
        colL = colLetter(1 + c)
        For r = 1 To 8
            fArr(r, c) = Replace(rowTpl(r), "{C}", colL)
        Next r
    Next c
    For r = 1 To 8: fArr(r, totalCol) = totFrm(r): Next r

    ' ...и одна пакетная запись формул на лист
    ws.Range(ws.Cells(4, 2), ws.Cells(11, totalCol)).FormulaLocal = fArr

    ws.Columns("A:" & colLetter(totalCol)).AutoFit
End Sub

'===============================================================
' ШАГ 4: ОФОРМЛЕНИЕ (стили, скрытие колонок, перенос, ширины,
' условное форматирование по Легенде)
'===============================================================
Public Sub ApplyAllFormatting(ByVal wb As Workbook)
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name <> TBL_PRODUCTS And lo.Name <> TBL_REFERENCE Then
                On Error Resume Next
                lo.TableStyle = STYLE_MAIN
                On Error GoTo 0
            End If
        Next lo
    Next ws

    Dim wsp As Worksheet: Set wsp = wb.Worksheets(SH_PROJECTS)
    Dim lop As ListObject: Set lop = wsp.ListObjects(TBL_PROJECTS)
    lop.Range.EntireColumn.Hidden = False

    Dim hid As Variant: hid = HIDDEN_COLUMNS()
    Dim i As Long
    For i = LBound(hid) To UBound(hid)
        On Error Resume Next
        lop.ListColumns(CStr(hid(i))).Range.EntireColumn.Hidden = True
        On Error GoTo 0
    Next i

    Dim wrp As Variant: wrp = WRAP_COLUMNS()
    For i = LBound(wrp) To UBound(wrp)
        On Error Resume Next
        lop.ListColumns(CStr(wrp(i))).Range.WrapText = True
        On Error GoTo 0
    Next i

    Dim widths As Variant: widths = COLUMN_WIDTHS()
    Dim parts As Variant
    For i = LBound(widths) To UBound(widths)
        parts = Split(CStr(widths(i)), "|")
        On Error Resume Next
        lop.ListColumns(CStr(parts(0))).Range.EntireColumn.ColumnWidth = CDbl(parts(1))
        On Error GoTo 0
    Next i

    PaintLegendColors wsp, lop
    RebuildConditionalFormatting wsp, lop
    wsp.Activate
End Sub

Private Sub PaintLegendColors(ByVal wsp As Worksheet, ByVal lop As ListObject)
    If lop.DataBodyRange Is Nothing Then Exit Sub
    Dim stCol As Long, statusCol As Long
    On Error Resume Next
    stCol = lop.ListColumns("Статус проверки").Index
    On Error GoTo 0
    If stCol = 0 Then Exit Sub
    statusCol = stCol

    Dim legend As Variant: legend = LEGEND_ROWS()
    Dim r As Long, i As Long, parts As Variant, sText As String
    Dim lastRow As Long: lastRow = lop.DataBodyRange.Rows.Count
    For r = 1 To lastRow
        sText = Trim(CStr(lop.DataBodyRange.Cells(r, statusCol).Value))
        For i = LBound(legend) To UBound(legend)
            parts = Split(CStr(legend(i)), "|")
            If StrComp(sText, CStr(parts(0)), vbTextCompare) = 0 Then
                lop.DataBodyRange.Rows(r).Interior.Color = HexToColor(CStr(parts(2)))
                Exit For
            End If
        Next i
    Next r
End Sub

Private Sub RebuildConditionalFormatting(ByVal wsp As Worksheet, ByVal lop As ListObject)
    Dim firstL As String, lastL As String, statusL As String
    firstL = colLetter(lop.Range.Column)
    lastL = colLetter(lop.Range.Column + lop.ListColumns.Count - 1)
    On Error Resume Next
    statusL = colLetter(lop.ListColumns("Статус проверки").Range.Column)
    On Error GoTo 0
    If Len(statusL) = 0 Then Exit Sub

    Dim rng As Range
    Set rng = wsp.Range(firstL & "2:" & lastL & "10000")
    rng.FormatConditions.Delete

    Dim legend As Variant: legend = LEGEND_ROWS()
    Dim i As Long, parts As Variant, fc As FormatCondition
    For i = LBound(legend) To UBound(legend)
        parts = Split(CStr(legend(i)), "|")
        Set fc = rng.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=$" & statusL & "2=""" & parts(0) & """")
        fc.Interior.Color = HexToColor(CStr(parts(2)))
        fc.StopIfTrue = False
    Next i
End Sub

'===============================================================
' ШАГ 5: СТРОГАЯ ИЕРАРХИЯ ЛИСТОВ
' Проекты > Статистика > Качество заполнения по людям > Аналитика >
' Эталон_Данные > СправочникСтатусов > СправочникСостояний >
' СправочникПродуктов > Ошибки (> служебные Легенда, Выгрузка)
'===============================================================
Public Sub ReorderAllSheets(Optional ByVal wb As Workbook = Nothing)
    If wb Is Nothing Then Set wb = ThisWorkbook
    Dim targetOrder As Variant: targetOrder = SHEET_ORDER()
    Dim i As Long, ws As Worksheet
    For i = LBound(targetOrder) To UBound(targetOrder)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(targetOrder(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
            On Error Resume Next
            ws.Move Before:=wb.Worksheets(i + 1)
            On Error GoTo 0
        End If
    Next i
End Sub

'===============================================================
' КОМПЛЕКСНАЯ СБОРКА: Листы -> Таблицы -> Формулы -> Оформление
' -> Иерархия. Используется и при создании, и при апгрейде.
'===============================================================
Public Sub BuildWorkbookStructure(ByVal wb As Workbook)
    EnsureAllSheets wb
    CreateAllTables wb
    ApplyAllFormulas wb
    ApplyAllFormatting wb
    ReorderAllSheets wb
End Sub

'===============================================================
' АПГРЕЙД СТРУКТУРЫ СУЩЕСТВУЮЩЕЙ КНИГИ БЕЗ ПОТЕРИ ДАННЫХ:
' недостающие листы/столбцы добавляются, формулы переустанавливаются.
'===============================================================
Public Sub UpgradeWorkbookStructure(ByVal wb As Workbook)
    LogStep "UpgradeWorkbookStructure: начало"

    ' ШАГ 1: все листы обязательной девятки (без очистки данных)
    Dim order As Variant: order = SHEET_ORDER()
    Dim i As Long
    For i = LBound(order) To UBound(order)
        GetOrCreateSheet wb, CStr(order(i)), False
    Next i

    ' ШАГ 2: недостающие таблицы-зависимости (справочники, эталон)
    If GetTable(wb, SH_PRODUCT_REF, TBL_PRODUCTS) Is Nothing Then
        CreateTableFromHeaders wb.Worksheets(SH_PRODUCT_REF), TBL_PRODUCTS, _
            PRODUCT_COLUMNS(), STYLE_PRODUCTS
    End If
    If GetTable(wb, SH_REFERENCE, TBL_REFERENCE) Is Nothing Then
        CreateTableFromHeaders wb.Worksheets(SH_REFERENCE), TBL_REFERENCE, _
            REFERENCE_COLUMNS(), STYLE_REFERENCE
    End If
    If GetTable(wb, SH_STATE_REF, TBL_STATES) Is Nothing Then
        CreateTableFromHeaders wb.Worksheets(SH_STATE_REF), TBL_STATES, _
            Array("Допустимое значение"), STYLE_MAIN
        PopulateSingleColumn wb.Worksheets(SH_STATE_REF), ALLOWED_STATES()
    End If
    If GetTable(wb, SH_STATUS_REF, TBL_MANUALSTATUS) Is Nothing Then
        CreateTableFromHeaders wb.Worksheets(SH_STATUS_REF), TBL_MANUALSTATUS, _
            Array("Статус"), STYLE_MAIN
        PopulateSingleColumn wb.Worksheets(SH_STATUS_REF), MANUAL_STATUSES()
    End If
    If GetTable(wb, SH_LEGEND, TBL_LEGEND) Is Nothing Then
        CreateTableFromHeaders wb.Worksheets(SH_LEGEND), TBL_LEGEND, _
            Array("Статус проверки", "Описание", "Цвет (HEX)"), STYLE_MAIN
        PopulateLegend wb.Worksheets(SH_LEGEND)
    End If
    If GetTable(wb, SH_ERRORS, TBL_ERRORS) Is Nothing Then
        CreateTableFromHeaders wb.Worksheets(SH_ERRORS), TBL_ERRORS, _
            ERROR_COLUMNS(), STYLE_MAIN
    End If
    UpdateReferenceNamedRanges wb

    ' тблПроекты: создаём при отсутствии либо достраиваем недостающие столбцы
    Dim loP As ListObject
    Set loP = GetTable(wb, SH_PROJECTS, TBL_PROJECTS)
    If loP Is Nothing Then
        Set loP = CreateTableFromHeaders(wb.Worksheets(SH_PROJECTS), TBL_PROJECTS, _
                                         PROJECT_COLUMNS(), STYLE_MAIN)
    Else
        EnsureProjectColumns loP
    End If
    loP.ListColumns("Код в базе 1С:УПП").Range.NumberFormatLocal = "@"
    loP.ListColumns(COL_KEY).Range.NumberFormatLocal = DATE_FORMAT_LOCAL

    ' ШАГ 3: формулы — только теперь, когда все зависимости есть
    ApplyAllFormulas wb

    ' ШАГ 4-5: оформление и иерархия
    ApplyAllFormatting wb
    ReorderAllSheets wb

    LogStep "UpgradeWorkbookStructure: завершено"
End Sub

' Добавление недостающих столбцов в существующую тблПроекты
Private Sub EnsureProjectColumns(ByVal lo As ListObject)
    Dim cols As Variant: cols = PROJECT_COLUMNS()
    Dim i As Long, existing As ListColumn, lc As ListColumn
    For i = LBound(cols) To UBound(cols)
        Set existing = Nothing
        On Error Resume Next
        Set existing = lo.ListColumns(CStr(cols(i)))
        On Error GoTo 0
        If existing Is Nothing Then
            Set lc = lo.ListColumns.Add
            lc.Name = CStr(cols(i))
        End If
    Next i
End Sub
