Attribute VB_Name = "ModuleBuilder"
'===============================================================
' МОДУЛЬ СОЗДАНИЯ ЛИСТОВ И ОФОРМЛЕНИЯ (ModuleBuilder)
'===============================================================
Option Explicit

'===============================================================
' КОНСТАНТЫ: РАЗДЕЛЕНИЕ СТОЛБЦОВ НА ДАННЫЕ И ФОРМУЛЬНЫЕ
' Критическое ограничение: Запись через .Value разрешена ТОЛЬКО в столбцы данных
'===============================================================

' Столбцы, разрешенные для записи данных через .Value (исходные данные из выгрузки)
Public Const DATA_COLUMNS As String = _
    "Контрагент|Дата создания|Код в базе 1С:УПП|Код проекта в 1С:УПП|" & _
    "Наименование проекта|Состояние|Тип проекта|Продукт|" & _
    "Куратор|Менеджер ОП|Руководитель проекта|Автор"

' Столбцы, заполняемые ТОЛЬКО формулами (запрещена запись .Value)
Public Const FORMULA_COLUMNS As String = _
    "Группа ПГС|Группа PLM|" & _
    "Проверка Контрагент|Проверка Дата создания|Проверка Код|Проверка Код проекта в 1С:УПП|" & _
    "Проверка Наименование проекта|Проверка Состояние|Недопустимое Состояние|" & _
    "Проверка Куратор|Проверка Менеджер ОП|Список некорректных полей|" & _
    "Ошибка обязательных полей|Число ошибок|Проект закрыт|" & _
    "Дубль кода|Дубль ссылки|Статус проверки авто|Комментарий к статусу авто|" & _
    "Статус проверки|Комментарий к статусу|Есть в эталоне"

' Столбцы ручного ввода (не перезаписываются при обновлении, сохраняются пользователем)
Public Const MANUAL_COLUMNS As String = _
    "Ссылка на проект|Ответственный за статус|Статус проверки ручной|Комментарий к статусу ручной"

'===============================================================
' УДАЛЕНИЕ УСТАРЕВШИХ СТОЛБЦОВ
'===============================================================
Public Sub RemoveOldColumns(ByVal lo As ListObject)
    Dim lcOld As ListColumn
    Set lcOld = Nothing
    On Error Resume Next: Set lcOld = lo.ListColumns("Ключ проекта"): On Error GoTo 0
    If Not lcOld Is Nothing Then lcOld.Delete
    
    Set lcOld = Nothing
    On Error Resume Next: Set lcOld = lo.ListColumns("Код только цифры"): On Error GoTo 0
    If Not lcOld Is Nothing Then lcOld.Delete
    
    Set lcOld = Nothing
    On Error Resume Next: Set lcOld = lo.ListColumns("Есть в сводной"): On Error GoTo 0
    If Not lcOld Is Nothing Then lcOld.Delete
End Sub

'===============================================================
' ДОБАВЛЕНИЕ СТОЛБЦА ПОСЛЕ УКАЗАННОГО
'===============================================================
Public Sub EnsureColumnAfter(ByVal lo As ListObject, ByVal newName As String, ByVal afterName As String)
    Dim lc As ListColumn
    Set lc = Nothing
    On Error Resume Next: Set lc = lo.ListColumns(newName): On Error GoTo 0
    If Not lc Is Nothing Then Exit Sub
    Dim pos As Long: pos = lo.ListColumns(afterName).Index + 1
    Set lc = lo.ListColumns.Add(Position:=pos)
    lc.name = newName
End Sub

'===============================================================
' ЗАПИСЬ СТРОКИ ВЫГРУЗКИ В ТАБЛИЦУ
' Критическое ограничение: Запись .Value выполняется ТОЛЬКО в столбцы данных.
' Столбцы ручного ввода (MANUAL_COLUMNS) не перезаписываются при обновлении существующих строк.
' Формульные столбцы заполняются через SetColumnFormula (Calculated Columns).
'===============================================================
Public Sub WriteSourceRowToListRow( _
    ByVal lo As ListObject, ByVal lr As ListRow, _
    ByVal srcValues As Variant, ByVal r As Long, _
    ByVal colIdx As Variant, ByVal parsedDate As Date, ByVal codeText As String)
    
    ' Явная запись только в разрешенные столбцы данных (DATA_COLUMNS)
    ' Порядок и список столбцов синхронизирован с константой DATA_COLUMNS
    
    ' 1. Контрагент
    lr.Range(lo.ListColumns("Контрагент").Index).value = GetCleanText(srcValues(r, colIdx(1)))
    ' 2. Дата создания
    lr.Range(lo.ListColumns("Дата создания").Index).value = parsedDate
    ' 3. Код в базе 1С:УПП
    lr.Range(lo.ListColumns("Код в базе 1С:УПП").Index).value = codeText
    ' 4. Код проекта в 1С:УПП
    lr.Range(lo.ListColumns("Код проекта в 1С:УПП").Index).value = GetCleanText(srcValues(r, colIdx(4)))
    ' 5. Наименование проекта
    lr.Range(lo.ListColumns("Наименование проекта").Index).value = GetCleanText(srcValues(r, colIdx(5)))
    ' 6. Состояние
    lr.Range(lo.ListColumns("Состояние").Index).value = GetCleanText(srcValues(r, colIdx(6)))
    ' 7. Тип проекта
    lr.Range(lo.ListColumns("Тип проекта").Index).value = GetCleanText(srcValues(r, colIdx(7)))
    ' 8. Продукт
    lr.Range(lo.ListColumns("Продукт").Index).value = GetCleanText(srcValues(r, colIdx(8)))
    ' 9. Куратор
    lr.Range(lo.ListColumns("Куратор").Index).value = GetCleanText(srcValues(r, colIdx(9)))
    ' 10. Менеджер ОП
    lr.Range(lo.ListColumns("Менеджер ОП").Index).value = GetCleanText(srcValues(r, colIdx(10)))
    ' 11. Руководитель проекта
    lr.Range(lo.ListColumns("Руководитель проекта").Index).value = GetCleanText(srcValues(r, colIdx(11)))
    ' 12. Автор
    lr.Range(lo.ListColumns("Автор").Index).value = GetCleanText(srcValues(r, colIdx(12)))
    
    ' ПРИМЕЧАНИЕ: Столбцы ручного ввода (Ссылка на проект, Ответственный за статус, 
    ' Статус проверки ручной, Комментарий к статусу ручной) НЕ перезаписываются здесь.
    ' Они сохраняются при обновлении существующих строк и остаются пустыми при добавлении новых.
    
    ' Формульные столбцы (Группа ПГС, Группа PLM, Проверка*, Статус проверки авто и др.)
    ' заполняются автоматически через механизм Calculated Columns умной таблицы Excel.
End Sub

'===============================================================
' ФОРМУЛЫ ДЛЯ тблПроекты
'===============================================================
Public Sub SetAllProjectFormulas(ByVal lo As ListObject)
    SetColumnFormula lo, "Проверка Контрагент", "=IF(ISBLANK([@[Дата создания]]),FALSE,OR(LEN(TRIM([@[Контрагент]]))=0,LEN(TRIM([@[Контрагент]]))=1,AND(LEN(TRIM([@[Контрагент]]))>1,EXACT(TRIM([@[Контрагент]]),REPT(LEFT(TRIM([@[Контрагент]]),1),LEN(TRIM([@[Контрагент]])))))))"
    SetColumnFormula lo, "Проверка Дата создания", "=IF(ISBLANK([@[Дата создания]]),FALSE,OR(ISBLANK([@[Дата создания]]),NOT(ISNUMBER([@[Дата создания]]))))"
    SetColumnFormula lo, "Проверка Код", "=IF(AND(ISBLANK([@[Дата создания]]),TRIM([@[Код в базе 1С:УПП]])=""""),FALSE,OR(LEN(TRIM([@[Код в базе 1С:УПП]]))=0,LEFT(TRIM([@[Код в базе 1С:УПП]]),2)<>""Я-""))"
    SetColumnFormula lo, "Проверка Код проекта в 1С:УПП", "=IF(ISBLANK([@[Дата создания]]),FALSE,OR(LEN(TRIM([@[Код проекта в 1С:УПП]]))=0,LEN(TRIM([@[Код проекта в 1С:УПП]]))=1,AND(LEN(TRIM([@[Код проекта в 1С:УПП]]))>1,EXACT(TRIM([@[Код проекта в 1С:УПП]]),REPT(LEFT(TRIM([@[Код проекта в 1С:УПП]]),1),LEN(TRIM([@[Код проекта в 1С:УПП]])))))))"
    SetColumnFormula lo, "Проверка Наименование проекта", "=IF(ISBLANK([@[Дата создания]]),FALSE,OR(LEN(TRIM([@[Наименование проекта]]))=0,LEN(TRIM([@[Наименование проекта]]))=1,AND(LEN(TRIM([@[Наименование проекта]]))>1,EXACT(TRIM([@[Наименование проекта]]),REPT(LEFT(TRIM([@[Наименование проекта]]),1),LEN(TRIM([@[Наименование проекта]])))))))"
    SetColumnFormula lo, "Проверка Состояние", "=IF(ISBLANK([@[Дата создания]]),FALSE,OR(LEN(TRIM([@[Состояние]]))=0,LEN(TRIM([@[Состояние]]))=1,AND(LEN(TRIM([@[Состояние]]))>1,EXACT(TRIM([@[Состояние]]),REPT(LEFT(TRIM([@[Состояние]]),1),LEN(TRIM([@[Состояние]])))))))"
    SetColumnFormula lo, "Недопустимое Состояние", "=IF(ISBLANK([@[Дата создания]]),FALSE,IF(LEN(TRIM([@[Состояние]]))=0,FALSE,ISNA(MATCH(TRIM([@[Состояние]]),ДопустимыеСостояния,0))))"
    SetColumnFormula lo, "Проверка Куратор", "=IF(ISBLANK([@[Дата создания]]),FALSE,OR(LEN(TRIM([@[Куратор]]))=0,LEN(TRIM([@[Куратор]]))=1,AND(LEN(TRIM([@[Куратор]]))>1,EXACT(TRIM([@[Куратор]]),REPT(LEFT(TRIM([@[Куратор]]),1),LEN(TRIM([@[Куратор]])))))))"
    SetColumnFormula lo, "Проверка Менеджер ОП", "=IF(ISBLANK([@[Дата создания]]),FALSE,OR(LEN(TRIM([@[Менеджер ОП]]))=0,LEN(TRIM([@[Менеджер ОП]]))=1,AND(LEN(TRIM([@[Менеджер ОП]]))>1,EXACT(TRIM([@[Менеджер ОП]]),REPT(LEFT(TRIM([@[Менеджер ОП]]),1),LEN(TRIM([@[Менеджер ОП]])))))))"
    SetColumnFormula lo, "Список некорректных полей", "=IF(ISBLANK([@[Дата создания]]),"""",TEXTJOIN("", "",TRUE,IF([@[Проверка Контрагент]],""Контрагент"",""""),IF([@[Проверка Дата создания]],""Дата создания"",""""),IF([@[Проверка Код]],""Код в базе 1С:УПП"",""""),IF([@[Проверка Код проекта в 1С:УПП]],""Код проекта в 1С:УПП"",""""),IF([@[Проверка Наименование проекта]],""Наименование проекта"",""""),IF(OR([@[Проверка Состояние]],[@[Недопустимое Состояние]]),""Состояние"",""""),IF([@[Проверка Куратор]],""Куратор"",""""),IF([@[Проверка Менеджер ОП]],""Менеджер ОП"","""")))"
    SetColumnFormula lo, "Ошибка обязательных полей", "=IF([@[Список некорректных полей]]="""",FALSE,TRUE)"
    SetColumnFormula lo, "Число ошибок", "=([@[Проверка Контрагент]]+[@[Проверка Дата создания]]+[@[Проверка Код]]+[@[Проверка Код проекта в 1С:УПП]]+[@[Проверка Наименование проекта]]+[@[Проверка Состояние]]+[@[Недопустимое Состояние]]+[@[Проверка Куратор]]+[@[Проверка Менеджер ОП]])"
    SetColumnFormula lo, "Проект закрыт", "=IF([@[Дата создания]]="""","""",OR([@[Состояние]]=""Завершен"",[@[Состояние]]=""Прекращен с отрицательным результатом"",[@[Состояние]]=""Не состоялся""))"
    SetColumnFormula lo, "Группа ПГС", "=IF(ISBLANK([@[Продукт]]),""""," & "IF(ISNUMBER(SEARCH(""ПГС"",IFERROR(VLOOKUP(TRIM([@[Продукт]]),тблСправочникПродуктов,2,FALSE),""""))),""да""," & "IF(ISNUMBER(SEARCH(""ТИМ"",IFERROR(VLOOKUP(TRIM([@[Продукт]]),тблСправочникПродуктов,2,FALSE),""""))),""да"","""")))"
    SetColumnFormula lo, "Группа PLM", "=IF(ISBLANK([@[Продукт]]),""""," & "IF(ISNUMBER(SEARCH(""PLM"",IFERROR(VLOOKUP(TRIM([@[Продукт]]),тблСправочникПродуктов,2,FALSE),""""))),""да"",""""))"
    SetColumnFormula lo, "Дубль кода", "=IF(OR(ISBLANK([@[Дата создания]]),LEN(TRIM([@[Код в базе 1С:УПП]]))=0),FALSE,COUNTIFS([Код в базе 1С:УПП],[@[Код в базе 1С:УПП]],[Статус проверки ручной],""<>Архивировать"")>1)"
    SetColumnFormula lo, "Дубль ссылки", "=IF(OR(ISBLANK([@[Дата создания]]),LEN(TRIM([@[Ссылка на проект]]))=0),FALSE,COUNTIFS([Ссылка на проект],[@[Ссылка на проект]],[Статус проверки ручной],""<>Архивировать"")>1)"
    SetColumnFormula lo, "Статус проверки авто", "=IF(ISBLANK([@[Дата создания]]),"""",IF(OR([@[Ошибка обязательных полей]],[@[Дубль кода]],[@[Дубль ссылки]]),IF(LEN(TRIM([@[Ответственный за статус]]))=0,""Не назначен Ответственный за статус"",""Не соответствует требованиям""),IF(LEN(TRIM([@[Ссылка на проект]]))=0,""Требуется ссылка"",IF(OR(TRIM([@[Состояние]])=""Завершен"",TRIM([@[Состояние]])=""Прекращен с отрицательным результатом"",TRIM([@[Состояние]])=""Не состоялся""),""Готов к передаче в архив"",""Соответствует требованиям""))))"
    SetColumnFormula lo, "Комментарий к статусу авто", "=IF(ISBLANK([@[Дата создания]]),"""",TRIM(IF([@[Ошибка обязательных полей]],""Некорректные данные в поле: ""&[@[Список некорректных полей]],"""")&IF([@[Дубль кода]],IF([@[Ошибка обязательных полей]],""; "","""")&""Дублируется Код в базе 1С:УПП"","""")&IF([@[Дубль ссылки]],IF(OR([@[Ошибка обязательных полей]],[@[Дубль кода]]),""; "","""")&""Дублируется Ссылка на проект"","""")))"
    SetColumnFormula lo, "Статус проверки", "=IF(ISBLANK([@[Дата создания]]),"""",IF(OR([@[Статус проверки ручной]]=""Архивировать"",[@[Статус проверки ручной]]=""Уточнить!"",[@[Статус проверки ручной]]=""В работе""),[@[Статус проверки ручной]],[@[Статус проверки авто]]))"
    SetColumnFormula lo, "Комментарий к статусу", "=TRIM([@[Комментарий к статусу авто]]&IF(AND([@[Комментарий к статусу авто]]<>"",TRIM([@[Комментарий к статусу ручной]])<>""),""; "","""")&TRIM([@[Комментарий к статусу ручной]]))"
    SetColumnFormula lo, "Есть в эталоне", "=ЕСЛИ(ЕОШИБКА(ПОИСКПОЗ([@[Дата создания]];тблЭталон[Дата создания];0));""Нет"";""Да"")"
End Sub

'===============================================================
' УСТАНОВКА ФОРМУЛЫ ЧЕРЕЗ CALCULATED COLUMN
' Оптимизация: вместо записи формулы во весь диапазон DataBodyRange,
' записываем только в первую ячейку. Excel сам распространяет формулу
' на весь столбец умной таблицы (механизм Calculated Column).
' Это дает значительный прирост производительности на больших таблицах.
'===============================================================
Public Sub SetColumnFormula(ByVal lo As ListObject, ByVal columnName As String, ByVal formulaText As String)
    On Error Resume Next
    Dim lc As ListColumn
    Set lc = lo.ListColumns(columnName)
    If lc Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    
    If Not lc.DataBodyRange Is Nothing Then
        ' Оптимизация: записываем формулу только в первую ячейку столбца
        ' Excel автоматически расширяет её на весь столбец (Calculated Column)
        lc.DataBodyRange.Cells(1, 1).formula = formulaText
    End If
    On Error GoTo 0
End Sub

'===============================================================
' ОФОРМЛЕНИЕ ЛИСТА ПРОЕКТЫ
'===============================================================
Public Sub ApplyProjectsFormatting(ByVal wsProjects As Worksheet, ByVal wsLegend As Worksheet)
    Dim lo As ListObject: Set lo = wsProjects.ListObjects("тблПроекты")
    If lo Is Nothing Then Exit Sub
    lo.tableStyle = "TableStyleLight13"
    lo.Range.EntireColumn.Hidden = False
    
    HideColumnByName lo, "Проверка Контрагент"
    HideColumnByName lo, "Проверка Дата создания"
    HideColumnByName lo, "Проверка Код"
    HideColumnByName lo, "Проверка Код проекта в 1С:УПП"
    HideColumnByName lo, "Проверка Наименование проекта"
    HideColumnByName lo, "Проверка Состояние"
    HideColumnByName lo, "Недопустимое Состояние"
    HideColumnByName lo, "Проверка Куратор"
    HideColumnByName lo, "Проверка Менеджер ОП"
    HideColumnByName lo, "Список некорректных полей"
    HideColumnByName lo, "Ошибка обязательных полей"
    HideColumnByName lo, "Дубль кода"
    HideColumnByName lo, "Дубль ссылки"
    HideColumnByName lo, "Статус проверки авто"
    HideColumnByName lo, "Комментарий к статусу авто"
    HideColumnByName lo, "Число ошибок"
    HideColumnByName lo, "Проект закрыт"
    
    WrapColumnByName lo, "Контрагент"
    WrapColumnByName lo, "Наименование проекта"
    WrapColumnByName lo, "Продукт"
    WrapColumnByName lo, "Куратор"
    WrapColumnByName lo, "Менеджер ОП"
    WrapColumnByName lo, "Руководитель проекта"
    WrapColumnByName lo, "Автор"
    WrapColumnByName lo, "Ссылка на проект"
    WrapColumnByName lo, "Ответственный за статус"
    WrapColumnByName lo, "Комментарий к статусу ручной"
    WrapColumnByName lo, "Комментарий к статусу"
    WrapColumnByName lo, "Тип проекта"
    
    SetColumnWidthByName lo, "Контрагент", 25
    SetColumnWidthByName lo, "Код в базе 1С:УПП", 18
    SetColumnWidthByName lo, "Код проекта в 1С:УПП", 20
    SetColumnWidthByName lo, "Наименование проекта", 40
    SetColumnWidthByName lo, "Состояние", 20
    SetColumnWidthByName lo, "Тип проекта", 20
    SetColumnWidthByName lo, "Продукт", 20
    SetColumnWidthByName lo, "Куратор", 20
    SetColumnWidthByName lo, "Менеджер ОП", 20
    SetColumnWidthByName lo, "Руководитель проекта", 25
    SetColumnWidthByName lo, "Автор", 20
    SetColumnWidthByName lo, "Ссылка на проект", 30
    SetColumnWidthByName lo, "Ответственный за статус", 20
    SetColumnWidthByName lo, "Комментарий к статусу ручной", 30
    SetColumnWidthByName lo, "Комментарий к статусу", 45
    
    PaintLegend wsLegend
    RebuildProjectsConditionalFormatting wsProjects, wsLegend
End Sub

Private Sub HideColumnByName(ByVal lo As ListObject, ByVal name As String)
    On Error Resume Next: lo.ListColumns(name).Range.EntireColumn.Hidden = True: On Error GoTo 0
End Sub

Private Sub WrapColumnByName(ByVal lo As ListObject, ByVal name As String)
    On Error Resume Next: lo.ListColumns(name).Range.EntireColumn.WrapText = True: On Error GoTo 0
End Sub

Private Sub SetColumnWidthByName(ByVal lo As ListObject, ByVal name As String, ByVal w As Double)
    On Error Resume Next: lo.ListColumns(name).Range.EntireColumn.ColumnWidth = w: On Error GoTo 0
End Sub

'===============================================================
' УСЛОВНОЕ ФОРМАТИРОВАНИЕ
'===============================================================
Private Sub RebuildProjectsConditionalFormatting(ByVal wsProjects As Worksheet, ByVal wsLegend As Worksheet)
    Dim lo As ListObject: Set lo = wsProjects.ListObjects("тблПроекты")
    If lo Is Nothing Then Exit Sub
    Dim loLegend As ListObject: Set loLegend = wsLegend.ListObjects("тблЛегендаСтатусов")
    If loLegend Is Nothing Then Exit Sub
    If loLegend.DataBodyRange Is Nothing Then Exit Sub
    
    Dim firstLetter As String, lastLetter As String, statusLetter As String
    firstLetter = colLetter(lo.Range.Column)
    lastLetter = colLetter(lo.Range.Column + lo.ListColumns.count - 1)
    statusLetter = colLetter(lo.ListColumns("Статус проверки").Range.Column)
    
    Dim rng As Range
    Set rng = wsProjects.Range(firstLetter & "2:" & lastLetter & "10000")
    rng.FormatConditions.Delete
    
    Dim r As Long, statusText As String, hexText As String
    For r = 1 To loLegend.DataBodyRange.Rows.count
        statusText = Trim(CStr(loLegend.DataBodyRange.Cells(r, 1).value))
        hexText = Trim(CStr(loLegend.DataBodyRange.Cells(r, 3).value))
        If Len(statusText) > 0 And Len(hexText) > 0 Then
            AddCondition rng, "=$" & statusLetter & "2=""" & statusText & """", HexToColor(hexText)
        End If
    Next r
End Sub

Private Sub AddCondition(ByVal rng As Range, ByVal formula As String, ByVal color As Long)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)
    fc.Interior.color = color
End Sub

'===============================================================
' ЛЕГЕНДА: ЦВЕТА И СИНХРОНИЗАЦИЯ
'===============================================================
Private Sub SyncLegendColorsFromFill(ByVal ws As Worksheet)
    Dim lo As ListObject: Set lo = ws.ListObjects("тблЛегендаСтатусов")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim r As Long
    For r = 1 To lo.DataBodyRange.Rows.count
        Dim cellFill As Range: Set cellFill = lo.DataBodyRange.Cells(r, 1)
        Dim existingHex As String: existingHex = Trim(CStr(lo.DataBodyRange.Cells(r, 3).value))
        Dim hexText As String
        If cellFill.Interior.ColorIndex <> xlNone Then
            hexText = ColorToHex(cellFill.Interior.color)
        Else
            hexText = existingHex
        End If
        lo.DataBodyRange.Cells(r, 3).value = hexText
    Next r
End Sub

Private Sub PaintLegend(ByVal ws As Worksheet)
    Dim lo As ListObject: Set lo = ws.ListObjects("тблЛегендаСтатусов")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim r As Long, hexText As String, rngRow As Range
    For r = 1 To lo.DataBodyRange.Rows.count
        hexText = Trim(CStr(lo.DataBodyRange.Cells(r, 3).value))
        Set rngRow = lo.DataBodyRange.Rows(r)
        If Len(hexText) > 0 Then
            rngRow.Interior.color = HexToColor(hexText)
        Else
            rngRow.Interior.ColorIndex = xlNone
        End If
    Next r
End Sub

Public Sub SaveColorsToSettings(ByRef st() As String, ByRef cL() As String, ByVal count As Long)
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("Настройки"): On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "Настройки"
    End If
    ws.Columns("A:B").Clear
    ws.Range("A1").value = "Настройки цветов статусов"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    Dim note As String
    note = "Служебная вкладка. Хранит цвета статусов проверки."
    ws.Range("A2").value = note: ws.Range("A2").WrapText = True
    ws.Range("A4").value = "Статус проверки": ws.Range("B4").value = "Цвет (HEX)"
    Dim i As Long
    For i = 0 To count - 1
        ws.Cells(i + 5, 1).value = st(i): ws.Cells(i + 5, 2).value = cL(i)
        If Len(cL(i)) > 0 Then
            ws.Range(ws.Cells(i + 5, 1), ws.Cells(i + 5, 2)).Interior.color = HexToColor(cL(i))
        Else
            ws.Range(ws.Cells(i + 5, 1), ws.Cells(i + 5, 2)).Interior.ColorIndex = xlNone
        End If
    Next i
    ws.Columns("A").ColumnWidth = 45: ws.Columns("B").ColumnWidth = 12
    ws.Rows(2).EntireRow.AutoFit
End Sub

Private Sub ApplySavedColors(ByRef st() As String, ByRef cL() As String, ByVal count As Long)
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("Настройки"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim r As Long, i As Long, sName As String, sHex As String
    For r = 5 To lastRow
        sName = Trim(CStr(ws.Cells(r, 1).value)): sHex = Trim(CStr(ws.Cells(r, 2).value))
        If Len(sName) > 0 Then
            For i = 0 To count - 1
                If StrComp(st(i), sName, vbTextCompare) = 0 Then cL(i) = sHex
            Next i
        End If
    Next r
End Sub

Private Sub SyncLegendInWorkbook(ByVal wb As Workbook, ByVal ws As Worksheet, ByRef st() As String, ByRef ds() As String, ByRef cL() As String, ByVal count As Long)
    Dim lo As ListObject
    Set lo = Nothing
    On Error Resume Next: Set lo = ws.ListObjects("тблЛегендаСтатусов"): On Error GoTo 0
    If lo Is Nothing Then
        ws.Range("A1").value = "Статус проверки": ws.Range("B1").value = "Описание": ws.Range("C1").value = "Цвет (HEX)"
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C2"), , xlYes)
        lo.name = "тблЛегендаСтатусов": lo.tableStyle = "TableStyleLight13"
    End If
    If lo.ListColumns.count < 3 Then
        Dim firstRow As Long, lastRow As Long
        firstRow = lo.Range.row: lastRow = lo.Range.row + lo.Range.Rows.count - 1
        lo.Resize ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, 3))
        ws.Cells(firstRow, 3).value = "Цвет (HEX)"
    End If
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    Dim i As Long
    For i = 0 To count - 1
        Dim lr As ListRow: Set lr = lo.ListRows.Add
        lr.Range(1).value = st(i): lr.Range(2).value = ds(i): lr.Range(3).value = cL(i)
    Next i
End Sub

Public Function ReadLegendFromWorkbook(ByVal wb As Workbook, ByRef st() As String, ByRef ds() As String, ByRef cL() As String) As Boolean
    Dim lo As ListObject
    Set lo = Nothing
    On Error Resume Next: Set lo = wb.Worksheets("Легенда").ListObjects("тблЛегендаСтатусов"): On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then ReadLegendFromWorkbook = False: Exit Function
    Dim colCount As Long: colCount = lo.ListColumns.count
    ReDim st(0 To 50): ReDim ds(0 To 50): ReDim cL(0 To 50)
    Dim n As Long: n = 0
    Dim r As Long, s As String
    For r = 1 To lo.DataBodyRange.Rows.count
        s = SafeText(lo.DataBodyRange.Cells(r, 1).value)
        If Len(s) > 0 Then
            st(n) = s
            If colCount >= 2 Then ds(n) = SafeText(lo.DataBodyRange.Cells(r, 2).value) Else ds(n) = ""
            If colCount >= 3 Then cL(n) = SafeText(lo.DataBodyRange.Cells(r, 3).value) Else cL(n) = ""
            n = n + 1
        End If
    Next r
    ReadLegendFromWorkbook = (n > 0)
End Function

Private Function ReadLegendFromFile(ByVal filePath As String, ByRef st() As String, ByRef ds() As String, ByRef cL() As String) As Boolean
    Dim wb As Workbook: Set wb = Nothing
    On Error Resume Next: Set wb = Workbooks.Open(fileName:=filePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False): On Error GoTo 0
    If wb Is Nothing Then ReadLegendFromFile = False: Exit Function
    ReadLegendFromFile = ReadLegendFromWorkbook(wb, st, ds, cL)
    wb.Close SaveChanges:=False
End Function

Private Sub GetDefaultLegend(ByRef st() As String, ByRef ds() As String, ByRef cL() As String)
    st = Split("Архивировать|Соответствует требованиям|Не назначен Ответственный за статус|Не соответствует требованиям|Уточнить!|Требуется ссылка|Готов к передаче в архив|В работе", "|")
    ds = Split("Проект проверен пользователем и одобрен для сдачи в архив|Проект не проверен пользователем. Обязательные поля корректно заполнены|Обязательные поля заполнены некорректно. Не назначен пользователь, ответственный за проверку проекта|Ответственный пользователь назначен. Обязательные поля заполнены некорректно|Ответственный пользователь назначен. Возможны ошибки в данных, корректность которых ответственный проверяет|Обязательные поля заполнены корректно, но не заполнено поле Ссылка на проект|Обязательные поля и ссылка заполнены корректно. Проект имеет завершенное состояние и готов к передаче в архив|Статус назначается вручную. Проект находится в работе и еще не завершен", "|")
    cL = Split("C6EFCE|FFFFFF|FFEB9C|FFC7CE|FFD966|BDD7EE|E2EFDA|D9D9D9", "|")
End Sub

Private Function FindInArray(ByRef arr() As String, ByVal count As Long, ByVal value As String) As Long
    Dim i As Long
    For i = 0 To count - 1
        If StrComp(arr(i), value, vbTextCompare) = 0 Then FindInArray = i: Exit Function
    Next i
    FindInArray = -1
End Function

Private Sub BuildMergedLegend(ByRef exSt() As String, ByRef exDs() As String, ByRef exCl() As String, ByVal hasExisting As Boolean, ByRef outSt() As String, ByRef outDs() As String, ByRef outCl() As String, ByRef outCount As Long)
    Dim dSt() As String, dDs() As String, dCl() As String
    GetDefaultLegend dSt, dDs, dCl
    ReDim outSt(0 To 50): ReDim outDs(0 To 50): ReDim outCl(0 To 50)
    Dim n As Long: n = 0
    Dim i As Long
    If hasExisting Then
        For i = LBound(exSt) To UBound(exSt)
            If Len(exSt(i)) > 0 Then
                outSt(n) = exSt(i): outDs(n) = exDs(i): outCl(n) = exCl(i): n = n + 1
            End If
        Next i
    End If
    Dim k As Long
    For k = LBound(dSt) To UBound(dSt)
        Dim idx As Long: idx = FindInArray(outSt, n, dSt(k))
        If idx < 0 Then
            outSt(n) = dSt(k): outDs(n) = dDs(k): outCl(n) = dCl(k): n = n + 1
        Else
            If Len(outDs(idx)) = 0 Then outDs(idx) = dDs(k)
            If Len(outCl(idx)) = 0 Then outCl(idx) = dCl(k)
        End If
    Next k
    outCount = n
End Sub

'===============================================================
' СПРАВОЧНИКИ
'===============================================================
Private Sub EnsureManualReferenceComplete(ByVal wb As Workbook)
    Dim lo As ListObject: Set lo = FindTableInWorkbook(wb, "тблРучныеСтатусы")
    If lo Is Nothing Then Exit Sub
    Dim hasWork As Boolean: hasWork = False
    Dim r As Long
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.DataBodyRange.Rows.count
            If StrComp(Trim(CStr(lo.DataBodyRange.Cells(r, 1).value)), "В работе", vbTextCompare) = 0 Then hasWork = True: Exit For
        Next r
    End If
    If Not hasWork Then
        Dim lr As ListRow: Set lr = lo.ListRows.Add: lr.Range(1).value = "В работе"
    End If
End Sub

Private Sub UpdateReferenceNamedRanges(ByVal wb As Workbook)
    Dim loManual As ListObject, loAllowed As ListObject
    Set loManual = FindTableInWorkbook(wb, "тблРучныеСтатусы")
    Set loAllowed = FindTableInWorkbook(wb, "тблДопустимыеСостояния")
    If Not loManual Is Nothing Then
        If Not loManual.DataBodyRange Is Nothing Then
            On Error Resume Next: wb.Names("СписокРучныхСтатусов").Delete: On Error GoTo 0
            wb.Names.Add name:="СписокРучныхСтатусов", RefersTo:="='" & loManual.DataBodyRange.Worksheet.name & "'!" & loManual.DataBodyRange.Address
        End If
    End If
    If Not loAllowed Is Nothing Then
        If Not loAllowed.DataBodyRange Is Nothing Then
            On Error Resume Next: wb.Names("ДопустимыеСостояния").Delete: On Error GoTo 0
            wb.Names.Add name:="ДопустимыеСостояния", RefersTo:="='" & loAllowed.DataBodyRange.Worksheet.name & "'!" & loAllowed.DataBodyRange.Address
        End If
    End If
End Sub

Private Function FindTableInWorkbook(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        Set lo = Nothing
        On Error Resume Next: Set lo = ws.ListObjects(tableName): On Error GoTo 0
        If Not lo Is Nothing Then Set FindTableInWorkbook = lo: Exit Function
    Next ws
    Set FindTableInWorkbook = Nothing
End Function

'===============================================================
' СОЗДАНИЕ ЛИСТОВ (вынесено из Module1)
'===============================================================
Public Sub CreateLegendSheet(ByVal ws As Worksheet, ByVal wb As Workbook, ByVal legacyFinalPath As String, ByVal legacyCurrentPath As String)
    ws.Cells.Clear
    Dim exSt() As String, exDs() As String, exCl() As String
    Dim hasExisting As Boolean: hasExisting = False
    If Len(Dir(legacyFinalPath)) > 0 Then hasExisting = ReadLegendFromFile(legacyFinalPath, exSt, exDs, exCl)
    If Not hasExisting And Len(Dir(legacyCurrentPath)) > 0 Then hasExisting = ReadLegendFromFile(legacyCurrentPath, exSt, exDs, exCl)
    Dim mSt() As String, mDs() As String, mCl() As String, mCount As Long
    BuildMergedLegend exSt, exDs, exCl, hasExisting, mSt, mDs, mCl, mCount
    ApplySavedColors mSt, mCl, mCount
    ws.Range("A1").value = "Статус проверки": ws.Range("B1").value = "Описание": ws.Range("C1").value = "Цвет (HEX)"
    Dim i As Long
    For i = 0 To mCount - 1
        ws.Cells(i + 2, 1).value = mSt(i): ws.Cells(i + 2, 2).value = mDs(i): ws.Cells(i + 2, 3).value = mCl(i)
    Next i
    Dim loLegend As ListObject
    Set loLegend = ws.ListObjects.Add(xlSrcRange, ws.Range("A1", ws.Cells(mCount + 1, 3)), , xlYes)
    loLegend.name = "тблЛегендаСтатусов": loLegend.tableStyle = "TableStyleLight13"
    PaintLegend ws: ws.Columns("A:C").AutoFit
End Sub

Public Sub CreateManualReferenceSheet(ByVal ws As Worksheet, ByVal wb As Workbook)
    ws.Cells.Clear
    ws.Range("A1").value = "Статус"
    ws.Range("A2").value = "Архивировать": ws.Range("A3").value = "Уточнить!": ws.Range("A4").value = "В работе"
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:A4"), , xlYes)
    lo.name = "тблРучныеСтатусы": lo.tableStyle = "TableStyleLight13"
    On Error Resume Next: wb.Names("СписокРучныхСтатусов").Delete: On Error GoTo 0
    wb.Names.Add name:="СписокРучныхСтатусов", RefersTo:="='СправочникСтатусов'!$A$2:$A$4"
    ws.Columns("A").AutoFit
End Sub

Public Sub CreateAllowedReferenceSheet(ByVal ws As Worksheet, ByVal wb As Workbook)
    ws.Cells.Clear
    ws.Range("A1").value = "Допустимое значение"
    ws.Range("A2").value = "Предпродажная подготовка": ws.Range("A3").value = "Инициация проекта"
    ws.Range("A4").value = "Планирование": ws.Range("A5").value = "Выполнение работ"
    ws.Range("A6").value = "Завершен": ws.Range("A7").value = "Заморожен"
    ws.Range("A8").value = "Прекращен с отрицательным результатом": ws.Range("A9").value = "Не состоялся"
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:A9"), , xlYes)
    lo.name = "тблДопустимыеСостояния": lo.tableStyle = "TableStyleLight13"
    On Error Resume Next: wb.Names("ДопустимыеСостояния").Delete: On Error GoTo 0
    wb.Names.Add name:="ДопустимыеСостояния", RefersTo:="='СправочникСостояний'!$A$2:$A$9"
    ws.Columns("A").AutoFit
End Sub

Public Sub CreateProjectsSheet(ByVal ws As Worksheet, ByVal wsLegend As Worksheet)
    ws.Cells.Clear
    Dim headersText As String
    headersText = "Контрагент|Дата создания|Код в базе 1С:УПП|Код проекта в 1С:УПП|Наименование проекта|Состояние|Тип проекта|Продукт|Группа ПГС|Группа PLM|Куратор|Менеджер ОП|Руководитель проекта|Автор"
    headersText = headersText & "|Ссылка на проект|Ответственный за статус|Статус проверки ручной|Комментарий к статусу ручной"
    headersText = headersText & "|Проверка Контрагент|Проверка Дата создания|Проверка Код|Проверка Код проекта в 1С:УПП|Проверка Наименование проекта|Проверка Состояние|Недопустимое Состояние|Проверка Куратор|Проверка Менеджер ОП"
    headersText = headersText & "|Список некорректных полей|Ошибка обязательных полей|Число ошибок|Проект закрыт|Дубль кода|Дубль ссылки|Статус проверки авто|Комментарий к статусу авто|Статус проверки|Комментарий к статусу"
    Dim headers As Variant: headers = Split(headersText, "|")
    Dim i As Long
    For i = LBound(headers) To UBound(headers): ws.Cells(1, i + 1).value = headers(i): Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(2, UBound(headers) + 1)).value = ""
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1)), , xlYes)
    lo.name = "тблПроекты": lo.tableStyle = "TableStyleLight13"
    lo.ListColumns("Код в базе 1С:УПП").Range.NumberFormatLocal = "@"
    SetAllProjectFormulas lo
    With lo.ListColumns("Статус проверки ручной").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=СписокРучныхСтатусов"
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True
        .ErrorTitle = "Недопустимый статус"
        .errorMessage = "Допустимые значения: Архивировать, Уточнить!, В работе или пустое значение."
    End With
    ApplyProjectsFormatting ws, wsLegend
End Sub

Public Sub CreateErrorsSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    Dim headersText As String: headersText = "Дата и время обработки|Тип ошибки|Ключ проекта|Код в базе 1С:УПП|Наименование проекта|Описание ошибки|Источник"
    Dim headers As Variant: headers = Split(headersText, "|")
    Dim i As Long
    For i = LBound(headers) To UBound(headers): ws.Cells(1, i + 1).value = headers(i): Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(2, UBound(headers) + 1)).value = ""
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1)), , xlYes)
    lo.name = "тблОшибки": lo.tableStyle = "TableStyleLight13"
    ws.Columns("A:G").AutoFit
End Sub

Public Sub CreateSourceSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    Dim headersText As String: headersText = "Контрагент|Дата создания|Код в базе 1С:УПП|Код проекта в 1С:УПП|Наименование проекта|Состояние|Тип проекта|Продукт|Куратор|Менеджер ОП|Руководитель проекта|Автор"
    Dim headers As Variant: headers = Split(headersText, "|")
    Dim i As Long
    For i = LBound(headers) To UBound(headers): ws.Cells(1, i + 1).value = headers(i): Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(2, UBound(headers) + 1)).value = ""
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) + 1)), , xlYes)
    lo.name = "тблВыгрузка": lo.tableStyle = "TableStyleLight13"
    ws.Columns("A:J").AutoFit
End Sub

'===============================================================
' КАРТА ПРОДУКТОВ И ГРУППЫ ПГС/PLM
'===============================================================
Public Sub BuildProductMapping(ByVal srcLo As ListObject)
    Dim wsSet As Worksheet
    Set wsSet = Nothing
    On Error Resume Next: Set wsSet = ThisWorkbook.Worksheets("Настройки"): On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub
    
    Dim keep As Collection: Set keep = New Collection
    Dim lastRow As Long: lastRow = wsSet.Cells(wsSet.Rows.count, 4).End(xlUp).row
    Dim i As Long, prod As String
    For i = 2 To lastRow
        prod = Trim(CStr(wsSet.Cells(i, 4).value))
        If Len(prod) > 0 Then
            If Not CollectionHasKey(keep, "K|" & prod) Then keep.Add CStr(wsSet.Cells(i, 5).value), "K|" & prod
        End If
    Next i
    
    Dim all As Collection: Set all = New Collection
    Dim expCol As Collection: Set expCol = CollectProductsFromExport()
    Dim vItem As Variant
    For Each vItem In expCol
        If Not CollectionHasKey(all, "K|" & CStr(vItem)) Then all.Add CStr(vItem), "K|" & CStr(vItem)
    Next vItem
    
    If Not srcLo Is Nothing Then
        If Not srcLo.DataBodyRange Is Nothing Then
            Dim colIdx As Long: colIdx = srcLo.ListColumns("Продукт").Index
            Dim r As Long, parts() As String, t As Long, pt As String
            For r = 1 To srcLo.DataBodyRange.Rows.count
                parts = Split(Trim(CStr(srcLo.DataBodyRange.Cells(r, colIdx).value)), ";")
                For t = 0 To UBound(parts)
                    pt = Trim(parts(t))
                    If Len(pt) > 0 Then
                        If Not CollectionHasKey(all, "K|" & pt) Then all.Add pt, "K|" & pt
                    End If
                Next t
            Next r
        End If
    End If
    
    Dim n As Long: n = all.count
    
    ' === ИСПРАВЛЕНИЕ: проверка на пустую коллекцию ===
    If n = 0 Then
        LogStep "BuildProductMapping: нет продуктов для обработки"
        Exit Sub
    End If
    
    Dim arr() As String: ReDim arr(1 To n)
    Dim k As Long: k = 0
    For Each vItem In all: k = k + 1: arr(k) = CStr(vItem): Next vItem
    
    Dim a As Long, b As Long, tmp As String
    For a = 1 To n - 1
        For b = a + 1 To n
            If arr(b) < arr(a) Then tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
        Next b
    Next a
    
    wsSet.Columns("D:E").ClearContents
    wsSet.Range("D1").value = "Продукт": wsSet.Range("E1").value = "Группа (ПГС / PLM)"
    wsSet.Range("D1:E1").Font.Bold = True
    For i = 1 To n
        wsSet.Cells(i + 1, 4).value = arr(i)
        If CollectionHasKey(keep, "K|" & arr(i)) Then wsSet.Cells(i + 1, 5).value = keep.item("K|" & arr(i))
    Next i
    
    If n > 0 Then
        With wsSet.Range(wsSet.Cells(2, 5), wsSet.Cells(n + 1, 5)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="ПГС,PLM"
            .IgnoreBlank = True: .InCellDropdown = True
        End With
    End If
    wsSet.Columns("D").ColumnWidth = 30: wsSet.Columns("E").ColumnWidth = 18
End Sub

Private Function CollectProductsFromExport() As Collection
    Dim res As Collection: Set res = New Collection
    Dim path As String: path = ThisWorkbook.path & "\" & "Выгрузка проектов_export.xlsx"
    If Len(Dir(path)) = 0 Then Set CollectProductsFromExport = res: Exit Function
    Dim wb As Workbook: Set wb = Nothing
    On Error Resume Next: Set wb = Workbooks.Open(fileName:=path, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False): On Error GoTo 0
    If wb Is Nothing Then Set CollectProductsFromExport = res: Exit Function
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next: Set ws = wb.Worksheets("Лист1"): On Error GoTo 0
    If Not ws Is Nothing Then
        Dim colIdx As Long: colIdx = 0
        Dim c As Long, lastCol As Long: lastCol = GetLastColumn(ws)
        For c = 1 To lastCol
            If StrComp(GetCleanText(ws.Cells(1, c).value), "Продукт", vbTextCompare) = 0 Then colIdx = c: Exit For
        Next c
        If colIdx > 0 Then
            Dim lastRow As Long: lastRow = GetLastRow(ws)
            Dim r As Long, parts() As String, t As Long, p As String
            For r = 2 To lastRow
                parts = Split(GetCleanText(ws.Cells(r, colIdx).value), ";")
                For t = 0 To UBound(parts)
                    p = Trim(parts(t))
                    If Len(p) > 0 Then
                        If Not CollectionHasKey(res, "K|" & p) Then res.Add p, "K|" & p
                    End If
                Next t
            Next r
        End If
    End If
    wb.Close SaveChanges:=False
    Set CollectProductsFromExport = res
End Function

Public Sub FillGroupColumns(ByVal lo As ListObject)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    Dim wsSet As Worksheet
    Set wsSet = Nothing
    On Error Resume Next: Set wsSet = ThisWorkbook.Worksheets("Настройки"): On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub
    
    Dim colProd As Long, colPgs As Long, colPlm As Long
    On Error Resume Next
    colProd = lo.ListColumns("Продукт").Index
    colPgs = lo.ListColumns("Группа ПГС").Index
    colPlm = lo.ListColumns("Группа PLM").Index
    On Error GoTo 0
    If colProd = 0 Or colPgs = 0 Or colPlm = 0 Then Exit Sub
    
    Dim mapPgs As Collection: Set mapPgs = New Collection
    Dim mapPlm As Collection: Set mapPlm = New Collection
    Dim lastRow As Long: lastRow = wsSet.Cells(wsSet.Rows.count, 4).End(xlUp).row
    Dim i As Long, prod As String, grp As String
    For i = 2 To lastRow
        prod = Trim(CStr(wsSet.Cells(i, 4).value)): grp = Trim(CStr(wsSet.Cells(i, 5).value))
        If Len(prod) > 0 Then
            If InStr(1, grp, "ПГС", vbTextCompare) > 0 Then
                If Not CollectionHasKey(mapPgs, "K|" & prod) Then mapPgs.Add True, "K|" & prod
            End If
            If InStr(1, grp, "PLM", vbTextCompare) > 0 Then
                If Not CollectionHasKey(mapPlm, "K|" & prod) Then mapPlm.Add True, "K|" & prod
            End If
        End If
    Next i
    
    Dim r As Long, parts() As String, t As Long, pt As String
    Dim isPgs As Boolean, isPlm As Boolean
    For r = 1 To lo.DataBodyRange.Rows.count
        isPgs = False: isPlm = False
        parts = Split(Trim(CStr(lo.DataBodyRange.Cells(r, colProd).value)), ";")
        For t = 0 To UBound(parts)
            pt = Trim(parts(t))
            If CollectionHasKey(mapPgs, "K|" & pt) Then isPgs = True
            If CollectionHasKey(mapPlm, "K|" & pt) Then isPlm = True
        Next t
        If isPgs Then lo.DataBodyRange.Cells(r, colPgs).value = "да" Else lo.DataBodyRange.Cells(r, colPgs).value = ""
        If isPlm Then lo.DataBodyRange.Cells(r, colPlm).value = "да" Else lo.DataBodyRange.Cells(r, colPlm).value = ""
    Next r
End Sub

'===============================================================
' СПРАВОЧНИК ПРОДУКТОВ В ИТОГОВОМ ФАЙЛЕ
'===============================================================
Public Sub CreateProductReferenceSheet(ByVal wb As Workbook)
    LogStep "CreateProductReferenceSheet: начало"
    Dim wsRef As Worksheet
    Set wsRef = Nothing
    On Error Resume Next: Set wsRef = wb.Worksheets("СправочникПродуктов"):
    On Error GoTo 0
    If wsRef Is Nothing Then
        Set wsRef = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        wsRef.name = "СправочникПродуктов"
        LogStep "Лист СправочникПродуктов создан"
    Else
        wsRef.Cells.Clear
        Do While wsRef.ListObjects.count > 0
            wsRef.ListObjects(1).Delete:
        Loop
        LogStep "Лист СправочникПродуктов очищен"
    End If
    wsRef.Range("A1").value = "Продукт": wsRef.Range("B1").value = "Группа"
    wsRef.Range("A1:B1").Font.Bold = True
    
    Dim wsSettings As Worksheet
    Set wsSettings = Nothing
    On Error Resume Next: Set wsSettings = ThisWorkbook.Worksheets("Настройки"): On Error GoTo 0
    If wsSettings Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = wsSettings.Cells(wsSettings.Rows.count, 4).End(xlUp).row
    If lastRow < 2 Then Exit Sub
    
    Dim r As Long, outRow As Long: outRow = 2
    For r = 2 To lastRow
        Dim prod As String: prod = Trim(CStr(wsSettings.Cells(r, 4).value))
        If Len(prod) > 0 Then
            wsRef.Cells(outRow, 1).value = prod
            wsRef.Cells(outRow, 2).value = Trim(CStr(wsSettings.Cells(r, 5).value))
            outRow = outRow + 1
        End If
    Next r
    
    If outRow > 2 Then
        Dim lo As ListObject
        Set lo = wsRef.ListObjects.Add(xlSrcRange, wsRef.Range("A1:B" & (outRow - 1)), , xlYes)
        lo.name = "тблСправочникПродуктов": lo.tableStyle = "TableStyleMedium15"
    End If
    wsRef.Columns("A").ColumnWidth = 40: wsRef.Columns("B").ColumnWidth = 20
    LogStep "CreateProductReferenceSheet: завершено"
End Sub

'===============================================================
' СТАТИСТИКА (перенесена из Module1 без изменений)
'===============================================================
Public Sub CreateStatisticsSheet(ByVal wb As Workbook, ByVal ws As Worksheet)
    ' Код CreateStatistics из Module1 (без изменений)
    ' Вставьте сюда полный код процедуры CreateStatistics из Module1
    ' (он слишком большой, чтобы дублировать здесь — используйте существующий код)
    LogStep "CreateStatisticsSheet: начало"
ws.Cells.Clear
    Do While ws.ListObjects.count > 0
        ws.ListObjects(1).Delete
    Loop
    Dim ci As Long
    For ci = ws.ChartObjects.count To 1 Step -1
        ws.ChartObjects(ci).Delete
    Next ci
    Dim h As Long
    For h = ws.Hyperlinks.count To 1 Step -1
        ws.Hyperlinks(h).Delete
    Next h

    ws.Range("A1").value = "Статистика реестра проектов"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    '================ 1. Ключевые показатели =================
    Dim rowT1 As Long
    rowT1 = 3
    ws.Cells(rowT1, 1).value = "1. Ключевые показатели качества"
    ws.Cells(rowT1, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT1, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 1. Ключевые показатели качества"

    Dim projectTypes() As String
    projectTypes = Split("Задача типовая|Пресейл|Разовая работа|Типовой проект|Уникальный проект", "|")
    Dim tCount As Long
    tCount = UBound(projectTypes) + 1

    ws.Cells(4, 1).value = "Показатель"
    Dim c As Long
    Dim cL As String
    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(4, 1 + c).value = projectTypes(c - 1)
        ws.Cells(4, 1 + c).Font.Bold = True
        ws.Columns(cL).ColumnWidth = 16
    Next c

    Dim totalCol As Long
    totalCol = 2 + tCount
    cL = colLetter(totalCol)
    ws.Cells(4, totalCol).value = "Всего"
    ws.Cells(4, totalCol).Font.Bold = True
    ws.Columns(cL).ColumnWidth = 16

    ws.Cells(5, 1).value = "Всего карточек (из БД)"
    ws.Cells(6, 1).value = "Без ошибок заполнения"
    ws.Cells(7, 1).value = "С ошибками заполнения"
    ws.Cells(8, 1).value = "Индекс качества"
    ws.Cells(9, 1).value = "Закрытые проекты (всего)"
    ws.Cells(10, 1).value = "Закрытые проекты с ошибками заполнения"
    ws.Cells(11, 1).value = "Индекс качества закрытых проектов"

    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(5, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$4)"
        ws.Cells(6, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$4,тблПроекты[Число ошибок],0)"
        ws.Cells(7, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$4,тблПроекты[Число ошибок],"">0"")"
        ws.Cells(8, 1 + c).formula = "=IF(" & cL & "5=0,""""," & cL & "6/" & cL & "5)"
        ws.Cells(9, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$4,тблПроекты[Проект закрыт],TRUE)"
        ws.Cells(10, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$4,тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
        ws.Cells(11, 1 + c).formula = "=IF(" & cL & "9=0,"""",(" & cL & "9-" & cL & "10)/" & cL & "9)"
    Next c

    cL = colLetter(totalCol)
    ws.Cells(5, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"")"
    ws.Cells(6, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Число ошибок],0)"
    ws.Cells(7, totalCol).formula = "=SUM(" & colLetter(2) & "7:" & colLetter(1 + tCount) & "7)"
    ws.Cells(8, totalCol).formula = "=IF(" & cL & "5=0,""""," & cL & "6/" & cL & "5)"
    ws.Cells(9, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проект закрыт],TRUE)"
    ws.Cells(10, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
    ws.Cells(11, totalCol).formula = "=IF(" & cL & "9=0,"""",(" & cL & "9-" & cL & "10)/" & cL & "9)"

    ws.Range(ws.Cells(8, 2), ws.Cells(8, totalCol)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(11, 2), ws.Cells(11, totalCol)).NumberFormat = "0.0%"

    With ws.Range(ws.Cells(8, 1), ws.Cells(8, totalCol))
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(11, 1), ws.Cells(11, totalCol))
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(10, 1), ws.Cells(10, totalCol)).Interior.color = RGB(255, 199, 206)
    ApplyBorders ws.Range(ws.Cells(4, 1), ws.Cells(11, totalCol))

    '================ 1-ТИМ =================
    Dim rowT2 As Long
    rowT2 = 13
    ws.Cells(rowT2, 1).value = "1-ТИМ. Ключевые показатели качества"
    ws.Cells(rowT2, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT2, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 1-ТИМ. Ключевые показатели качества"

    ws.Cells(rowT2 + 1, 1).value = "Показатель"
    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT2 + 1, 1 + c).value = projectTypes(c - 1)
        ws.Cells(rowT2 + 1, 1 + c).Font.Bold = True
    Next c
    cL = colLetter(totalCol)
    ws.Cells(rowT2 + 1, totalCol).value = "Всего"
    ws.Cells(rowT2 + 1, totalCol).Font.Bold = True

    ws.Cells(rowT2 + 2, 1).value = "Всего карточек (из БД)"
    ws.Cells(rowT2 + 3, 1).value = "Без ошибок заполнения"
    ws.Cells(rowT2 + 4, 1).value = "С ошибками заполнения"
    ws.Cells(rowT2 + 5, 1).value = "Индекс качества"
    ws.Cells(rowT2 + 6, 1).value = "Закрытые проекты (всего)"
    ws.Cells(rowT2 + 7, 1).value = "Закрытые проекты с ошибками заполнения"
    ws.Cells(rowT2 + 8, 1).value = "Индекс качества закрытых проектов"

    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT2 + 2, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT2 + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
        ws.Cells(rowT2 + 3, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT2 + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""",тблПроекты[Число ошибок],0)"
        ws.Cells(rowT2 + 4, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT2 + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""",тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT2 + 5, 1 + c).formula = "=IF(" & cL & (rowT2 + 2) & "=0,""""," & cL & (rowT2 + 3) & "/" & cL & (rowT2 + 2) & ")"
        ws.Cells(rowT2 + 6, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT2 + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE)"
        ws.Cells(rowT2 + 7, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT2 + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT2 + 8, 1 + c).formula = "=IF(" & cL & (rowT2 + 6) & "=0,"""",(" & cL & (rowT2 + 6) & "-" & cL & (rowT2 + 7) & ")/" & cL & (rowT2 + 6) & ")"
    Next c

    cL = colLetter(totalCol)
    ws.Cells(rowT2 + 2, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Cells(rowT2 + 3, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""",тблПроекты[Число ошибок],0)"
    ws.Cells(rowT2 + 4, totalCol).formula = "=SUM(" & colLetter(2) & (rowT2 + 4) & ":" & colLetter(1 + tCount) & (rowT2 + 4) & ")"
    ws.Cells(rowT2 + 5, totalCol).formula = "=IF(" & cL & (rowT2 + 2) & "=0,""""," & cL & (rowT2 + 3) & "/" & cL & (rowT2 + 2) & ")"
    ws.Cells(rowT2 + 6, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE)"
    ws.Cells(rowT2 + 7, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
    ws.Cells(rowT2 + 8, totalCol).formula = "=IF(" & cL & (rowT2 + 6) & "=0,"""",(" & cL & (rowT2 + 6) & "-" & cL & (rowT2 + 7) & ")/" & cL & (rowT2 + 6) & ")"

    ws.Range(ws.Cells(rowT2 + 5, 2), ws.Cells(rowT2 + 5, totalCol)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(rowT2 + 8, 2), ws.Cells(rowT2 + 8, totalCol)).NumberFormat = "0.0%"
    With ws.Range(ws.Cells(rowT2 + 5, 1), ws.Cells(rowT2 + 5, totalCol))
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(rowT2 + 8, 1), ws.Cells(rowT2 + 8, totalCol))
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(rowT2 + 7, 1), ws.Cells(rowT2 + 7, totalCol)).Interior.color = RGB(255, 199, 206)
    ApplyBorders ws.Range(ws.Cells(rowT2 + 1, 1), ws.Cells(rowT2 + 8, totalCol))

    '================ 1-PLM =================
    Dim rowT3 As Long
    rowT3 = rowT2 + 10
    ws.Cells(rowT3, 1).value = "1-PLM. Ключевые показатели качества"
    ws.Cells(rowT3, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT3, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 1-PLM. Ключевые показатели качества"

    ws.Cells(rowT3 + 1, 1).value = "Показатель"
    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT3 + 1, 1 + c).value = projectTypes(c - 1)
        ws.Cells(rowT3 + 1, 1 + c).Font.Bold = True
    Next c
    cL = colLetter(totalCol)
    ws.Cells(rowT3 + 1, totalCol).value = "Всего"
    ws.Cells(rowT3 + 1, totalCol).Font.Bold = True

    ws.Cells(rowT3 + 2, 1).value = "Всего карточек (из БД)"
    ws.Cells(rowT3 + 3, 1).value = "Без ошибок заполнения"
    ws.Cells(rowT3 + 4, 1).value = "С ошибками заполнения"
    ws.Cells(rowT3 + 5, 1).value = "Индекс качества"
    ws.Cells(rowT3 + 6, 1).value = "Закрытые проекты (всего)"
    ws.Cells(rowT3 + 7, 1).value = "Закрытые проекты с ошибками заполнения"
    ws.Cells(rowT3 + 8, 1).value = "Индекс качества закрытых проектов"

    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT3 + 2, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT3 + 1) & ",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
        ws.Cells(rowT3 + 3, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT3 + 1) & ",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""",тблПроекты[Число ошибок],0)"
        ws.Cells(rowT3 + 4, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT3 + 1) & ",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""",тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT3 + 5, 1 + c).formula = "=IF(" & cL & (rowT3 + 2) & "=0,""""," & cL & (rowT3 + 3) & "/" & cL & (rowT3 + 2) & ")"
        ws.Cells(rowT3 + 6, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT3 + 1) & ",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""",тблПроекты[Проект закрыт],TRUE)"
        ws.Cells(rowT3 + 7, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT3 + 1) & ",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT3 + 8, 1 + c).formula = "=IF(" & cL & (rowT3 + 6) & "=0,"""",(" & cL & (rowT3 + 6) & "-" & cL & (rowT3 + 7) & ")/" & cL & (rowT3 + 6) & ")"
    Next c

    cL = colLetter(totalCol)
    ws.Cells(rowT3 + 2, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Cells(rowT3 + 3, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""",тблПроекты[Число ошибок],0)"
    ws.Cells(rowT3 + 4, totalCol).formula = "=SUM(" & colLetter(2) & (rowT3 + 4) & ":" & colLetter(1 + tCount) & (rowT3 + 4) & ")"
    ws.Cells(rowT3 + 5, totalCol).formula = "=IF(" & cL & (rowT3 + 2) & "=0,""""," & cL & (rowT3 + 3) & "/" & cL & (rowT3 + 2) & ")"
    ws.Cells(rowT3 + 6, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""",тблПроекты[Проект закрыт],TRUE)"
    ws.Cells(rowT3 + 7, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
    ws.Cells(rowT3 + 8, totalCol).formula = "=IF(" & cL & (rowT3 + 6) & "=0,"""",(" & cL & (rowT3 + 6) & "-" & cL & (rowT3 + 7) & ")/" & cL & (rowT3 + 6) & ")"

    ws.Range(ws.Cells(rowT3 + 5, 2), ws.Cells(rowT3 + 5, totalCol)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(rowT3 + 8, 2), ws.Cells(rowT3 + 8, totalCol)).NumberFormat = "0.0%"
    With ws.Range(ws.Cells(rowT3 + 5, 1), ws.Cells(rowT3 + 5, totalCol))
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(rowT3 + 8, 1), ws.Cells(rowT3 + 8, totalCol))
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(rowT3 + 7, 1), ws.Cells(rowT3 + 7, totalCol)).Interior.color = RGB(255, 199, 206)
    ApplyBorders ws.Range(ws.Cells(rowT3 + 1, 1), ws.Cells(rowT3 + 8, totalCol))

    '================ 1 Смешанные =================
    Dim rowT4a As Long
    rowT4a = rowT3 + 10
    ws.Cells(rowT4a, 1).value = "1 Смешанные (продукты ТИМ и PLM). Ключевые показатели качества"
    ws.Cells(rowT4a, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT4a, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 1 Смешанные (продукты ТИМ и PLM). Ключевые показатели качества"

    ws.Cells(rowT4a + 1, 1).value = "Показатель"
    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT4a + 1, 1 + c).value = projectTypes(c - 1)
        ws.Cells(rowT4a + 1, 1 + c).Font.Bold = True
    Next c
    cL = colLetter(totalCol)
    ws.Cells(rowT4a + 1, totalCol).value = "Всего"
    ws.Cells(rowT4a + 1, totalCol).Font.Bold = True

    ws.Cells(rowT4a + 2, 1).value = "Всего карточек (из БД)"
    ws.Cells(rowT4a + 3, 1).value = "Без ошибок заполнения"
    ws.Cells(rowT4a + 4, 1).value = "С ошибками заполнения"
    ws.Cells(rowT4a + 5, 1).value = "Индекс качества"
    ws.Cells(rowT4a + 6, 1).value = "Закрытые проекты (всего)"
    ws.Cells(rowT4a + 7, 1).value = "Закрытые проекты с ошибками заполнения"
    ws.Cells(rowT4a + 8, 1).value = "Индекс качества закрытых проектов"

    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT4a + 2, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4a + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
        ws.Cells(rowT4a + 3, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4a + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"",тблПроекты[Число ошибок],0)"
        ws.Cells(rowT4a + 4, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4a + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"",тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT4a + 5, 1 + c).formula = "=IF(" & cL & (rowT4a + 2) & "=0,""""," & cL & (rowT4a + 3) & "/" & cL & (rowT4a + 2) & ")"
        ws.Cells(rowT4a + 6, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4a + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"",тблПроекты[Проект закрыт],TRUE)"
        ws.Cells(rowT4a + 7, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4a + 1) & ",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT4a + 8, 1 + c).formula = "=IF(" & cL & (rowT4a + 6) & "=0,"""",(" & cL & (rowT4a + 6) & "-" & cL & (rowT4a + 7) & ")/" & cL & (rowT4a + 6) & ")"
    Next c

    cL = colLetter(totalCol)
    ws.Cells(rowT4a + 2, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Cells(rowT4a + 3, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"",тблПроекты[Число ошибок],0)"
    ws.Cells(rowT4a + 4, totalCol).formula = "=SUM(" & colLetter(2) & (rowT4a + 4) & ":" & colLetter(1 + tCount) & (rowT4a + 4) & ")"
    ws.Cells(rowT4a + 5, totalCol).formula = "=IF(" & cL & (rowT4a + 2) & "=0,""""," & cL & (rowT4a + 3) & "/" & cL & (rowT4a + 2) & ")"
    ws.Cells(rowT4a + 6, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"",тблПроекты[Проект закрыт],TRUE)"
    ws.Cells(rowT4a + 7, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
    ws.Cells(rowT4a + 8, totalCol).formula = "=IF(" & cL & (rowT4a + 6) & "=0,"""",(" & cL & (rowT4a + 6) & "-" & cL & (rowT4a + 7) & ")/" & cL & (rowT4a + 6) & ")"

    ws.Range(ws.Cells(rowT4a + 5, 2), ws.Cells(rowT4a + 5, totalCol)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(rowT4a + 8, 2), ws.Cells(rowT4a + 8, totalCol)).NumberFormat = "0.0%"
    With ws.Range(ws.Cells(rowT4a + 5, 1), ws.Cells(rowT4a + 5, totalCol))
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(rowT4a + 8, 1), ws.Cells(rowT4a + 8, totalCol))
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(rowT4a + 7, 1), ws.Cells(rowT4a + 7, totalCol)).Interior.color = RGB(255, 199, 206)
    ApplyBorders ws.Range(ws.Cells(rowT4a + 1, 1), ws.Cells(rowT4a + 8, totalCol))

    '================ 1-Неопределенные =================
    Dim rowT4b As Long
    rowT4b = rowT4a + 10
    ws.Cells(rowT4b, 1).value = "1-Неопределенные. Ключевые показатели качества"
    ws.Cells(rowT4b, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT4b, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 1-Неопределенные. Ключевые показатели качества"

    ws.Cells(rowT4b + 1, 1).value = "Показатель"
    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT4b + 1, 1 + c).value = projectTypes(c - 1)
        ws.Cells(rowT4b + 1, 1 + c).Font.Bold = True
    Next c
    cL = colLetter(totalCol)
    ws.Cells(rowT4b + 1, totalCol).value = "Всего"
    ws.Cells(rowT4b + 1, totalCol).Font.Bold = True

    ws.Cells(rowT4b + 2, 1).value = "Всего карточек (из БД)"
    ws.Cells(rowT4b + 3, 1).value = "Без ошибок заполнения"
    ws.Cells(rowT4b + 4, 1).value = "С ошибками заполнения"
    ws.Cells(rowT4b + 5, 1).value = "Индекс качества"
    ws.Cells(rowT4b + 6, 1).value = "Закрытые проекты (всего)"
    ws.Cells(rowT4b + 7, 1).value = "Закрытые проекты с ошибками заполнения"
    ws.Cells(rowT4b + 8, 1).value = "Индекс качества закрытых проектов"

    For c = 1 To tCount
        cL = colLetter(1 + c)
        ws.Cells(rowT4b + 2, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4b + 1) & ",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
        ws.Cells(rowT4b + 3, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4b + 1) & ",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""",тблПроекты[Число ошибок],0)"
        ws.Cells(rowT4b + 4, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4b + 1) & ",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""",тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT4b + 5, 1 + c).formula = "=IF(" & cL & (rowT4b + 2) & "=0,""""," & cL & (rowT4b + 3) & "/" & cL & (rowT4b + 2) & ")"
        ws.Cells(rowT4b + 6, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4b + 1) & ",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE)"
        ws.Cells(rowT4b + 7, 1 + c).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Тип проекта]," & cL & "$" & (rowT4b + 1) & ",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
        ws.Cells(rowT4b + 8, 1 + c).formula = "=IF(" & cL & (rowT4b + 6) & "=0,"""",(" & cL & (rowT4b + 6) & "-" & cL & (rowT4b + 7) & ")/" & cL & (rowT4b + 6) & ")"
    Next c

    cL = colLetter(totalCol)
    ws.Cells(rowT4b + 2, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Cells(rowT4b + 3, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""",тблПроекты[Число ошибок],0)"
    ws.Cells(rowT4b + 4, totalCol).formula = "=SUM(" & colLetter(2) & (rowT4b + 4) & ":" & colLetter(1 + tCount) & (rowT4b + 4) & ")"
    ws.Cells(rowT4b + 5, totalCol).formula = "=IF(" & cL & (rowT4b + 2) & "=0,""""," & cL & (rowT4b + 3) & "/" & cL & (rowT4b + 2) & ")"
    ws.Cells(rowT4b + 6, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE)"
    ws.Cells(rowT4b + 7, totalCol).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""",тблПроекты[Проект закрыт],TRUE,тблПроекты[Число ошибок],"">0"")"
    ws.Cells(rowT4b + 8, totalCol).formula = "=IF(" & cL & (rowT4b + 6) & "=0,"""",(" & cL & (rowT4b + 6) & "-" & cL & (rowT4b + 7) & ")/" & cL & (rowT4b + 6) & ")"

    ws.Range(ws.Cells(rowT4b + 5, 2), ws.Cells(rowT4b + 5, totalCol)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(rowT4b + 8, 2), ws.Cells(rowT4b + 8, totalCol)).NumberFormat = "0.0%"
    With ws.Range(ws.Cells(rowT4b + 5, 1), ws.Cells(rowT4b + 5, totalCol))
        .Font.Bold = True
    End With
    With ws.Range(ws.Cells(rowT4b + 8, 1), ws.Cells(rowT4b + 8, totalCol))
        .Font.Bold = True
    End With
    ws.Range(ws.Cells(rowT4b + 7, 1), ws.Cells(rowT4b + 7, totalCol)).Interior.color = RGB(255, 199, 206)
    ApplyBorders ws.Range(ws.Cells(rowT4b + 1, 1), ws.Cells(rowT4b + 8, totalCol))

    '================ Оглавление =================
    Dim tocCol As Long
    tocCol = totalCol + 2
    Dim tocRow As Long
    tocRow = 3
    ws.Cells(tocRow, tocCol).value = "Оглавление"
    ws.Cells(tocRow, tocCol).Font.Bold = True
    ws.Cells(tocRow, tocCol).Font.Size = 12

    On Error Resume Next
    wb.Names("TOC").Delete
    On Error GoTo 0
    wb.Names.Add name:="TOC", RefersTo:="=" & ws.name & "!" & ws.Cells(tocRow, tocCol).Address

    Dim tocTitles(1 To 16) As String
    tocTitles(1) = "1. Ключевые показатели качества"
    tocTitles(2) = "1-ТИМ. Ключевые показатели качества"
    tocTitles(3) = "1-PLM. Ключевые показатели качества"
    tocTitles(4) = "1 Смешанные (продукты ТИМ и PLM). Ключевые показатели качества"
    tocTitles(5) = "1-Неопределенные. Ключевые показатели качества"
    tocTitles(6) = "2. Распределение карточек по числу ошибок"
    tocTitles(7) = "3. Поля: корректность и заполненность"
    tocTitles(8) = "3-ТИМ. Поля: корректность и заполненность"
    tocTitles(9) = "3-PLM. Поля: корректность и заполненность"
    tocTitles(10) = "3 Смешанные (продукты ТИМ и PLM). Поля: корректность и заполненность"
    tocTitles(11) = "3-Неопределенные. Поля: корректность и заполненность"
    tocTitles(12) = "4. Дефекты данных БД"
    tocTitles(13) = "5. Качество заполнения карточек по Авторам"
    tocTitles(14) = "6. Качество заполнения карточек по Кураторам"
    tocTitles(15) = "7. Качество заполнения карточек по Менеджерам ОП"
    tocTitles(16) = "8. Качество заполнения карточек по Руководителям проекта"

    Dim rowT5 As Long: rowT5 = rowT4b + 10
    Dim rowT6 As Long: rowT6 = rowT5 + 7
    Dim rowT6b As Long: rowT6b = rowT6 + 14
    Dim rowT6c As Long: rowT6c = rowT6b + 14
    Dim rowT6d As Long: rowT6d = rowT6c + 14
    Dim rowT6e As Long: rowT6e = rowT6d + 14
    Dim rowT7 As Long: rowT7 = rowT6e + 14
    Dim tableRows(1 To 16) As Long
    tableRows(1) = rowT1
    tableRows(2) = rowT2
    tableRows(3) = rowT3
    tableRows(4) = rowT4a
    tableRows(5) = rowT4b
    tableRows(6) = rowT5
    tableRows(7) = rowT6
    tableRows(8) = rowT6b
    tableRows(9) = rowT6c
    tableRows(10) = rowT6d
    tableRows(11) = rowT6e
    tableRows(12) = rowT7

    Dim ti As Long
    For ti = 1 To 14
        ws.Hyperlinks.Add Anchor:=ws.Cells(tocRow + ti, tocCol), Address:="", SubAddress:="=" & ws.name & "!A" & tableRows(ti), TextToDisplay:=tocTitles(ti)
    Next ti
    ws.Columns(colLetter(tocCol)).ColumnWidth = 55

    '================ 2. Распределение карточек по числу ошибок =================
    ws.Cells(rowT5, 1).value = "< 2. Распределение карточек по числу ошибок"
    ws.Cells(rowT5, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT5, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 2. Распределение карточек по числу ошибок"
    ws.Range("A" & (rowT5 + 1)).value = "Число ошибок"
    ws.Range("B" & (rowT5 + 1)).value = "Количество"
    ws.Range("A" & (rowT5 + 2)).value = "0 ошибок"
    ws.Range("B" & (rowT5 + 2)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Число ошибок],0)"
    ws.Range("A" & (rowT5 + 3)).value = "1 ошибка"
    ws.Range("B" & (rowT5 + 3)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Число ошибок],1)"
    ws.Range("A" & (rowT5 + 4)).value = "2 ошибки"
    ws.Range("B" & (rowT5 + 4)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Число ошибок],2)"
    ws.Range("A" & (rowT5 + 5)).value = "3 и более"
    ws.Range("B" & (rowT5 + 5)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Число ошибок],"">2"")"
    ApplyBorders ws.Range(ws.Cells(rowT5 + 1, 1), ws.Cells(rowT5 + 5, 2))

    '================ 3. Поля: корректность и заполненность (ОБЩАЯ) =================
    ws.Cells(rowT6, 1).value = "< 3. Поля: корректность и заполненность"
    ws.Cells(rowT6, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT6, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 3. Поля: корректность и заполненность"
    ws.Range("A" & (rowT6 + 1)).value = "Поле"
    ws.Range("B" & (rowT6 + 1)).value = "Тип"
    ws.Range("C" & (rowT6 + 1)).value = "Всего"
    ws.Range("D" & (rowT6 + 1)).value = "Корректно"
    ws.Range("E" & (rowT6 + 1)).value = "Проблемных"
    ws.Range("F" & (rowT6 + 1)).value = "% корректных"
    ws.Range("A" & (rowT6 + 2)).value = "Контрагент": ws.Range("B" & (rowT6 + 2)).value = "обяз."
    ws.Range("A" & (rowT6 + 3)).value = "Дата создания": ws.Range("B" & (rowT6 + 3)).value = "обяз."
    ws.Range("A" & (rowT6 + 4)).value = "Код в базе 1С:УПП": ws.Range("B" & (rowT6 + 4)).value = "обяз."
    ws.Range("A" & (rowT6 + 5)).value = "Код проекта в 1С:УПП": ws.Range("B" & (rowT6 + 5)).value = "рекоменд."
    ws.Range("A" & (rowT6 + 6)).value = "Наименование проекта": ws.Range("B" & (rowT6 + 6)).value = "обяз."
    ws.Range("A" & (rowT6 + 7)).value = "Состояние": ws.Range("B" & (rowT6 + 7)).value = "обяз."
    ws.Range("A" & (rowT6 + 8)).value = "Куратор": ws.Range("B" & (rowT6 + 8)).value = "обяз."
    ws.Range("A" & (rowT6 + 9)).value = "Менеджер ОП": ws.Range("B" & (rowT6 + 9)).value = "обяз."
    ws.Range("A" & (rowT6 + 10)).value = "Продукт": ws.Range("B" & (rowT6 + 10)).value = "рекоменд."
    ws.Range("A" & (rowT6 + 11)).value = "Руководитель проекта": ws.Range("B" & (rowT6 + 11)).value = "рекоменд."

    Dim r As Long
    For r = rowT6 + 2 To rowT6 + 11
        ws.Cells(r, 3).formula = "=$G$5"
        ws.Cells(r, 4).formula = "=C" & r & "-E" & r
        ws.Cells(r, 6).formula = "=IF(C" & r & "=0,"""",D" & r & "/C" & r & ")"
    Next r
    ws.Range("E" & (rowT6 + 2)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Контрагент],TRUE)"
    ws.Range("E" & (rowT6 + 3)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Дата создания],TRUE)"
    ws.Range("E" & (rowT6 + 4)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Код],TRUE)"
    ws.Range("E" & (rowT6 + 5)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Код проекта в 1С:УПП],"""")"
    ws.Range("E" & (rowT6 + 6)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Наименование проекта],TRUE)"
    ws.Range("E" & (rowT6 + 7)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Состояние],TRUE)+COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Недопустимое Состояние],TRUE)"
    ws.Range("E" & (rowT6 + 8)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Куратор],TRUE)"
    ws.Range("E" & (rowT6 + 9)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Менеджер ОП],TRUE)"
    ws.Range("E" & (rowT6 + 10)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Продукт],"""")"
    ws.Range("E" & (rowT6 + 11)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Руководитель проекта],"""")"
    ws.Range("F" & (rowT6 + 2) & ":F" & (rowT6 + 11)).NumberFormat = "0.0%"
    ApplyBorders ws.Range(ws.Cells(rowT6 + 1, 1), ws.Cells(rowT6 + 11, 6))

    '================ 3-ТИМ =================
    ws.Cells(rowT6b, 1).value = "< 3-ТИМ. Поля: корректность и заполненность"
    ws.Cells(rowT6b, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT6b, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 3-ТИМ. Поля: корректность и заполненность"
    ws.Range("A" & (rowT6b + 1)).value = "Поле"
    ws.Range("B" & (rowT6b + 1)).value = "Тип"
    ws.Range("C" & (rowT6b + 1)).value = "Всего"
    ws.Range("D" & (rowT6b + 1)).value = "Корректно"
    ws.Range("E" & (rowT6b + 1)).value = "Проблемных"
    ws.Range("F" & (rowT6b + 1)).value = "% корректных"
    ws.Range("A" & (rowT6b + 2)).value = "Контрагент": ws.Range("B" & (rowT6b + 2)).value = "обяз."
    ws.Range("A" & (rowT6b + 3)).value = "Дата создания": ws.Range("B" & (rowT6b + 3)).value = "обяз."
    ws.Range("A" & (rowT6b + 4)).value = "Код в базе 1С:УПП": ws.Range("B" & (rowT6b + 4)).value = "обяз."
    ws.Range("A" & (rowT6b + 5)).value = "Код проекта в 1С:УПП": ws.Range("B" & (rowT6b + 5)).value = "рекоменд."
    ws.Range("A" & (rowT6b + 6)).value = "Наименование проекта": ws.Range("B" & (rowT6b + 6)).value = "обяз."
    ws.Range("A" & (rowT6b + 7)).value = "Состояние": ws.Range("B" & (rowT6b + 7)).value = "обяз."
    ws.Range("A" & (rowT6b + 8)).value = "Куратор": ws.Range("B" & (rowT6b + 8)).value = "обяз."
    ws.Range("A" & (rowT6b + 9)).value = "Менеджер ОП": ws.Range("B" & (rowT6b + 9)).value = "обяз."
    ws.Range("A" & (rowT6b + 10)).value = "Продукт": ws.Range("B" & (rowT6b + 10)).value = "рекоменд."
    ws.Range("A" & (rowT6b + 11)).value = "Руководитель проекта": ws.Range("B" & (rowT6b + 11)).value = "рекоменд."
    For r = rowT6b + 2 To rowT6b + 11
        ws.Cells(r, 3).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
        ws.Cells(r, 4).formula = "=C" & r & "-E" & r
        ws.Cells(r, 6).formula = "=IF(C" & r & "=0,"""",D" & r & "/C" & r & ")"
    Next r
    ws.Range("E" & (rowT6b + 2)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Контрагент],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 3)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Дата создания],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 4)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Код],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 5)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Код проекта в 1С:УПП],"""",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 6)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Наименование проекта],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 7)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Состояние],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")+COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Недопустимое Состояние],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 8)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Куратор],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 9)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Менеджер ОП],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 10)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Продукт],"""",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6b + 11)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Руководитель проекта],"""",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],"""")"
    ws.Range("F" & (rowT6b + 2) & ":F" & (rowT6b + 11)).NumberFormat = "0.0%"
    ApplyBorders ws.Range(ws.Cells(rowT6b + 1, 1), ws.Cells(rowT6b + 11, 6))

    '================ 3-PLM =================
    ws.Cells(rowT6c, 1).value = "< 3-PLM. Поля: корректность и заполненность"
    ws.Cells(rowT6c, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT6c, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 3-PLM. Поля: корректность и заполненность"
    ws.Range("A" & (rowT6c + 1)).value = "Поле"
    ws.Range("B" & (rowT6c + 1)).value = "Тип"
    ws.Range("C" & (rowT6c + 1)).value = "Всего"
    ws.Range("D" & (rowT6c + 1)).value = "Корректно"
    ws.Range("E" & (rowT6c + 1)).value = "Проблемных"
    ws.Range("F" & (rowT6c + 1)).value = "% корректных"
    ws.Range("A" & (rowT6c + 2)).value = "Контрагент": ws.Range("B" & (rowT6c + 2)).value = "обяз."
    ws.Range("A" & (rowT6c + 3)).value = "Дата создания": ws.Range("B" & (rowT6c + 3)).value = "обяз."
    ws.Range("A" & (rowT6c + 4)).value = "Код в базе 1С:УПП": ws.Range("B" & (rowT6c + 4)).value = "обяз."
    ws.Range("A" & (rowT6c + 5)).value = "Код проекта в 1С:УПП": ws.Range("B" & (rowT6c + 5)).value = "рекоменд."
    ws.Range("A" & (rowT6c + 6)).value = "Наименование проекта": ws.Range("B" & (rowT6c + 6)).value = "обяз."
    ws.Range("A" & (rowT6c + 7)).value = "Состояние": ws.Range("B" & (rowT6c + 7)).value = "обяз."
    ws.Range("A" & (rowT6c + 8)).value = "Куратор": ws.Range("B" & (rowT6c + 8)).value = "обяз."
    ws.Range("A" & (rowT6c + 9)).value = "Менеджер ОП": ws.Range("B" & (rowT6c + 9)).value = "обяз."
    ws.Range("A" & (rowT6c + 10)).value = "Продукт": ws.Range("B" & (rowT6c + 10)).value = "рекоменд."
    ws.Range("A" & (rowT6c + 11)).value = "Руководитель проекта": ws.Range("B" & (rowT6c + 11)).value = "рекоменд."
    For r = rowT6c + 2 To rowT6c + 11
        ws.Cells(r, 3).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
        ws.Cells(r, 4).formula = "=C" & r & "-E" & r
        ws.Cells(r, 6).formula = "=IF(C" & r & "=0,"""",D" & r & "/C" & r & ")"
    Next r
    ws.Range("E" & (rowT6c + 2)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Контрагент],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 3)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Дата создания],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 4)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Код],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 5)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Код проекта в 1С:УПП],"""",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 6)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Наименование проекта],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 7)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Состояние],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")+COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Недопустимое Состояние],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 8)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Куратор],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 9)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Менеджер ОП],TRUE,тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 10)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Продукт],"""",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("E" & (rowT6c + 11)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Руководитель проекта],"""",тблПроекты[Группа PLM],""да"",тблПроекты[Группа ПГС],"""")"
    ws.Range("F" & (rowT6c + 2) & ":F" & (rowT6c + 11)).NumberFormat = "0.0%"
    ApplyBorders ws.Range(ws.Cells(rowT6c + 1, 1), ws.Cells(rowT6c + 11, 6))

    '================ 3 Смешанные =================
    ws.Cells(rowT6d, 1).value = "< 3 Смешанные (продукты ТИМ и PLM). Поля: корректность и заполненность"
    ws.Cells(rowT6d, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT6d, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 3 Смешанные (продукты ТИМ и PLM). Поля: корректность и заполненность"
    ws.Range("A" & (rowT6d + 1)).value = "Поле"
    ws.Range("B" & (rowT6d + 1)).value = "Тип"
    ws.Range("C" & (rowT6d + 1)).value = "Всего"
    ws.Range("D" & (rowT6d + 1)).value = "Корректно"
    ws.Range("E" & (rowT6d + 1)).value = "Проблемных"
    ws.Range("F" & (rowT6d + 1)).value = "% корректных"
    ws.Range("A" & (rowT6d + 2)).value = "Контрагент": ws.Range("B" & (rowT6d + 2)).value = "обяз."
    ws.Range("A" & (rowT6d + 3)).value = "Дата создания": ws.Range("B" & (rowT6d + 3)).value = "обяз."
    ws.Range("A" & (rowT6d + 4)).value = "Код в базе 1С:УПП": ws.Range("B" & (rowT6d + 4)).value = "обяз."
    ws.Range("A" & (rowT6d + 5)).value = "Код проекта в 1С:УПП": ws.Range("B" & (rowT6d + 5)).value = "рекоменд."
    ws.Range("A" & (rowT6d + 6)).value = "Наименование проекта": ws.Range("B" & (rowT6d + 6)).value = "обяз."
    ws.Range("A" & (rowT6d + 7)).value = "Состояние": ws.Range("B" & (rowT6d + 7)).value = "обяз."
    ws.Range("A" & (rowT6d + 8)).value = "Куратор": ws.Range("B" & (rowT6d + 8)).value = "обяз."
    ws.Range("A" & (rowT6d + 9)).value = "Менеджер ОП": ws.Range("B" & (rowT6d + 9)).value = "обяз."
    ws.Range("A" & (rowT6d + 10)).value = "Продукт": ws.Range("B" & (rowT6d + 10)).value = "рекоменд."
    ws.Range("A" & (rowT6d + 11)).value = "Руководитель проекта": ws.Range("B" & (rowT6d + 11)).value = "рекоменд."
    For r = rowT6d + 2 To rowT6d + 11
        ws.Cells(r, 3).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
        ws.Cells(r, 4).formula = "=C" & r & "-E" & r
        ws.Cells(r, 6).formula = "=IF(C" & r & "=0,"""",D" & r & "/C" & r & ")"
    Next r
    ws.Range("E" & (rowT6d + 2)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Контрагент],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 3)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Дата создания],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 4)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Код],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 5)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Код проекта в 1С:УПП],"""",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 6)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Наименование проекта],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 7)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Состояние],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")+COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Недопустимое Состояние],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 8)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Куратор],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 9)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Менеджер ОП],TRUE,тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 10)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Продукт],"""",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("E" & (rowT6d + 11)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Руководитель проекта],"""",тблПроекты[Группа ПГС],""да"",тблПроекты[Группа PLM],""да"")"
    ws.Range("F" & (rowT6d + 2) & ":F" & (rowT6d + 11)).NumberFormat = "0.0%"
    ApplyBorders ws.Range(ws.Cells(rowT6d + 1, 1), ws.Cells(rowT6d + 11, 6))

    '================ 3-Неопределенные =================
    ws.Cells(rowT6e, 1).value = "< 3-Неопределенные. Поля: корректность и заполненность"
    ws.Cells(rowT6e, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT6e, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 3-Неопределенные. Поля: корректность и заполненность"
    ws.Range("A" & (rowT6e + 1)).value = "Поле"
    ws.Range("B" & (rowT6e + 1)).value = "Тип"
    ws.Range("C" & (rowT6e + 1)).value = "Всего"
    ws.Range("D" & (rowT6e + 1)).value = "Корректно"
    ws.Range("E" & (rowT6e + 1)).value = "Проблемных"
    ws.Range("F" & (rowT6e + 1)).value = "% корректных"
    ws.Range("A" & (rowT6e + 2)).value = "Контрагент": ws.Range("B" & (rowT6e + 2)).value = "обяз."
    ws.Range("A" & (rowT6e + 3)).value = "Дата создания": ws.Range("B" & (rowT6e + 3)).value = "обяз."
    ws.Range("A" & (rowT6e + 4)).value = "Код в базе 1С:УПП": ws.Range("B" & (rowT6e + 4)).value = "обяз."
    ws.Range("A" & (rowT6e + 5)).value = "Код проекта в 1С:УПП": ws.Range("B" & (rowT6e + 5)).value = "рекоменд."
    ws.Range("A" & (rowT6e + 6)).value = "Наименование проекта": ws.Range("B" & (rowT6e + 6)).value = "обяз."
    ws.Range("A" & (rowT6e + 7)).value = "Состояние": ws.Range("B" & (rowT6e + 7)).value = "обяз."
    ws.Range("A" & (rowT6e + 8)).value = "Куратор": ws.Range("B" & (rowT6e + 8)).value = "обяз."
    ws.Range("A" & (rowT6e + 9)).value = "Менеджер ОП": ws.Range("B" & (rowT6e + 9)).value = "обяз."
    ws.Range("A" & (rowT6e + 10)).value = "Продукт": ws.Range("B" & (rowT6e + 10)).value = "рекоменд."
    ws.Range("A" & (rowT6e + 11)).value = "Руководитель проекта": ws.Range("B" & (rowT6e + 11)).value = "рекоменд."
    For r = rowT6e + 2 To rowT6e + 11
        ws.Cells(r, 3).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
        ws.Cells(r, 4).formula = "=C" & r & "-E" & r
        ws.Cells(r, 6).formula = "=IF(C" & r & "=0,"""",D" & r & "/C" & r & ")"
    Next r
    ws.Range("E" & (rowT6e + 2)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Контрагент],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 3)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Дата создания],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 4)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Код],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 5)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Код проекта в 1С:УПП],"""",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 6)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Наименование проекта],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 7)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Состояние],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")+COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Недопустимое Состояние],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 8)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Куратор],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 9)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Менеджер ОП],TRUE,тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 10)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Продукт],"""",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("E" & (rowT6e + 11)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Руководитель проекта],"""",тблПроекты[Группа ПГС],"""",тблПроекты[Группа PLM],"""")"
    ws.Range("F" & (rowT6e + 2) & ":F" & (rowT6e + 11)).NumberFormat = "0.0%"
    ApplyBorders ws.Range(ws.Cells(rowT6e + 1, 1), ws.Cells(rowT6e + 11, 6))

    '================ 4. Дефекты БД =================
    ws.Cells(rowT7, 1).value = "< 4. Дефекты данных БД"
    ws.Cells(rowT7, 1).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(rowT7, 1), Address:="", SubAddress:="TOC", TextToDisplay:="< 4. Дефекты данных БД"
    ws.Range("A" & (rowT7 + 1)).value = "Дефект"
    ws.Range("B" & (rowT7 + 1)).value = "Количество"
    ws.Range("A" & (rowT7 + 2)).value = "Код в УПП пуст или не начинается с ""Я-"""
    ws.Range("B" & (rowT7 + 2)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Проверка Код],TRUE)"
    ws.Range("A" & (rowT7 + 3)).value = "Код проекта в 1С:УПП не заполнен"
    ws.Range("B" & (rowT7 + 3)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Код проекта в 1С:УПП],"""")"
    ws.Range("A" & (rowT7 + 4)).value = "Дубли кода в УПП"
    ws.Range("B" & (rowT7 + 4)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Дубль кода],TRUE)"
    ws.Range("A" & (rowT7 + 5)).value = "Состояние вне допустимого списка"
    ws.Range("B" & (rowT7 + 5)).formula = "=COUNTIFS(тблПроекты[Дата создания],""<>"",тблПроекты[Недопустимое Состояние],TRUE)"
    ApplyBorders ws.Range(ws.Cells(rowT7 + 1, 1), ws.Cells(rowT7 + 5, 2))

    '================ 5-8. Качество по людям =================
    Dim nextRow As Long
    nextRow = rowT7 + 7

    Dim startR5 As Long
    startR5 = nextRow
    tableRows(13) = startR5
    FillPersonTable wb, ws, startR5, 1, "5. Качество заполнения карточек по Авторам", "Автор", nextRow

    nextRow = nextRow + 2
    Dim startR6 As Long
    startR6 = nextRow
    tableRows(14) = startR6
    FillPersonTable wb, ws, startR6, 1, "6. Качество заполнения карточек по Кураторам", "Куратор", nextRow

    nextRow = nextRow + 2
    Dim startR7 As Long
    startR7 = nextRow
    tableRows(15) = startR7
    FillPersonTable wb, ws, startR7, 1, "7. Качество заполнения карточек по Менеджерам ОП", "Менеджер ОП", nextRow

    nextRow = nextRow + 2
    Dim startR8 As Long
    startR8 = nextRow
    tableRows(16) = startR8
    FillPersonTable wb, ws, startR8, 1, "8. Качество заполнения карточек по Руководителям проекта", "Руководитель проекта", nextRow

    For ti = 13 To 16
        ws.Hyperlinks.Add Anchor:=ws.Cells(tocRow + ti, tocCol), Address:="", SubAddress:="=" & ws.name & "!A" & tableRows(ti), TextToDisplay:=tocTitles(ti)
    Next ti

    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 14
    Dim wc As Long
    For wc = 2 To totalCol
        ws.Columns(colLetter(wc)).ColumnWidth = 14
    Next wc

    '================ ДИАГРАММЫ =================
    Dim chartLeft As Double
    chartLeft = ws.Columns(totalCol + 1).Left
    Dim chartWidth As Double
    chartWidth = 550

    Dim topY3 As Double
    topY3 = ws.Rows(rowT6 + 1).Top
    Dim h3 As Double
    h3 = (ws.Rows(rowT6 + 11).Top + ws.Rows(rowT6 + 11).Height) - topY3
    Dim ch3 As ChartObject
    Set ch3 = ws.ChartObjects.Add(chartLeft, topY3, chartWidth, h3)
    With ch3.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Диаграмма 3. Проблемные значения по полям (общая)" & vbCrLf & "(обязательные - синие, рекомендуемые - оранжевые)"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
    End With
    Dim ser3 As Series
    Set ser3 = ch3.Chart.SeriesCollection.NewSeries
    ser3.name = "Проблемных"
    ser3.values = ws.Range("E" & (rowT6 + 2) & ":E" & (rowT6 + 11))
    ser3.XValues = ws.Range("A" & (rowT6 + 2) & ":A" & (rowT6 + 11))
    Dim p As Long
    For p = 1 To 10
        If ws.Cells(rowT6 + 1 + p, 2).value = "рекоменд." Then
            ser3.Points(p).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        Else
            ser3.Points(p).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
        End If
    Next p

    Dim topY3b As Double
    topY3b = ws.Rows(rowT6b + 1).Top
    Dim h3b As Double
    h3b = (ws.Rows(rowT6b + 11).Top + ws.Rows(rowT6b + 11).Height) - topY3b
    Dim ch3b As ChartObject
    Set ch3b = ws.ChartObjects.Add(chartLeft, topY3b, chartWidth, h3b)
    With ch3b.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Диаграмма 3-ТИМ. Проблемные значения по полям (только ПГС)" & vbCrLf & "(обязательные - синие, рекомендуемые - оранжевые)"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
    End With
    Dim ser3b As Series
    Set ser3b = ch3b.Chart.SeriesCollection.NewSeries
    ser3b.name = "Проблемных"
    ser3b.values = ws.Range("E" & (rowT6b + 2) & ":E" & (rowT6b + 11))
    ser3b.XValues = ws.Range("A" & (rowT6b + 2) & ":A" & (rowT6b + 11))
    For p = 1 To 10
        If ws.Cells(rowT6b + 1 + p, 2).value = "рекоменд." Then
            ser3b.Points(p).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        Else
            ser3b.Points(p).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
        End If
    Next p

    Dim topY3c As Double
    topY3c = ws.Rows(rowT6c + 1).Top
    Dim h3c As Double
    h3c = (ws.Rows(rowT6c + 11).Top + ws.Rows(rowT6c + 11).Height) - topY3c
    Dim ch3c As ChartObject
    Set ch3c = ws.ChartObjects.Add(chartLeft, topY3c, chartWidth, h3c)
    With ch3c.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Диаграмма 3-PLM. Проблемные значения по полям (только PLM)" & vbCrLf & "(обязательные - синие, рекомендуемые - оранжевые)"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
    End With
    Dim ser3c As Series
    Set ser3c = ch3c.Chart.SeriesCollection.NewSeries
    ser3c.name = "Проблемных"
    ser3c.values = ws.Range("E" & (rowT6c + 2) & ":E" & (rowT6c + 11))
    ser3c.XValues = ws.Range("A" & (rowT6c + 2) & ":A" & (rowT6c + 11))
    For p = 1 To 10
        If ws.Cells(rowT6c + 1 + p, 2).value = "рекоменд." Then
            ser3c.Points(p).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        Else
            ser3c.Points(p).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
        End If
    Next p

    Dim topY3d As Double
    topY3d = ws.Rows(rowT6d + 1).Top
    Dim h3d As Double
    h3d = (ws.Rows(rowT6d + 11).Top + ws.Rows(rowT6d + 11).Height) - topY3d
    Dim ch3d As ChartObject
    Set ch3d = ws.ChartObjects.Add(chartLeft, topY3d, chartWidth, h3d)
    With ch3d.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Диаграмма 3 Смешанные. Проблемные значения по полям (ТИМ и PLM)" & vbCrLf & "(обязательные - синие, рекомендуемые - оранжевые)"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
    End With
    Dim ser3d As Series
    Set ser3d = ch3d.Chart.SeriesCollection.NewSeries
    ser3d.name = "Проблемных"
    ser3d.values = ws.Range("E" & (rowT6d + 2) & ":E" & (rowT6d + 11))
    ser3d.XValues = ws.Range("A" & (rowT6d + 2) & ":A" & (rowT6d + 11))
    For p = 1 To 10
        If ws.Cells(rowT6d + 1 + p, 2).value = "рекоменд." Then
            ser3d.Points(p).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        Else
            ser3d.Points(p).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
        End If
    Next p

    Dim topY3e As Double
    topY3e = ws.Rows(rowT6e + 1).Top
    Dim h3e As Double
    h3e = (ws.Rows(rowT6e + 11).Top + ws.Rows(rowT6e + 11).Height) - topY3e
    Dim ch3e As ChartObject
    Set ch3e = ws.ChartObjects.Add(chartLeft, topY3e, chartWidth, h3e)
    With ch3e.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Диаграмма 3-Неопределенные. Проблемные значения по полям" & vbCrLf & "(обязательные - синие, рекомендуемые - оранжевые)"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
    End With
    Dim ser3e As Series
    Set ser3e = ch3e.Chart.SeriesCollection.NewSeries
    ser3e.name = "Проблемных"
    ser3e.values = ws.Range("E" & (rowT6e + 2) & ":E" & (rowT6e + 11))
    ser3e.XValues = ws.Range("A" & (rowT6e + 2) & ":A" & (rowT6e + 11))
    For p = 1 To 10
        If ws.Cells(rowT6e + 1 + p, 2).value = "рекоменд." Then
            ser3e.Points(p).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        Else
            ser3e.Points(p).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
        End If
    Next p
    LogStep "CreateStatisticsSheet: завершено"
End Sub


'===============================================================
' Таблица качества по человеку/полю
'===============================================================
Private Sub FillPersonTable(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal startRow As Long, ByVal baseCol As Long, ByVal title As String, ByVal fieldName As String, ByRef lastUsedRow As Long)
    ' Код FillPersonTable из Module1 (без изменений)

    lastUsedRow = startRow
    Dim lo As ListObject
    Set lo = wb.Worksheets("Проекты").ListObjects("тблПроекты")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim colIdx As Long
    colIdx = lo.ListColumns(fieldName).Index
    Dim n As Long
    n = lo.DataBodyRange.Rows.count

    Dim seen As Collection
    Set seen = New Collection
    Dim vals() As String
    ReDim vals(1 To n)
    Dim m As Long
    m = 0
    Dim i As Long
    Dim v As String
    For i = 1 To n
        v = Trim(CStr(lo.DataBodyRange.Cells(i, colIdx).value))
        If Len(v) > 0 Then
            If Not CollectionHasKey(seen, v) Then
                seen.Add True, v
                m = m + 1
                vals(m) = v
            End If
        End If
    Next i
    If m = 0 Then Exit Sub

    Dim a As Long
    Dim b As Long
    Dim t As String
    For a = 1 To m - 1
        For b = a + 1 To m
            If vals(b) < vals(a) Then
                t = vals(a): vals(a) = vals(b): vals(b) = t
            End If
        Next b
    Next a

    Dim cName As String
    cName = colLetter(baseCol)

    ws.Cells(startRow, baseCol).value = "< " & title
    ws.Cells(startRow, baseCol).Font.Bold = True
    ws.Hyperlinks.Add Anchor:=ws.Cells(startRow, baseCol), Address:="", SubAddress:="TOC", TextToDisplay:="< " & title
    ws.Cells(startRow + 1, baseCol).value = fieldName
    ws.Cells(startRow + 1, baseCol + 1).value = "Всего"
    ws.Cells(startRow + 1, baseCol + 2).value = "Без ошибок"
    ws.Cells(startRow + 1, baseCol + 3).value = "С ошибками"
    ws.Cells(startRow + 1, baseCol + 4).value = "Индекс качества"

    Dim r As Long
    For i = 1 To m
        r = startRow + 1 + i
        ws.Cells(r, baseCol).value = vals(i)
        ws.Cells(r, baseCol + 1).formula = "=COUNTIFS(тблПроекты[" & fieldName & "]," & cName & r & ",тблПроекты[Дата создания],""<>"")"
        ws.Cells(r, baseCol + 2).formula = "=COUNTIFS(тблПроекты[" & fieldName & "]," & cName & r & ",тблПроекты[Число ошибок],0)"
        ws.Cells(r, baseCol + 3).formula = "=COUNTIFS(тблПроекты[" & fieldName & "]," & cName & r & ",тблПроекты[Число ошибок],"">0"")"
        ws.Cells(r, baseCol + 4).formula = "=IF(" & colLetter(baseCol + 1) & r & "=0,""""," & colLetter(baseCol + 2) & r & "/" & colLetter(baseCol + 1) & r & ")"
    Next i

    ws.Range(ws.Cells(startRow + 2, baseCol + 4), ws.Cells(startRow + 1 + m, baseCol + 4)).NumberFormat = "0.0%"
    With ws.Range(ws.Cells(startRow + 2, baseCol + 4), ws.Cells(startRow + 1 + m, baseCol + 4))
        .Font.Bold = True
    End With

    ws.Columns(cName).ColumnWidth = 32
    ws.Columns(colLetter(baseCol + 1)).ColumnWidth = 8
    ws.Columns(colLetter(baseCol + 2)).ColumnWidth = 12
    ws.Columns(colLetter(baseCol + 3)).ColumnWidth = 12
    ws.Columns(colLetter(baseCol + 4)).ColumnWidth = 14

    ApplyBorders ws.Range(ws.Cells(startRow + 1, baseCol), ws.Cells(startRow + 1 + m, baseCol + 4))

    If m > 0 Then
        Dim tblName As String
        Select Case fieldName
            Case "Автор": tblName = "тблСтатАвторы"
            Case "Куратор": tblName = "тблСтатКураторы"
            Case "Менеджер ОП": tblName = "тблСтатМенеджеры"
            Case "Руководитель проекта": tblName = "тблСтатРП"
            Case Else: tblName = "тблСтат" & fieldName
        End Select

        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(startRow + 1, baseCol), ws.Cells(startRow + 1 + m, baseCol + 4))
        Dim loOld As ListObject
        On Error Resume Next
        Set loOld = ws.ListObjects(tblName)
        If Not loOld Is Nothing Then loOld.Delete
        On Error GoTo 0

        Dim loNew As ListObject
        Set loNew = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        loNew.name = tblName
        loNew.tableStyle = "TableStyleLight13"
        loNew.ShowAutoFilterDropDown = True
    End If
    'lastUsedRow = startRow + 1 + m

    lastUsedRow = startRow
End Sub

