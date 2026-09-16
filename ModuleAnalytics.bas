Attribute VB_Name = "ModuleAnalytics"
'===============================================================
' МОДУЛЬ АНАЛИТИКИ (ModuleAnalytics) v3
' Гибридная архитектура: VBA копирует эталон, формулы считают
'===============================================================
Option Explicit

'--- Константы ---
Private Const FILE_FINAL As String = "Проекты РЦ АСКОН_Волга в Архив.xlsx"
Private Const SHEET_ANALYTICS As String = "Аналитика"
Private Const SHEET_REFERENCE_DATA As String = "Эталон_Данные"
Private Const SHEET_PRODUCT_REF As String = "СправочникПродуктов"
Private Const TBL_REFERENCE As String = "тблЭталон"
Private Const REFERENCE_FOLDER As String = "Reference"

'===============================================================
' МАКРОС 1: Создание структуры аналитики (шаблоны без данных)
'===============================================================
Public Sub CreateAnalyticsStructure(ByVal rootFolder As String)
    
    LogStep "=== CreateAnalyticsStructure: НАЧАЛО ==="
    
    Dim finalPath As String
    finalPath = rootFolder & FILE_FINAL
    
    If Len(Dir(finalPath)) = 0 Then
        CriticalError "Проверка файлов", 0, "Не найден итоговый файл: " & finalPath
        Exit Sub
    End If
    
    If IsWorkbookOpenInApp(FILE_FINAL) Then
        CriticalError "Проверка файлов", 0, "Итоговый файл уже открыт. Закройте его и повторите запуск."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim wb As Workbook: Set wb = Nothing
    On Error GoTo CreateFail
    
    LogStep "Открытие итогового файла"
    Set wb = Workbooks.Open(fileName:=finalPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
    If wb Is Nothing Then Err.Raise vbObjectError + 30, , "Не удалось открыть: " & finalPath
    
    ' Удаление старых листов аналитики
    ' ИСПРАВЛЕНИЕ: Удалено удаление базовых справочных листов (СправочникПродуктов, Проекты и др.)
    ' чтобы не разрушать формулы в тблПроекты, ссылающиеся на тблСправочникПродуктов
    LogStep "Удаление старых листов аналитики"
    DeleteSheetIfExists wb, "_Сводная v1+"
    DeleteSheetIfExists wb, "Сводная статистика"
    DeleteSheetIfExists wb, SHEET_ANALYTICS
    DeleteSheetIfExists wb, SHEET_REFERENCE_DATA
    ' Строка удаления SHEET_PRODUCT_REF удалена - этот лист создаётся в CreateProjectFiles
    ' и не должен удаляться/пересоздаваться в CreateAnalyticsStructure
    
    ' ИСПРАВЛЕНИЕ: Удалено повторное создание СправочникПродуктов
    ' Этот лист уже гарантированно создан на этапе CreateProjectFiles
    ' Вызов ModuleBuilder.CreateProductReferenceSheet wb удалён
    
    ' Создание листа Эталон_Данные (пустой шаблон)
    LogStep "Создание листа Эталон_Данные"
    ProgressSet 60, "Создание листа эталона..."
    Dim wsRef As Worksheet
    Set wsRef = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    wsRef.name = SHEET_REFERENCE_DATA
    wsRef.Visible = xlSheetVisible
    
    ' Проверка видимости
    LogStep "Лист Эталон_Данные создан, видимость=" & wsRef.Visible & " (xlSheetVisible=" & xlSheetVisible & ")"
    
    ' Создаём пустую таблицу с заголовками
    CreateEmptyReferenceTable wsRef
    
    ' Проверка таблицы
    Dim loCheck As ListObject
    Set loCheck = Nothing
    On Error Resume Next
    Set loCheck = wsRef.ListObjects(TBL_REFERENCE)
    On Error GoTo 0
    If loCheck Is Nothing Then
        LogStep "ВНИМАНИЕ: таблица " & TBL_REFERENCE & " не создана на листе Эталон_Данные"
    Else
        LogStep "Таблица " & TBL_REFERENCE & " создана, строк=" & loCheck.ListRows.count
    End If

    ' Создание листа Аналитика (с формулами, но без данных)
    LogStep "Создание листа Аналитика"
    ProgressSet 70, "Создание листа Аналитика..."
    Dim wsAn As Worksheet
    Set wsAn = wb.Worksheets.Add(After:=wsRef)
    wsAn.name = SHEET_ANALYTICS
    wsAn.Visible = xlSheetVisible
    
    ' Создаём структуру аналитики с формулами (ссылки на пустой эталон)
    BuildAnalyticsFormulas wb, wsAn
    
    ' Пересчёт
    LogStep "Пересчёт формул"
    Application.Calculate
    
    ' Упорядочивание листов ПЕРЕД сохранением
    LogStep "Упорядочивание листов"
    ReorderSheetsInWorkbook wb
    
    ' Сохранение
    LogStep "Сохранение файла"
    wb.Save
    wb.Close SaveChanges:=False: Set wb = Nothing
    
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ProgressSet 100, "Готово!"
    
    Exit Sub

CreateFail:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not wb Is Nothing Then
        On Error Resume Next: wb.Close SaveChanges:=False: On Error GoTo 0
    End If
    CriticalError "CreateAnalyticsStructure", Err.Number, Err.description
End Sub

'===============================================================
' МАКРОС 2: Обновление структуры аналитики
'===============================================================
Public Sub UpgradeAnalyticsStructure(ByVal rootFolder As String)

    LogStep "=== UpgradeAnalyticsStructure: НАЧАЛО ==="
    
    Dim finalPath As String
    finalPath = rootFolder & FILE_FINAL
    
    If Len(Dir(finalPath)) = 0 Then
        CriticalError "Проверка файлов", 0, "Не найден итоговый файл: " & finalPath
        Exit Sub
    End If
    
    If IsWorkbookOpenInApp(FILE_FINAL) Then
        CriticalError "Проверка файлов", 0, "Итоговый файл уже открыт. Закройте его и повторите запуск."
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim wb As Workbook: Set wb = Nothing
    On Error GoTo UpgradeFail
    
    LogStep "Открытие итогового файла"
    Set wb = Workbooks.Open(fileName:=finalPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
    If wb Is Nothing Then Err.Raise vbObjectError + 30, , "Не удалось открыть: " & finalPath
    
    ' Удаление старых листов
    LogStep "Удаление старых листов аналитики"
    DeleteSheetIfExists wb, "_Сводная v1+"
    DeleteSheetIfExists wb, "Сводная статистика"
    
    ' Пересоздание структуры Аналитика
    LogStep "Пересоздание структуры Аналитика"
    RecreateAnalyticsStructure wb
    
    ' ИСПРАВЛЕНИЕ: Удалена синхронизация СправочникПродуктов
    ' Эта процедура (UpgradeAnalyticsStructure) предназначена ТОЛЬКО для обновления
    ' листов Аналитика и Эталон_Данные, не трогая базовые справочники
    ' СправочникПродуктов уже существует и не должен пересоздаваться здесь
    ' Вызов ModuleBuilder.CreateProductReferenceSheet wb удалён
    
    ' Пересчёт
    LogStep "Пересчёт формул"
    Application.Calculate
    
    ' Упорядочивание листов
    LogStep "Упорядочивание листов"
    ReorderSheetsInWorkbook wb
    
    ' Сохранение
    LogStep "Сохранение файла"
    wb.Save
    wb.Close SaveChanges:=False: Set wb = Nothing
    
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Логика и оформление аналитики обновлены." & vbCrLf & _
           "Данные сохранены.", vbInformation
    
    Exit Sub

UpgradeFail:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not wb Is Nothing Then
        On Error Resume Next: wb.Close SaveChanges:=False: On Error GoTo 0
    End If
    CriticalError "UpgradeAnalyticsStructure", Err.Number, Err.description
End Sub

'===============================================================
' МАКРОС 3: Пересчитать аналитику (заполнить данными)
'===============================================================
Public Sub RefreshAnalytics(ByVal rootFolder As String)
    InitLogging "RefreshAnalytics_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
    LogStep "=== RefreshAnalytics: НАЧАЛО ==="
    
    Dim finalPath As String
    finalPath = rootFolder & FILE_FINAL
    
    If Len(Dir(finalPath)) = 0 Then
        MsgBox "Не найден итоговый файл: " & finalPath, vbCritical
        Exit Sub
    End If
    
    If IsWorkbookOpenInApp(FILE_FINAL) Then
        MsgBox "Итоговый файл уже открыт. Закройте его и повторите запуск.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim wb As Workbook: Set wb = Nothing
    On Error GoTo RefreshFail
    
    LogStep "Открытие итогового файла"
    Set wb = Workbooks.Open(fileName:=finalPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
    If wb Is Nothing Then Err.Raise vbObjectError + 30, , "Не удалось открыть: " & finalPath
    
    ' Проверка наличия листа Эталон_Данные
    Dim wsRef As Worksheet
    Set wsRef = Nothing
    On Error Resume Next
    Set wsRef = wb.Worksheets(SHEET_REFERENCE_DATA)
    On Error GoTo RefreshFail
    
    If wsRef Is Nothing Then
        wb.Close SaveChanges:=False: Set wb = Nothing
        MsgBox "Лист 'Эталон_Данные' не найден. Сначала запустите 'Создать всё с нуля'.", vbCritical
        Exit Sub
    End If
    
    ' Поиск эталонного файла
    LogStep "Поиск эталонного файла"
    ProgressSet 10, "Поиск эталона..."
    Dim refPath As String
    refPath = FindReferenceFile(rootFolder & REFERENCE_FOLDER)
    
    If Len(refPath) > 0 Then
        LogStep "Эталон найден: " & refPath
        ProgressSet 20, "Открытие эталона..."
        
        ' Копируем данные из эталона на лист Эталон_Данные
        Dim copyOk As Boolean
        copyOk = CopyReferenceToSheet(wb, wsRef, refPath)
        
        If Not copyOk Then
            wb.Close SaveChanges:=False: Set wb = Nothing
            ProgressHide
            MsgBox "Не удалось открыть эталонный файл:" & vbCrLf & refPath & vbCrLf & vbCrLf & _
                   "Возможно, файл уже открыт в Excel или повреждён.", vbCritical
            Exit Sub
        End If
        
        ProgressSet 80, "Пересчёт формул..."
        Application.Calculate
        
        LogStep "Сохранение файла"
        wb.Save
        wb.Close SaveChanges:=False: Set wb = Nothing
        
        ProgressSet 100, "Готово!"
        
        MsgBox "Аналитика пересчитана." & vbCrLf & _
               "Эталонный файл: " & Dir(refPath) & vbCrLf & _
               "Данные скопированы на лист 'Эталон_Данные'.", vbInformation
    Else
        wb.Close SaveChanges:=False: Set wb = Nothing
        ProgressHide
        MsgBox "Эталонный файл не найден в папке Reference." & vbCrLf & _
               "Поместите эталон в папку:" & vbCrLf & _
               rootFolder & REFERENCE_FOLDER, vbExclamation
    End If
    
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Exit Sub

RefreshFail:
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not wb Is Nothing Then
        On Error Resume Next: wb.Close SaveChanges:=False: On Error GoTo 0
    End If
    CriticalError "RefreshAnalytics", Err.Number, Err.description
End Sub

'===============================================================
' СОЗДАНИЕ ПУСТОЙ ТАБЛИЦЫ ЭТАЛОНА
'===============================================================
Private Sub CreateEmptyReferenceTable(ByVal wsRef As Worksheet)
    LogStep "CreateEmptyReferenceTable: создание пустой таблицы"
    
    ' Заголовки (как в тблПроекты)
    Dim headers As Variant
    headers = Array( _
        "Контрагент", "Дата создания", "Код в базе 1С:УПП", _
        "Код проекта в 1С:УПП", "Наименование проекта", "Состояние", _
        "Тип проекта", "Продукт", "Группа ПГС", "Группа PLM", _
        "Куратор", "Менеджер ОП", "Руководитель проекта", "Автор", _
        "Число ошибок", "Проект закрыт" _
    )
    
    ' Записываем заголовки
    Dim h As Long
    For h = 0 To UBound(headers)
        wsRef.Cells(1, h + 1).value = headers(h)
    Next h
    
    ' Форматирование заголовков
    With wsRef.Range(wsRef.Cells(1, 1), wsRef.Cells(1, UBound(headers) + 1))
        .Font.Bold = True
        .Interior.color = RGB(221, 235, 247)
        .WrapText = True
    End With
    
    ' Формат даты
    wsRef.Columns(2).NumberFormatLocal = "дд.мм.гггг чч:мм:сс"
    
    ' Создаём умную таблицу (2 строки: заголовок + 1 пустая)
    Dim lo As ListObject
    Set lo = wsRef.ListObjects.Add(xlSrcRange, wsRef.Range(wsRef.Cells(1, 1), wsRef.Cells(2, UBound(headers) + 1)), , xlYes)
    lo.name = TBL_REFERENCE
    lo.tableStyle = "TableStyleLight15"
    
    ' Удаляем пустую строку данных
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Delete
    End If
    
    ' Ширина столбцов
    wsRef.Columns(1).ColumnWidth = 25
    wsRef.Columns(2).ColumnWidth = 18
    wsRef.Columns(3).ColumnWidth = 18
    wsRef.Columns(4).ColumnWidth = 20
    wsRef.Columns(5).ColumnWidth = 40
    wsRef.Columns(6).ColumnWidth = 20
    wsRef.Columns(7).ColumnWidth = 18
    wsRef.Columns(8).ColumnWidth = 25
    wsRef.Columns(9).ColumnWidth = 12
    wsRef.Columns(10).ColumnWidth = 12
    wsRef.Columns(11).ColumnWidth = 20
    wsRef.Columns(12).ColumnWidth = 20
    wsRef.Columns(13).ColumnWidth = 25
    wsRef.Columns(14).ColumnWidth = 20
    wsRef.Columns(15).ColumnWidth = 12
    wsRef.Columns(16).ColumnWidth = 12
    
    LogStep "CreateEmptyReferenceTable: завершено"
Exit Sub

CreateFail:
    LogStep "CreateEmptyReferenceTable: ОШИБКА в строке " & Erl & ", код=" & Err.Number & ", описание=" & Err.description
    Err.Raise Err.Number, , Err.description
End Sub

'===============================================================
' КОПИРОВАНИЕ ДАННЫХ ИЗ ЭТАЛОНА НА ЛИСТ "Эталон_Данные"
'===============================================================
Private Function CopyReferenceToSheet(ByVal wb As Workbook, ByVal wsRef As Worksheet, ByVal refPath As String) As Boolean
    CopyReferenceToSheet = False
    LogStep "CopyReferenceToSheet: начало"
    
    ' Очищаем лист Эталон_Данные
    wsRef.Cells.Clear
    Do While wsRef.ListObjects.count > 0
        wsRef.ListObjects(1).Delete
    Loop
    
    ' Открываем эталонный файл
    Dim wbRef As Workbook
    On Error Resume Next
    Set wbRef = Workbooks.Open(fileName:=refPath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    On Error GoTo 0
    
    If wbRef Is Nothing Then
        LogStep "Не удалось открыть эталонный файл: " & refPath
        Exit Function
    End If
    
    ' Ищем тблПроекты в эталоне
    Dim loRef As ListObject
    Dim wsRefProjects As Worksheet
    For Each wsRefProjects In wbRef.Worksheets
        On Error Resume Next
        Set loRef = wsRefProjects.ListObjects("тблПроекты")
        On Error GoTo 0
        If Not loRef Is Nothing Then Exit For
    Next wsRefProjects
    
    If loRef Is Nothing Or loRef.DataBodyRange Is Nothing Then
        LogStep "тблПроекты не найдена в эталоне"
        wbRef.Close SaveChanges:=False
        Exit Function
    End If
    
    ' Копируем данные из эталона
    Dim arrRef As Variant
    arrRef = loRef.DataBodyRange.value
    
    Dim refRows As Long: refRows = UBound(arrRef, 1)
    Dim refCols As Long: refCols = UBound(arrRef, 2)
    
    ' Заголовки
    Dim h As Long
    For h = 1 To refCols
        wsRef.Cells(1, h).value = loRef.ListColumns(h).name
    Next h
    
    ' Данные
    If refRows > 0 Then
        wsRef.Range(wsRef.Cells(2, 1), wsRef.Cells(refRows + 1, refCols)).value = arrRef
    End If
    
    ' Создаём умную таблицу
    Dim loRefLocal As ListObject
    Set loRefLocal = wsRef.ListObjects.Add(xlSrcRange, wsRef.Range(wsRef.Cells(1, 1), wsRef.Cells(refRows + 1, refCols)), , xlYes)
    loRefLocal.name = TBL_REFERENCE
    loRefLocal.tableStyle = "TableStyleLight15"
    
    ' Формат даты
    On Error Resume Next
    loRefLocal.ListColumns("Дата создания").Range.NumberFormatLocal = "дд.мм.гггг чч:мм:сс"
    On Error GoTo 0
    
    wbRef.Close SaveChanges:=False
    Set wbRef = Nothing
    
    CopyReferenceToSheet = True
    LogStep "CopyReferenceToSheet: завершено, строк=" & refRows
End Function

'===============================================================
' ПОСТРОЕНИЕ ФОРМУЛ НА ЛИСТЕ "Аналитика"
'===============================================================
Private Sub BuildAnalyticsFormulas(ByVal wb As Workbook, ByVal wsAn As Worksheet)
    LogStep "BuildAnalyticsFormulas: начало"
    
    ' Заголовок
    wsAn.Range("A1").value = "Аналитика: сравнение Эталон vs Текущий"
    wsAn.Range("A1").Font.Bold = True
    wsAn.Range("A1").Font.Size = 14
    
    Dim projectTypes As Variant
    projectTypes = Array("Задача типовая", "Пресейл", "Разовая работа", "Типовой проект", "Уникальный проект")
    
    '================ ТАБЛИЦА 1: Ключевые показатели =================
    LogStep "Таблица 1: Ключевые показатели"
    Dim rowT1 As Long: rowT1 = 3
    wsAn.Cells(rowT1, 1).value = "1. Ключевые показатели качества: Эталон vs Текущий"
    wsAn.Cells(rowT1, 1).Font.Bold = True
    
    wsAn.Cells(rowT1 + 1, 1).value = "Показатель"
    wsAn.Cells(rowT1 + 1, 2).value = "Эталон"
    wsAn.Cells(rowT1 + 1, 3).value = "Текущий"
    wsAn.Cells(rowT1 + 1, 4).value = "изм."
    wsAn.Cells(rowT1 + 1, 5).value = "изм. %"
    wsAn.Cells(rowT1 + 1, 1).Resize(1, 5).Font.Bold = True
    wsAn.Cells(rowT1 + 1, 1).Resize(1, 5).Interior.color = RGB(221, 235, 247)
    
    wsAn.Cells(rowT1 + 2, 1).value = "Всего карточек"
    wsAn.Cells(rowT1 + 2, 2).FormulaLocal = "=СЧЁТЗ(" & TBL_REFERENCE & "[Дата создания])"
    wsAn.Cells(rowT1 + 2, 3).FormulaLocal = "=СЧЁТЗ(тблПроекты[Дата создания])"
    wsAn.Cells(rowT1 + 2, 4).FormulaLocal = "=C" & (rowT1 + 2) & "-B" & (rowT1 + 2)
    wsAn.Cells(rowT1 + 2, 5).FormulaLocal = "=ЕСЛИ(B" & (rowT1 + 2) & "=0;"""";(C" & (rowT1 + 2) & "-B" & (rowT1 + 2) & ")/B" & (rowT1 + 2) & ")"
    
    wsAn.Cells(rowT1 + 3, 1).value = "Без ошибок"
    wsAn.Cells(rowT1 + 3, 2).FormulaLocal = "=СЧЁТЕСЛИ(" & TBL_REFERENCE & "[Число ошибок];0)"
    wsAn.Cells(rowT1 + 3, 3).FormulaLocal = "=СЧЁТЕСЛИ(тблПроекты[Число ошибок];0)"
    wsAn.Cells(rowT1 + 3, 4).FormulaLocal = "=C" & (rowT1 + 3) & "-B" & (rowT1 + 3)
    wsAn.Cells(rowT1 + 3, 5).FormulaLocal = "=ЕСЛИ(B" & (rowT1 + 3) & "=0;"""";(C" & (rowT1 + 3) & "-B" & (rowT1 + 3) & ")/B" & (rowT1 + 3) & ")"
    
    wsAn.Cells(rowT1 + 4, 1).value = "С ошибками"
    wsAn.Cells(rowT1 + 4, 2).FormulaLocal = "=СЧЁТЕСЛИ(" & TBL_REFERENCE & "[Число ошибок];"">0"")"
    wsAn.Cells(rowT1 + 4, 3).FormulaLocal = "=СЧЁТЕСЛИ(тблПроекты[Число ошибок];"">0"")"
    wsAn.Cells(rowT1 + 4, 4).FormulaLocal = "=C" & (rowT1 + 4) & "-B" & (rowT1 + 4)
    wsAn.Cells(rowT1 + 4, 5).FormulaLocal = "=ЕСЛИ(B" & (rowT1 + 4) & "=0;"""";(C" & (rowT1 + 4) & "-B" & (rowT1 + 4) & ")/B" & (rowT1 + 4) & ")"
    
    wsAn.Cells(rowT1 + 5, 1).value = "Индекс качества"
    wsAn.Cells(rowT1 + 5, 2).FormulaLocal = "=ЕСЛИ(B" & (rowT1 + 2) & "=0;"""";B" & (rowT1 + 3) & "/B" & (rowT1 + 2) & ")"
    wsAn.Cells(rowT1 + 5, 3).FormulaLocal = "=ЕСЛИ(C" & (rowT1 + 2) & "=0;"""";C" & (rowT1 + 3) & "/C" & (rowT1 + 2) & ")"
    wsAn.Cells(rowT1 + 5, 4).FormulaLocal = "=C" & (rowT1 + 5) & "-B" & (rowT1 + 5)
    wsAn.Cells(rowT1 + 5, 5).value = "п.п."
    
    wsAn.Cells(rowT1 + 6, 1).value = "Закрытые проекты"
    wsAn.Cells(rowT1 + 6, 2).FormulaLocal = _
        "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Состояние];""Завершен"")" & _
        "+СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Состояние];""Прекращен с отрицательным результатом"")" & _
        "+СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Состояние];""Не состоялся"")"
    wsAn.Cells(rowT1 + 6, 3).FormulaLocal = "=СЧЁТЕСЛИ(тблПроекты[Проект закрыт];ИСТИНА)"
    wsAn.Cells(rowT1 + 6, 4).FormulaLocal = "=C" & (rowT1 + 6) & "-B" & (rowT1 + 6)
    wsAn.Cells(rowT1 + 6, 5).FormulaLocal = "=ЕСЛИ(B" & (rowT1 + 6) & "=0;"""";(C" & (rowT1 + 6) & "-B" & (rowT1 + 6) & ")/B" & (rowT1 + 6) & ")"
    
    wsAn.Cells(rowT1 + 7, 1).value = "% закрытых"
    wsAn.Cells(rowT1 + 7, 2).FormulaLocal = "=ЕСЛИ(B" & (rowT1 + 2) & "=0;"""";B" & (rowT1 + 6) & "/B" & (rowT1 + 2) & ")"
    wsAn.Cells(rowT1 + 7, 3).FormulaLocal = "=ЕСЛИ(C" & (rowT1 + 2) & "=0;"""";C" & (rowT1 + 6) & "/C" & (rowT1 + 2) & ")"
    wsAn.Cells(rowT1 + 7, 4).FormulaLocal = "=C" & (rowT1 + 7) & "-B" & (rowT1 + 7)
    wsAn.Cells(rowT1 + 7, 5).value = "п.п."
    
    Dim pctRows As Variant: pctRows = Array(rowT1 + 5, rowT1 + 7)
    Dim pr As Variant
    For Each pr In pctRows
        wsAn.Cells(pr, 2).NumberFormat = "0.0%"
        wsAn.Cells(pr, 3).NumberFormat = "0.0%"
        wsAn.Cells(pr, 4).NumberFormat = "0.0%"
    Next pr
    
    Dim r As Long
    For r = 2 To 6
        wsAn.Cells(rowT1 + r, 5).NumberFormat = "0.0%"
    Next r
    
    ApplyBorders wsAn.Range(wsAn.Cells(rowT1 + 1, 1), wsAn.Cells(rowT1 + 7, 5))
    
    '================ ТАБЛИЦА 2: Динамика ошибок по атрибутам =================
    LogStep "Таблица 2: Динамика ошибок по атрибутам"
    Dim rowT2 As Long: rowT2 = rowT1 + 9
    wsAn.Cells(rowT2, 1).value = "2. Динамика ошибок по атрибутам (Эталон > Текущий)"
    wsAn.Cells(rowT2, 1).Font.Bold = True
    
    wsAn.Cells(rowT2 + 1, 1).value = "Атрибут"
    wsAn.Cells(rowT2 + 1, 2).value = "Эталон"
    wsAn.Cells(rowT2 + 1, 3).value = "Текущий"
    wsAn.Cells(rowT2 + 1, 4).value = "изм."
    wsAn.Cells(rowT2 + 1, 5).value = "изм. %"
    wsAn.Cells(rowT2 + 1, 1).Resize(1, 5).Font.Bold = True
    wsAn.Cells(rowT2 + 1, 1).Resize(1, 5).Interior.color = RGB(221, 235, 247)
    
    Dim errorFields As Variant
    errorFields = Array("Контрагент", "Код в базе 1С:УПП", "Наименование проекта", _
                        "Состояние", "Тип проекта", "Продукт", _
                        "Куратор", "Менеджер ОП", "Руководитель проекта")
    
    Dim row As Long
    For r = 0 To UBound(errorFields)
        row = rowT2 + 2 + r
        wsAn.Cells(row, 1).value = errorFields(r)
        wsAn.Cells(row, 2).FormulaLocal = "=СЧЁТЕСЛИ(" & TBL_REFERENCE & "[" & errorFields(r) & "];"""")"
        wsAn.Cells(row, 3).FormulaLocal = "=СЧЁТЕСЛИ(тблПроекты[" & errorFields(r) & "];"""")"
        wsAn.Cells(row, 4).FormulaLocal = "=C" & row & "-B" & row
        wsAn.Cells(row, 5).FormulaLocal = "=ЕСЛИ(B" & row & "=0;"""";(C" & row & "-B" & row & ")/B" & row & ")"
        wsAn.Cells(row, 5).NumberFormat = "0.0%"
    Next r
    
    ApplyBorders wsAn.Range(wsAn.Cells(rowT2 + 1, 1), wsAn.Cells(rowT2 + 1 + UBound(errorFields), 5))
    
    '================ ТАБЛИЦА 3: Динамика ошибок по группам =================
    LogStep "Таблица 3: Динамика ошибок по группам"
    Dim rowT3 As Long: rowT3 = rowT2 + 2 + UBound(errorFields) + 3
    wsAn.Cells(rowT3, 1).value = "3. Динамика ошибок по группам ТИМ/PLM (Эталон > Текущий)"
    wsAn.Cells(rowT3, 1).Font.Bold = True
    
    wsAn.Cells(rowT3 + 1, 1).value = "Группа"
    wsAn.Cells(rowT3 + 1, 2).value = "Всего (Эталон)"
    wsAn.Cells(rowT3 + 1, 3).value = "С ошибками (Эталон)"
    wsAn.Cells(rowT3 + 1, 4).value = "Всего (Текущий)"
    wsAn.Cells(rowT3 + 1, 5).value = "С ошибками (Текущий)"
    wsAn.Cells(rowT3 + 1, 6).value = "?"
    wsAn.Cells(rowT3 + 1, 7).value = "? %"
    wsAn.Cells(rowT3 + 1, 1).Resize(1, 7).Font.Bold = True
    wsAn.Cells(rowT3 + 1, 1).Resize(1, 7).Interior.color = RGB(221, 235, 247)
    
    ' ТИМ
    wsAn.Cells(rowT3 + 2, 1).value = "ТИМ"
    wsAn.Cells(rowT3 + 2, 2).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Дата создания];""<>"";" & TBL_REFERENCE & "[Группа ПГС];""да"";" & TBL_REFERENCE & "[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 2, 3).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Число ошибок];"">0"";" & TBL_REFERENCE & "[Группа ПГС];""да"";" & TBL_REFERENCE & "[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 2, 4).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Дата создания];""<>"";тблПроекты[Группа ПГС];""да"";тблПроекты[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 2, 5).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Число ошибок];"">0"";тблПроекты[Группа ПГС];""да"";тблПроекты[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 2, 6).FormulaLocal = "=E" & (rowT3 + 2) & "-C" & (rowT3 + 2)
    wsAn.Cells(rowT3 + 2, 7).FormulaLocal = "=ЕСЛИ(C" & (rowT3 + 2) & "=0;"""";(E" & (rowT3 + 2) & "-C" & (rowT3 + 2) & ")/C" & (rowT3 + 2) & ")"
    
    ' PLM
    wsAn.Cells(rowT3 + 3, 1).value = "PLM"
    wsAn.Cells(rowT3 + 3, 2).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Дата создания];""<>"";" & TBL_REFERENCE & "[Группа PLM];""да"";" & TBL_REFERENCE & "[Группа ПГС];"""")"
    wsAn.Cells(rowT3 + 3, 3).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Число ошибок];"">0"";" & TBL_REFERENCE & "[Группа PLM];""да"";" & TBL_REFERENCE & "[Группа ПГС];"""")"
    wsAn.Cells(rowT3 + 3, 4).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Дата создания];""<>"";тблПроекты[Группа PLM];""да"";тблПроекты[Группа ПГС];"""")"
    wsAn.Cells(rowT3 + 3, 5).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Число ошибок];"">0"";тблПроекты[Группа PLM];""да"";тблПроекты[Группа ПГС];"""")"
    wsAn.Cells(rowT3 + 3, 6).FormulaLocal = "=E" & (rowT3 + 3) & "-C" & (rowT3 + 3)
    wsAn.Cells(rowT3 + 3, 7).FormulaLocal = "=ЕСЛИ(C" & (rowT3 + 3) & "=0;"""";(E" & (rowT3 + 3) & "-C" & (rowT3 + 3) & ")/C" & (rowT3 + 3) & ")"
    
    ' Смешанные
    wsAn.Cells(rowT3 + 4, 1).value = "Смешанные"
    wsAn.Cells(rowT3 + 4, 2).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Дата создания];""<>"";" & TBL_REFERENCE & "[Группа ПГС];""да"";" & TBL_REFERENCE & "[Группа PLM];""да"")"
    wsAn.Cells(rowT3 + 4, 3).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Число ошибок];"">0"";" & TBL_REFERENCE & "[Группа ПГС];""да"";" & TBL_REFERENCE & "[Группа PLM];""да"")"
    wsAn.Cells(rowT3 + 4, 4).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Дата создания];""<>"";тблПроекты[Группа ПГС];""да"";тблПроекты[Группа PLM];""да"")"
    wsAn.Cells(rowT3 + 4, 5).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Число ошибок];"">0"";тблПроекты[Группа ПГС];""да"";тблПроекты[Группа PLM];""да"")"
    wsAn.Cells(rowT3 + 4, 6).FormulaLocal = "=E" & (rowT3 + 4) & "-C" & (rowT3 + 4)
    wsAn.Cells(rowT3 + 4, 7).FormulaLocal = "=ЕСЛИ(C" & (rowT3 + 4) & "=0;"""";(E" & (rowT3 + 4) & "-C" & (rowT3 + 4) & ")/C" & (rowT3 + 4) & ")"
    
    ' Неопределенные
    wsAn.Cells(rowT3 + 5, 1).value = "Неопределенные"
    wsAn.Cells(rowT3 + 5, 2).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Дата создания];""<>"";" & TBL_REFERENCE & "[Группа ПГС];"""";" & TBL_REFERENCE & "[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 5, 3).FormulaLocal = "=СЧЁТЕСЛИМН(" & TBL_REFERENCE & "[Число ошибок];"">0"";" & TBL_REFERENCE & "[Группа ПГС];"""";" & TBL_REFERENCE & "[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 5, 4).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Дата создания];""<>"";тблПроекты[Группа ПГС];"""";тблПроекты[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 5, 5).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Число ошибок];"">0"";тблПроекты[Группа ПГС];"""";тблПроекты[Группа PLM];"""")"
    wsAn.Cells(rowT3 + 5, 6).FormulaLocal = "=E" & (rowT3 + 5) & "-C" & (rowT3 + 5)
    wsAn.Cells(rowT3 + 5, 7).FormulaLocal = "=ЕСЛИ(C" & (rowT3 + 5) & "=0;"""";(E" & (rowT3 + 5) & "-C" & (rowT3 + 5) & ")/C" & (rowT3 + 5) & ")"
    
    ' Итого
    wsAn.Cells(rowT3 + 6, 1).value = "Итого"
    wsAn.Cells(rowT3 + 6, 1).Font.Bold = True
    wsAn.Cells(rowT3 + 6, 2).FormulaLocal = "=СУММ(B" & (rowT3 + 2) & ":B" & (rowT3 + 5) & ")"
    wsAn.Cells(rowT3 + 6, 3).FormulaLocal = "=СУММ(C" & (rowT3 + 2) & ":C" & (rowT3 + 5) & ")"
    wsAn.Cells(rowT3 + 6, 4).FormulaLocal = "=СУММ(D" & (rowT3 + 2) & ":D" & (rowT3 + 5) & ")"
    wsAn.Cells(rowT3 + 6, 5).FormulaLocal = "=СУММ(E" & (rowT3 + 2) & ":E" & (rowT3 + 5) & ")"
    wsAn.Cells(rowT3 + 6, 6).FormulaLocal = "=E" & (rowT3 + 6) & "-C" & (rowT3 + 6)
    wsAn.Cells(rowT3 + 6, 7).FormulaLocal = "=ЕСЛИ(C" & (rowT3 + 6) & "=0;"""";(E" & (rowT3 + 6) & "-C" & (rowT3 + 6) & ")/C" & (rowT3 + 6) & ")"
    
    For r = 2 To 6
        wsAn.Cells(rowT3 + r, 7).NumberFormat = "0.0%"
    Next r
    
    ApplyBorders wsAn.Range(wsAn.Cells(rowT3 + 1, 1), wsAn.Cells(rowT3 + 6, 7))
    
    '================ ТАБЛИЦА 4: Проекты с изменениями данных =================
    LogStep "Таблица 4: Проекты с изменениями"
    Dim rowT4 As Long: rowT4 = rowT3 + 8
    wsAn.Cells(rowT4, 1).value = "4. Проекты с изменениями данных (Эталон > Текущий, по группам)"
    wsAn.Cells(rowT4, 1).Font.Bold = True
    
    wsAn.Cells(rowT4 + 1, 1).value = "Поле"
    wsAn.Cells(rowT4 + 1, 2).value = "ТИМ"
    wsAn.Cells(rowT4 + 1, 3).value = "PLM"
    wsAn.Cells(rowT4 + 1, 4).value = "Смешанные"
    wsAn.Cells(rowT4 + 1, 5).value = "Неопределенные"
    wsAn.Cells(rowT4 + 1, 6).value = "Всего"
    wsAn.Cells(rowT4 + 1, 1).Resize(1, 6).Font.Bold = True
    wsAn.Cells(rowT4 + 1, 1).Resize(1, 6).Interior.color = RGB(221, 235, 247)
    
    Dim changeFields As Variant
    changeFields = Array("Контрагент", "Код в базе 1С:УПП", "Код проекта в 1С:УПП", _
                         "Наименование проекта", "Состояние", "Продукт", "Куратор", "Менеджер ОП", "Руководитель проекта")
    
    For r = 0 To UBound(changeFields)
        row = rowT4 + 2 + r
        wsAn.Cells(row, 1).value = changeFields(r)
        
        ' ТИМ
        wsAn.Cells(row, 2).FormulaLocal = _
            "=СУММПРОИЗВ((тблПроекты[Группа ПГС]=""да"")*(тблПроекты[Группа PLM]="""")*" & _
            "(тблПроекты[" & changeFields(r) & "]<>ЕСЛИОШИБКА(ВПР(тблПроекты[Дата создания];" & TBL_REFERENCE & "[[Дата создания]:[" & changeFields(r) & "]];2;0);""""))*" & _
            "(тблПроекты[Дата создания]<>"""")*(тблПроекты[" & changeFields(r) & "]<>""""))"
        
        ' PLM
        wsAn.Cells(row, 3).FormulaLocal = _
            "=СУММПРОИЗВ((тблПроекты[Группа PLM]=""да"")*(тблПроекты[Группа ПГС]="""")*" & _
            "(тблПроекты[" & changeFields(r) & "]<>ЕСЛИОШИБКА(ВПР(тблПроекты[Дата создания];" & TBL_REFERENCE & "[[Дата создания]:[" & changeFields(r) & "]];2;0);""""))*" & _
            "(тблПроекты[Дата создания]<>"""")*(тблПроекты[" & changeFields(r) & "]<>""""))"
        
        ' Смешанные
        wsAn.Cells(row, 4).FormulaLocal = _
            "=СУММПРОИЗВ((тблПроекты[Группа ПГС]=""да"")*(тблПроекты[Группа PLM]=""да"")*" & _
            "(тблПроекты[" & changeFields(r) & "]<>ЕСЛИОШИБКА(ВПР(тблПроекты[Дата создания];" & TBL_REFERENCE & "[[Дата создания]:[" & changeFields(r) & "]];2;0);""""))*" & _
            "(тблПроекты[Дата создания]<>"""")*(тблПроекты[" & changeFields(r) & "]<>""""))"
        
        ' Неопределенные
        wsAn.Cells(row, 5).FormulaLocal = _
            "=СУММПРОИЗВ((тблПроекты[Группа ПГС]="""")*(тблПроекты[Группа PLM]="""")*" & _
            "(тблПроекты[" & changeFields(r) & "]<>ЕСЛИОШИБКА(ВПР(тблПроекты[Дата создания];" & TBL_REFERENCE & "[[Дата создания]:[" & changeFields(r) & "]];2;0);""""))*" & _
            "(тблПроекты[Дата создания]<>"""")*(тблПроекты[" & changeFields(r) & "]<>""""))"
        
        ' Всего
        wsAn.Cells(row, 6).FormulaLocal = _
            "=СУММПРОИЗВ((тблПроекты[" & changeFields(r) & "]<>ЕСЛИОШИБКА(ВПР(тблПроекты[Дата создания];" & TBL_REFERENCE & "[[Дата создания]:[" & changeFields(r) & "]];2;0);""""))*" & _
            "(тблПроекты[Дата создания]<>"""")*(тблПроекты[" & changeFields(r) & "]<>""""))"
    Next r
    
    ApplyBorders wsAn.Range(wsAn.Cells(rowT4 + 1, 1), wsAn.Cells(rowT4 + 1 + UBound(changeFields), 6))
    
    '================ ТАБЛИЦА 5: Проекты с изменениями (по типам) =================
    LogStep "Таблица 5: Проекты с изменениями по типам"
    Dim rowT5 As Long: rowT5 = rowT4 + 2 + UBound(changeFields) + 3
    wsAn.Cells(rowT5, 1).value = "5. Проекты с изменениями (Эталон > Текущий, по типам проектов)"
    wsAn.Cells(rowT5, 1).Font.Bold = True
    
    wsAn.Cells(rowT5 + 1, 1).value = "Тип проекта"
    wsAn.Cells(rowT5 + 1, 2).value = "С изменениями"
    wsAn.Cells(rowT5 + 1, 3).value = "Без изменений"
    wsAn.Cells(rowT5 + 1, 4).value = "Всего"
    wsAn.Cells(rowT5 + 1, 5).value = "% с изменениями"
    wsAn.Cells(rowT5 + 1, 1).Resize(1, 5).Font.Bold = True
    wsAn.Cells(rowT5 + 1, 1).Resize(1, 5).Interior.color = RGB(221, 235, 247)
    
    Dim pt As Long, f As Long
    Dim formulaChanged As String
    
    On Error Resume Next  ' < ДОБАВИТЬ обработку ошибок
    
    For pt = 0 To UBound(projectTypes)
        row = rowT5 + 2 + pt
        wsAn.Cells(row, 1).value = projectTypes(pt)
        
        ' С изменениями: хотя бы одно поле изменено
        ' УПРОЩЕННАЯ ФОРМУЛА — не зависит от столбца "Есть в эталоне"
        formulaChanged = "=СУММПРОИЗВ((тблПроекты[Тип проекта]=""" & projectTypes(pt) & """)*" & _
                         "(НЕ(ЕОШИБКА(ПОИСКПОЗ(тблПроекты[Дата создания];" & TBL_REFERENCE & "[Дата создания];0)))*" & _
                         "(("
        
        For f = 0 To UBound(changeFields)
            If f > 0 Then formulaChanged = formulaChanged & "+"
            formulaChanged = formulaChanged & _
                "(тблПроекты[" & changeFields(f) & "]<>ЕСЛИОШИБКА(ВПР(тблПроекты[Дата создания];" & _
                TBL_REFERENCE & "[[Дата создания]:[" & changeFields(f) & "]];2;0);""""))*" & _
                "(тблПроекты[" & changeFields(f) & "]<>"""")"
        Next f
        
        formulaChanged = formulaChanged & ")>0)))"
        
        ' Проверяем длину формулы
        If Len(formulaChanged) > 8192 Then
            LogStep "ВНИМАНИЕ: Формула для типа '" & projectTypes(pt) & "' превышает лимит (длина=" & Len(formulaChanged) & ")"
            ' Записываем упрощённую версию
            wsAn.Cells(row, 2).value = 0  ' Заглушка
        Else
            wsAn.Cells(row, 2).FormulaLocal = formulaChanged
        End If
        
        If Err.Number <> 0 Then
            LogStep "ОШИБКА при записи формулы для типа '" & projectTypes(pt) & "': код=" & Err.Number & ", описание=" & Err.description
            Err.Clear
            wsAn.Cells(row, 2).value = 0  ' Заглушка при ошибке
        End If
        
        ' Всего проектов этого типа
        wsAn.Cells(row, 4).FormulaLocal = "=СЧЁТЕСЛИМН(тблПроекты[Тип проекта];""" & projectTypes(pt) & """)"
        
        ' Без изменений
        wsAn.Cells(row, 3).FormulaLocal = "=D" & row & "-B" & row
        
        ' % с изменениями
        wsAn.Cells(row, 5).FormulaLocal = "=ЕСЛИ(D" & row & "=0;"""";B" & row & "/D" & row & ")"
        wsAn.Cells(row, 5).NumberFormat = "0.0%"
    Next pt
    
    On Error GoTo 0  ' < Выключаем обработку ошибок
    
    ' Итого
    Dim rowT5Total As Long: rowT5Total = rowT5 + 2 + UBound(projectTypes) + 1
    wsAn.Cells(rowT5Total, 1).value = "Итого"
    wsAn.Cells(rowT5Total, 1).Font.Bold = True
    wsAn.Cells(rowT5Total, 2).FormulaLocal = "=СУММ(B" & (rowT5 + 2) & ":B" & (rowT5Total - 1) & ")"
    wsAn.Cells(rowT5Total, 3).FormulaLocal = "=СУММ(C" & (rowT5 + 2) & ":C" & (rowT5Total - 1) & ")"
    wsAn.Cells(rowT5Total, 4).FormulaLocal = "=СУММ(D" & (rowT5 + 2) & ":D" & (rowT5Total - 1) & ")"
    wsAn.Cells(rowT5Total, 5).FormulaLocal = "=ЕСЛИ(D" & rowT5Total & "=0;"""";B" & rowT5Total & "/D" & rowT5Total & ")"
    wsAn.Cells(rowT5Total, 5).NumberFormat = "0.0%"
    
    ApplyBorders wsAn.Range(wsAn.Cells(rowT5 + 1, 1), wsAn.Cells(rowT5Total, 5))
    
    '================ ДИАГРАММЫ =================
    LogStep "Создание диаграмм"
    
    ' Диаграмма 1: Ключевые показатели
    Dim chartLeft As Double: chartLeft = wsAn.Cells(1, 6).Left + 15
    Dim chartTop As Double: chartTop = wsAn.Rows(rowT1).Top
    Dim chartWidth As Double: chartWidth = 450
    Dim chartHeight As Double: chartHeight = wsAn.Rows(rowT1 + 7).Top + wsAn.Rows(rowT1 + 7).Height - chartTop + 20
    
    Dim ch1 As ChartObject
    Set ch1 = wsAn.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)
    With ch1.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Ключевые показатели: Эталон vs Текущий"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    Dim ser1a As Series
    Set ser1a = ch1.Chart.SeriesCollection.NewSeries
    ser1a.name = "Эталон"
    ser1a.values = Application.Union( _
        wsAn.Range("B" & (rowT1 + 2)), _
        wsAn.Range("B" & (rowT1 + 3)), _
        wsAn.Range("B" & (rowT1 + 4)), _
        wsAn.Range("B" & (rowT1 + 6)) _
    )
    ser1a.XValues = Application.Union( _
        wsAn.Range("A" & (rowT1 + 2)), _
        wsAn.Range("A" & (rowT1 + 3)), _
        wsAn.Range("A" & (rowT1 + 4)), _
        wsAn.Range("A" & (rowT1 + 6)) _
    )
    ser1a.Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    
    Dim ser1b As Series
    Set ser1b = ch1.Chart.SeriesCollection.NewSeries
    ser1b.name = "Текущий"
    ser1b.values = Application.Union( _
        wsAn.Range("C" & (rowT1 + 2)), _
        wsAn.Range("C" & (rowT1 + 3)), _
        wsAn.Range("C" & (rowT1 + 4)), _
        wsAn.Range("C" & (rowT1 + 6)) _
    )
    ser1b.XValues = Application.Union( _
        wsAn.Range("A" & (rowT1 + 2)), _
        wsAn.Range("A" & (rowT1 + 3)), _
        wsAn.Range("A" & (rowT1 + 4)), _
        wsAn.Range("A" & (rowT1 + 6)) _
    )
    ser1b.Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
    
    ' Диаграмма 2: Динамика ошибок по атрибутам
    chartTop = wsAn.Rows(rowT2).Top
    chartHeight = wsAn.Rows(rowT2 + 1 + UBound(errorFields)).Top + wsAn.Rows(rowT2 + 1 + UBound(errorFields)).Height - chartTop + 20
    
    Dim ch2 As ChartObject
    Set ch2 = wsAn.ChartObjects.Add(chartLeft, chartTop, 500, chartHeight)
    With ch2.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Динамика ошибок по атрибутам"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    Dim ser2a As Series
    Set ser2a = ch2.Chart.SeriesCollection.NewSeries
    ser2a.name = "Эталон"
    ser2a.values = wsAn.Range("B" & (rowT2 + 2) & ":B" & (rowT2 + 1 + UBound(errorFields)))
    ser2a.XValues = wsAn.Range("A" & (rowT2 + 2) & ":A" & (rowT2 + 1 + UBound(errorFields)))
    ser2a.Format.Fill.ForeColor.RGB = RGB(68, 114, 196)
    
    Dim ser2b As Series
    Set ser2b = ch2.Chart.SeriesCollection.NewSeries
    ser2b.name = "Текущий"
    ser2b.values = wsAn.Range("C" & (rowT2 + 2) & ":C" & (rowT2 + 1 + UBound(errorFields)))
    ser2b.XValues = wsAn.Range("A" & (rowT2 + 2) & ":A" & (rowT2 + 1 + UBound(errorFields)))
    ser2b.Format.Fill.ForeColor.RGB = RGB(112, 173, 71)
    
    ' Диаграмма 3: Динамика ошибок по группам
    chartTop = wsAn.Rows(rowT3).Top
    chartHeight = wsAn.Rows(rowT3 + 6).Top + wsAn.Rows(rowT3 + 6).Height - chartTop + 20
    
    Dim ch3 As ChartObject
    Set ch3 = wsAn.ChartObjects.Add(chartLeft, chartTop, 450, chartHeight)
    With ch3.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Динамика ошибок по группам"
        .ChartTitle.Font.Size = 10
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    Dim ser3a As Series
    Set ser3a = ch3.Chart.SeriesCollection.NewSeries
    ser3a.name = "Эталон"
    ser3a.values = wsAn.Range("C" & (rowT3 + 2) & ":C" & (rowT3 + 5))
    ser3a.XValues = wsAn.Range("A" & (rowT3 + 2) & ":A" & (rowT3 + 5))
    ser3a.Format.Line.ForeColor.RGB = RGB(68, 114, 196)
    
    Dim ser3b As Series
    Set ser3b = ch3.Chart.SeriesCollection.NewSeries
    ser3b.name = "Текущий"
    ser3b.values = wsAn.Range("E" & (rowT3 + 2) & ":E" & (rowT3 + 5))
    ser3b.XValues = wsAn.Range("A" & (rowT3 + 2) & ":A" & (rowT3 + 5))
    ser3b.Format.Line.ForeColor.RGB = RGB(112, 173, 71)
    
    ' Ширина столбцов
    wsAn.Columns("A").ColumnWidth = 35
    wsAn.Columns("B:G").ColumnWidth = 15
    
    LogStep "BuildAnalyticsFormulas: завершено"
Exit Sub

BuildFail:
    LogStep "BuildAnalyticsFormulas: ОШИБКА в строке " & Erl & ", код=" & Err.Number & ", описание=" & Err.description
    Err.Raise Err.Number, , Err.description
End Sub

'===============================================================
' ПУСТАЯ АНАЛИТИКА (если эталон не найден)
'===============================================================
Private Sub BuildEmptyAnalytics(ByVal wsAn As Worksheet, ByVal refFolderPath As String)
    wsAn.Range("A1").value = "Аналитика"
    wsAn.Range("A1").Font.Bold = True
    wsAn.Range("A1").Font.Size = 14
    
    wsAn.Range("A3").value = "Эталонный файл не найден."
    wsAn.Range("A3").Font.Italic = True
    
    wsAn.Range("A5").value = "Поместите эталонный файл (.xlsx) в папку:"
    wsAn.Range("A6").value = refFolderPath
    wsAn.Range("A6").Font.Bold = True
    
    wsAn.Columns("A").ColumnWidth = 60
End Sub

'===============================================================
' ПЕРЕСОЗДАНИЕ СТРУКТУРЫ ЛИСТА "Аналитика"
'===============================================================
Private Sub RecreateAnalyticsStructure(ByVal wb As Workbook)
    LogStep "RecreateAnalyticsStructure: начало"
    
    Dim wsAn As Worksheet
    Set wsAn = Nothing
    On Error Resume Next
    Set wsAn = wb.Worksheets(SHEET_ANALYTICS)
    On Error GoTo 0
    
    If wsAn Is Nothing Then
        LogStep "Лист Аналитика не найден — создаём"
        Set wsAn = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        wsAn.name = SHEET_ANALYTICS
        wsAn.Visible = xlSheetVisible
    Else
        ' Очищаем только объекты (диаграммы, таблицы), данные не трогаем
        Dim ci As Long
        For ci = wsAn.ChartObjects.count To 1 Step -1
            wsAn.ChartObjects(ci).Delete
        Next ci
        Do While wsAn.ListObjects.count > 0
            wsAn.ListObjects(1).Delete
        Loop
    End If
    
    LogStep "RecreateAnalyticsStructure: завершено"
End Sub

'===============================================================
' ПОИСК ЭТАЛОННОГО ФАЙЛА В ПАПКЕ REFERENCE
'===============================================================
Private Function FindReferenceFile(ByVal refFolder As String) As String
    FindReferenceFile = ""
    
    ' Создаём папку, если её нет
    If Len(Dir(refFolder, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir refFolder
        On Error GoTo 0
        LogStep "Создана папка Reference: " & refFolder
        Exit Function
    End If
    
    ' Ищем файлы xlsx в папке (исключая временные ~$)
    Dim f As String
    Dim files As Collection
    Set files = New Collection
    
    f = Dir(refFolder & "\*.xlsx")
    Do While Len(f) > 0
        If Left(f, 1) <> "~" Then files.Add refFolder & "\" & f
        f = Dir()
    Loop
    
    If files.count = 0 Then
        LogStep "В папке Reference нет файлов xlsx"
        Exit Function
    End If
    
    ' Если один файл — используем его
    ' Если несколько — берём первый (по алфавиту)
    FindReferenceFile = files(1)
    LogStep "Найден эталонный файл: " & files(1)
End Function

'===============================================================
' УДАЛЕНИЕ ЛИСТА, ЕСЛИ СУЩЕСТВУЕТ
'===============================================================
Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
        Application.DisplayAlerts = False
        On Error Resume Next
        ws.Delete
        On Error GoTo 0
        Application.DisplayAlerts = True
        LogStep "Удалён лист: " & sheetName
    End If
End Sub

'===============================================================
' ПРИМЕНЕНИЕ ГРАНИЦ (используется из ModuleHelpers)
'===============================================================
' Примечание: ApplyBorders уже есть в ModuleHelpers, но для автономности оставляем локальную копию
' Если хотите убрать дубль — удалите эту процедуру и вызывайте ModuleHelpers.ApplyBorders
Private Sub ApplyBorders(ByVal rng As Range)
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

'===============================================================
' УПОРЯДОЧИВАНИЕ ЛИСТОВ ВНУТРИ ИТОГОВОГО ФАЙЛА
'===============================================================
Private Sub ReorderSheetsInWorkbook(ByVal wb As Workbook)
    LogStep "ReorderSheetsInWorkbook: начало"
    
    Dim targetOrder As Variant
    targetOrder = Array("Проекты", "Статистика", "Аналитика", "Эталон_Данные", _
                       "Легенда", "СправочникСтатусов", "СправочникСостояний", _
                       "СправочникПродуктов", "Ошибки", "Выгрузка")
    
    Dim i As Long
    For i = 0 To UBound(targetOrder)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(targetOrder(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ' Делаем лист видимым перед перемещением
            ws.Visible = xlSheetVisible
            ' Перемещаем лист на нужную позицию
            On Error Resume Next
            ws.Move Before:=wb.Worksheets(i + 1)
            On Error GoTo 0
            LogStep "Лист '" & targetOrder(i) & "' перемещён на позицию " & (i + 1)
        Else
            LogStep "Лист '" & targetOrder(i) & "' не найден — пропуск"
        End If
    Next i
    
    LogStep "ReorderSheetsInWorkbook: завершено"
End Sub
