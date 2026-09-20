Attribute VB_Name = "ModuleProjects"
'===============================================================
' МОДУЛЬ РАБОТЫ С ПРОЕКТАМИ (ModuleProjects)
' бизнес-логика обновления проектов
'===============================================================
Option Explicit


'===============================================================
' СОЗДАНИЕ ФАЙЛОВ РЕШЕНИЯ
'===============================================================
Public Sub CreateProjectFiles(ByVal rootFolder As String)
    LogStep "НАЧАЛО: CreateProjectFiles"
    
    Dim finalFilePath As String, currentFilePath As String
    finalFilePath = rootFolder & "Проекты РЦ АСКОН_Волга в Архив.xlsx"
    currentFilePath = rootFolder & "Выгрузка проектов_current.xlsx"
    
    LogStep "Проверка файлов: " & finalFilePath
    
    If IsWorkbookOpenInApp("Проекты РЦ АСКОН_Волга в Архив.xlsx") Then
        Err.Raise vbObjectError + 1, , "Итоговый файл уже открыт"
    End If
    If IsWorkbookOpenInApp("Выгрузка проектов_current.xlsx") Then
        Err.Raise vbObjectError + 2, , "Служебный файл уже открыт"
    End If
    
    If Not ConfirmOverwrite(finalFilePath) Then Exit Sub
    If Not ConfirmOverwrite(currentFilePath) Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wb As Workbook: Set wb = Workbooks.Add
    
    LogStep "Создание книги: " & wb.Name
    
    ' Создаём листы в строгом порядке: Проекты, Статистика, Аналитика, Эталон_Данные, Легенда, СправочникСтатусов, СправочникСостояний, СправочникПродуктов, Ошибки, Выгрузка
    Dim wsProjects As Worksheet, wsStat As Worksheet, wsAnalytics As Worksheet
    Dim wsRefData As Worksheet, wsLegend As Worksheet, wsManualRef As Worksheet
    Dim wsAllowedRef As Worksheet, wsProductRef As Worksheet, wsErrors As Worksheet, wsSource As Worksheet
    
    Set wsProjects = wb.Sheets(1): wsProjects.name = "Проекты"
    LogStep "Лист создан: " & wsProjects.name
    
    Set wsStat = wb.Sheets.Add(After:=wsProjects): wsStat.name = "Статистика"
    LogStep "Лист создан: " & wsStat.name
    
    Set wsAnalytics = wb.Sheets.Add(After:=wsStat): wsAnalytics.name = "Аналитика"
    LogStep "Лист создан: " & wsAnalytics.name
    
    Set wsRefData = wb.Sheets.Add(After:=wsAnalytics): wsRefData.name = "Эталон_Данные"
    LogStep "Лист создан: " & wsRefData.name
    
    Set wsLegend = wb.Sheets.Add(After:=wsRefData): wsLegend.name = "Легенда"
    LogStep "Лист создан: " & wsLegend.name
    
    Set wsManualRef = wb.Sheets.Add(After:=wsLegend): wsManualRef.name = "СправочникСтатусов"
    LogStep "Лист создан: " & wsManualRef.name
    
    Set wsAllowedRef = wb.Sheets.Add(After:=wsManualRef): wsAllowedRef.name = "СправочникСостояний"
    LogStep "Лист создан: " & wsAllowedRef.name
    
    Set wsProductRef = wb.Sheets.Add(After:=wsAllowedRef): wsProductRef.name = "СправочникПродуктов"
    LogStep "Лист создан: " & wsProductRef.name
    
    Set wsErrors = wb.Sheets.Add(After:=wsProductRef): wsErrors.name = "Ошибки"
    LogStep "Лист создан: " & wsErrors.name
    
    Set wsSource = wb.Sheets.Add(After:=wsErrors): wsSource.name = "Выгрузка"
    LogStep "Лист создан: " & wsSource.name
    
    ' === ИЗМЕНЕНИЕ: Сначала создаём все справочные таблицы, особенно тблСправочникПродуктов ===
    ' Это необходимо, так как формулы в тблПроекты ссылаются на тблСправочникПродуктов
    LogStep "Создание таблиц справочников"
    CreateLegendSheet wsLegend, wb, finalFilePath, currentFilePath
    CreateManualReferenceSheet wsManualRef, wb
    CreateAllowedReferenceSheet wsAllowedRef, wb
    CreateProductReferenceSheet wb  ' Создаётся ДО CreateProjectsSheet
    LogStep "Таблица тблСправочникПродуктов создана на листе СправочникПродуктов"
    CreateErrorsSheet wsErrors
    CreateSourceSheet wsSource
    CreateStatisticsSheet wb, wsStat
    LogStep "Таблица тблСтатистика создана на листе Статистика"
    
    ' === ИЗМЕНЕНИЕ: Теперь создаём тблПроекты после всех справочников ===
    LogStep "Создание таблицы тблПроекты на листе Проекты"
    CreateProjectsSheet wsProjects, wsLegend
    LogStep "Таблица тблПроекты создана"
    
    ' Применяем стили ко всем таблицам
    Dim wsAny As Worksheet, loAny As ListObject
    For Each wsAny In wb.Worksheets
        For Each loAny In wsAny.ListObjects
            loAny.tableStyle = "TableStyleLight13"
        Next loAny
    Next wsAny
    
    wsProjects.ListObjects("тблПроекты").ListColumns("Дата создания").Range.NumberFormatLocal = "дд.мм.гггг чч:мм:сс"
    
    LogStep "Операция с файловой системой: Kill " & finalFilePath
    If Not DeleteFileIfExists(finalFilePath) Then
        Err.Raise vbObjectError + 3, , "Не удалось удалить: " & finalFilePath
    End If
    LogStep "Операция с файловой системой: Kill " & currentFilePath
    If Not DeleteFileIfExists(currentFilePath) Then
        Err.Raise vbObjectError + 4, , "Не удалось удалить: " & currentFilePath
    End If
    
    LogStep "Операция с файловой системой: SaveAs " & finalFilePath
    wb.SaveAs fileName:=finalFilePath, FileFormat:=xlOpenXMLWorkbook
    LogStep "Операция с файловой системой: SaveCopyAs " & currentFilePath
    wb.SaveCopyAs currentFilePath
    wb.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    LogStep "ЗАВЕРШЕНО: CreateProjectFiles"
End Sub

'===============================================================
' ОБНОВЛЕНИЕ ПРОЕКТОВ ИЗ ВЫГРУЗКИ
'===============================================================
Public Sub UpdateProjectsFromExport(ByVal rootFolder As String)
    ' === ИНИЦИАЛИЗАЦИЯ gStep для корректного логирования ошибок ===
    gStep = "Инициализация"
    LogStep "UpdateProjectsFromExport: начало"
    
    Dim sourcePath As String, finalPath As String, currentPath As String
    sourcePath = rootFolder & "Выгрузка проектов_export.xlsx"
    finalPath = rootFolder & "Проекты РЦ АСКОН_Волга в Архив.xlsx"
    currentPath = rootFolder & "Выгрузка проектов_current.xlsx"
    
    gStep = "Проверка файлов"
    If Len(Dir(sourcePath)) = 0 Then
        MsgBox "Файл выгрузки не найден:" & vbCrLf & sourcePath & vbCrLf & vbCrLf & _
               "Поместите файл ""Выгрузка проектов_export.xlsx"" в рабочую папку и повторите запуск.", vbExclamation
        Exit Sub
    End If
    If IsWorkbookOpenInApp("Выгрузка проектов_export.xlsx") Or _
       IsWorkbookOpenInApp("Проекты РЦ АСКОН_Волга в Архив.xlsx") Or _
       IsWorkbookOpenInApp("Выгрузка проектов_current.xlsx") Then
        Err.Raise vbObjectError + 11, , "Закройте файлы выгрузки и итоговые файлы"
    End If
    
    EnsureFolder rootFolder & "Archive_export"
    EnsureFolder rootFolder & "Archive_backup"
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    
    Dim wbSource As Workbook, wbTarget As Workbook
    Set wbSource = Nothing: Set wbTarget = Nothing
    On Error GoTo UpdateFail
    
    gStep = "Открытие файла выгрузки"
    ProgressSet 5, "Открытие файла выгрузки..."
    Set wbSource = Workbooks.Open(fileName:=sourcePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    
    gStep = "Поиск листа с данными"
    ProgressSet 10, "Поиск листа с данными..."
    Dim wsSource As Worksheet
    Set wsSource = FindSheetByHeaders(wbSource)
    If wsSource Is Nothing Then
        Dim wsInfo As Worksheet, diag As String
        diag = "Не удалось найти лист с нужными заголовками." & vbCrLf & vbCrLf & "Доступные листы:" & vbCrLf
        For Each wsInfo In wbSource.Worksheets: diag = diag & "  - '" & wsInfo.name & "'" & vbCrLf: Next wsInfo
        wbSource.Close SaveChanges:=False: Set wbSource = Nothing
        Err.Raise vbObjectError + 12, , diag
    End If
    
    gStep = "Проверка заголовков"
    ProgressSet 15, "Проверка заголовков..."
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastRow(wsSource): lastCol = GetLastColumn(wsSource)
    Dim colIdx As Variant, missingHeaders As String
    ReDim colIdx(1 To 12)
    If Not GetSourceColumnIndexes(wsSource, lastCol, colIdx, missingHeaders) Then
        wbSource.Close SaveChanges:=False: Set wbSource = Nothing
        Err.Raise vbObjectError + 13, , "Отсутствуют столбцы:" & vbCrLf & missingHeaders
    End If
    
    gStep = "Чтение данных выгрузки..."
    ProgressSet 20, "Чтение данных выгрузки..."
    On Error Resume Next
    wsSource.Columns(colIdx(3)).ColumnWidth = 40
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось изменить ширину столбца " & colIdx(3) & ". Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo UpdateFail
    
    Dim dataRange As Range
    Set dataRange = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, lastCol))
    Dim srcValues As Variant: srcValues = dataRange.value
    Set dataRange = Nothing
    wbSource.Close SaveChanges:=False: Set wbSource = Nothing
    
    Dim rowCount As Long: rowCount = lastRow - 1
    
    gStep = "Анализ данных..."
    ProgressSet 25, "Анализ данных..."
    Dim srcKey() As String, srcDate() As Date, srcCode() As String
    Dim srcName() As String, srcValid() As Boolean, srcProcessed() As Boolean
    ReDim srcKey(1 To rowCount), srcDate(1 To rowCount), srcCode(1 To rowCount)
    ReDim srcName(1 To rowCount), srcValid(1 To rowCount), srcProcessed(1 To rowCount)
    
    Dim errorsCol As Collection, keyFirst As Collection, dupKeys As Collection
    Set errorsCol = New Collection: Set keyFirst = New Collection: Set dupKeys = New Collection
    
    Dim r As Long, parsedDate As Variant, dt As Date
    Dim codeText As String, nameText As String, key As String
    
    ' Адаптивный шаг для прогресс-бара
    Dim stepSize As Long
    stepSize = Application.WorksheetFunction.RoundUp(rowCount / 50, 0)
    If stepSize < 1 Then stepSize = 1
    Dim nextUpdate As Long: nextUpdate = stepSize
    
    ' === ИЗМЕНЕНИЕ: Добавлено логирование каждые 100 итераций в цикле анализа ===
    For r = 1 To rowCount
        If r >= nextUpdate Or r = rowCount Then
            ProgressSmooth 25 + (CDbl(r) / rowCount) * 20, "Анализ строк (" & r & " из " & rowCount & ")..."
            nextUpdate = r + stepSize
        End If
        
        ' Логирование каждые 100 итераций
        If r Mod 100 = 0 Then
            gStep = "Анализ данных: строка " & r
            LogStep gStep
        End If
        
        On Error Resume Next
        srcValid(r) = False: srcProcessed(r) = False
        If IsEmptySourceRow(srcValues, r, colIdx) Then GoTo NextSourceRow
        
        codeText = GetCodeText(srcValues(r, colIdx(3)))
        nameText = GetCleanText(srcValues(r, colIdx(5)))
        parsedDate = ParseRuDateTime(srcValues(r, colIdx(2)))
        
        If Err.Number <> 0 Then
            LogStep "Ошибка в цикле анализа: строка " & r & ", ключ=" & key & ", ошибка: " & Err.description
            Err.Clear
            GoTo NextSourceRow
        End If
        On Error GoTo UpdateFail
        
        If IsEmpty(parsedDate) Then
            AddErrorToCollection errorsCol, "Ошибка даты", "", codeText, nameText, "Пустая или некорректная дата", "Выгрузка проектов_export.xlsx"
            GoTo NextSourceRow
        End If
        If VarType(parsedDate) = vbString Then
            If parsedDate = "EXCLUDED_DATE" Then
                AddErrorToCollection errorsCol, "Исключенная дата", "", codeText, nameText, "Служебная дата 01.01.0001", "Выгрузка проектов_export.xlsx"
                GoTo NextSourceRow
            End If
        End If
        
        dt = CDate(parsedDate)
        key = NormalizeDateKey(dt)
        srcKey(r) = key: srcDate(r) = dt: srcCode(r) = codeText: srcName(r) = nameText
        srcValid(r) = True
        
        If CollectionHasKey(keyFirst, key) Then
            If Not CollectionHasKey(dupKeys, key) Then dupKeys.Add True, key
        Else
            keyFirst.Add r, key
        End If
NextSourceRow:
    Next r
    
    gStep = "Проверка дублей..."
    ProgressSet 50, "Проверка дублей..."
    For r = 1 To rowCount
        If srcValid(r) Then
            If CollectionHasKey(dupKeys, srcKey(r)) Then
                AddErrorToCollection errorsCol, "Дубль ключа", srcKey(r), srcCode(r), srcName(r), "Дубль даты создания", "Выгрузка проектов_export.xlsx"
                srcProcessed(r) = True
            End If
        End If
    Next r
    
    Dim validUniqueCount As Long: validUniqueCount = 0
    For r = 1 To rowCount
        If srcValid(r) And Not CollectionHasKey(dupKeys, srcKey(r)) Then validUniqueCount = validUniqueCount + 1
    Next r
    
    If validUniqueCount = 0 Then
        ProgressHide
        Application.Calculation = oldCalc
        Application.ScreenUpdating = True: Application.EnableEvents = True: Application.DisplayAlerts = True
        MsgBox "В выгрузке нет корректных строк. Файл не перемещен.", vbExclamation
        Exit Sub
    End If
    
    gStep = "Открытие целевого файла..."
    ProgressSet 55, "Открытие целевого файла..."
    Dim targetPath As String
    If Len(Dir(finalPath)) > 0 Then
        targetPath = finalPath
    ElseIf Len(Dir(currentPath)) > 0 Then
        targetPath = currentPath
    Else
        Err.Raise vbObjectError + 14, , "Не найдены итоговые файлы. Запустите создание файлов."
    End If
    Set wbTarget = Workbooks.Open(fileName:=targetPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
    If wbTarget Is Nothing Then Err.Raise vbObjectError + 15, , "Не удалось открыть целевой файл"
    If wbTarget.ReadOnly Then
        wbTarget.Close SaveChanges:=False: Set wbTarget = Nothing
        Err.Raise vbObjectError + 16, , "Целевой файл только для чтения"
    End If
    
    gStep = "Создание резервной копии..."
    ProgressSet 60, "Создание резервной копии..."
    If targetPath = finalPath Then
        Dim backupFolder As String
        backupFolder = CreateTimestampFolder(rootFolder & "Archive_backup")
        wbTarget.SaveCopyAs GetUniqueFilePath(backupFolder, "Проекты РЦ АСКОН_Волга в Архив.xlsx")
    End If
    
    Dim wsProjects As Worksheet
    Set wsProjects = wbTarget.Worksheets("Проекты")
    Dim loProjects As ListObject
    Set loProjects = wsProjects.ListObjects("тблПроекты")
    
    loProjects.ListColumns("Код в базе 1С:УПП").Range.NumberFormatLocal = "@"
    Application.Calculate
    
    gStep = "Чтение существующих ключей..."
    ProgressSet 65, "Чтение существующих ключей..."
    Dim existingKeys As Collection: Set existingKeys = New Collection
    Dim lr As ListRow, existingKey As String
    If Not loProjects.DataBodyRange Is Nothing Then
        For Each lr In loProjects.ListRows
            existingKey = NormalizeDateKey(lr.Range(loProjects.ListColumns("Дата создания").Index).value)
            If Len(existingKey) > 0 Then
                If Not CollectionHasKey(existingKeys, existingKey) Then existingKeys.Add lr.Index, existingKey
            End If
        Next lr
    End If
    
    gStep = "Обновление существующих проектов..."
    ProgressSet 70, "Обновление существующих проектов..."
    Dim updatedCount As Long, addedCount As Long, targetRowIndex As Long
    updatedCount = 0: addedCount = 0
    nextUpdate = stepSize
    
    ' === ИЗМЕНЕНИЕ: Добавлено логирование и обработка ошибок в цикле обновления ===
    For r = 1 To rowCount
        If r >= nextUpdate Or r = rowCount Then
            ProgressSmooth 70 + (CDbl(r) / rowCount) * 10, "Обновление (" & r & " из " & rowCount & ")..."
            nextUpdate = r + stepSize
        End If
        
        ' Логирование каждые 100 итераций
        If r Mod 100 = 0 Then
            gStep = "Обновление существующих проектов: строка " & r
            LogStep gStep
        End If
        
        On Error Resume Next
        If srcValid(r) And Not srcProcessed(r) Then
            key = srcKey(r)
            If CollectionHasKey(existingKeys, key) Then
                targetRowIndex = CLng(existingKeys(key))
                If targetRowIndex >= 1 And targetRowIndex <= loProjects.ListRows.count Then
                    Set lr = loProjects.ListRows(targetRowIndex)
                    Dim writeCode As String: writeCode = srcCode(r)
                    If Len(writeCode) = 0 Then
                        writeCode = GetCleanText(lr.Range(loProjects.ListColumns("Код в базе 1С:УПП").Index).value)
                    End If
                    WriteSourceRowToListRow loProjects, lr, srcValues, r, colIdx, srcDate(r), writeCode
                    srcProcessed(r) = True: updatedCount = updatedCount + 1
                End If
            End If
        End If
        
        If Err.Number <> 0 Then
            LogStep "Ошибка в цикле обновления: строка " & r & ", ключ=" & key & ", ошибка: " & Err.description
            Err.Clear
            On Error GoTo UpdateFail
            GoTo ContinueUpdateLoop
        End If
        On Error GoTo UpdateFail
        
ContinueUpdateLoop:
    Next r
    
    ' Защита от массового дублирования
    gStep = "Защита от массового дублирования..."
    Dim existingRows As Long, realDataRows As Long
    existingRows = 0: realDataRows = 0
    If Not loProjects.DataBodyRange Is Nothing Then
        existingRows = loProjects.DataBodyRange.Rows.count
        For Each lr In loProjects.ListRows
            If Not IsEmpty(lr.Range(loProjects.ListColumns("Дата создания").Index).value) Then realDataRows = realDataRows + 1
        Next lr
    End If
    If realDataRows > 0 And updatedCount = 0 And validUniqueCount > 0 Then
        wbTarget.Close SaveChanges:=False: Set wbTarget = Nothing
        ProgressHide
        Application.Calculation = oldCalc
        Application.ScreenUpdating = True: Application.EnableEvents = True: Application.DisplayAlerts = True
        MsgBox "В таблице есть строки (" & existingRows & "), но ни одна строка выгрузки не совпала. Добавление отменено.", vbCritical
        Exit Sub
    End If
    
    gStep = "Добавление новых проектов..."
    ProgressSet 85, "Добавление новых проектов..."
    nextUpdate = stepSize
    
    ' === ИЗМЕНЕНИЕ: Добавлено логирование и обработка ошибок в цикле добавления ===
    For r = 1 To rowCount
        If r >= nextUpdate Or r = rowCount Then
            ProgressSmooth 85 + (CDbl(r) / rowCount) * 7, "Добавление (" & r & " из " & rowCount & ")..."
            nextUpdate = r + stepSize
        End If
        
        ' Логирование каждые 100 итераций
        If r Mod 100 = 0 Then
            gStep = "Добавление новых проектов: строка " & r
            LogStep gStep
        End If
        
        On Error Resume Next
        If srcValid(r) And Not srcProcessed(r) Then
            Set lr = loProjects.ListRows.Add
            WriteSourceRowToListRow loProjects, lr, srcValues, r, colIdx, srcDate(r), srcCode(r)
            srcProcessed(r) = True: addedCount = addedCount + 1
        End If
        
        If Err.Number <> 0 Then
            LogStep "Ошибка в цикле добавления: строка " & r & ", ключ=" & key & ", ошибка: " & Err.description
            Err.Clear
            On Error GoTo UpdateFail
            GoTo ContinueAddLoop
        End If
        On Error GoTo UpdateFail
        
ContinueAddLoop:
    Next r
    
    ' === ИЗМЕНЕНИЕ: Принудительное применение формул после добавления строк ===
    ' Обеспечиваем применение формул в столбцах FORMULA_COLUMNS таблицы тблПроекты
    gStep = "Применение формул в таблице"
    LogStep "Принудительный пересчет формул для тблПроекты"
    
    ' Включаем авто-расширение диапазона списка для корректного применения Calculated Columns
    On Error Resume Next
    Application.AutoCorrect.AutoExpandListRange = True
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось установить AutoExpandListRange. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo UpdateFail
    
    ' Принудительно вызываем пересчет формул для всей таблицы
    On Error Resume Next
    loProjects.Range.Calculate
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось пересчитать формулы таблицы. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo UpdateFail
    
    ' Дополнительный полный пересчет для гарантии применения формул
    On Error Resume Next
    Application.CalculateFull
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось выполнить CalculateFull. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo UpdateFail
    
    ' === ИЗМЕНЕНИЕ: Удалены вызовы BuildProductMapping и FillGroupColumns ===
    ' Столбцы "Группа ПГС" и "Группа PLM" теперь заполняются ТОЛЬКО формулами через SetAllProjectFormulas
    ' Механизм Calculated Column автоматически распространяет формулы на все строки таблицы
    ' Вызовы BuildProductMapping/FillGroupColumns записывали данные через .Value, что нарушало принцип формульных столбцов
    
    gStep = "Запись ошибок"
    ProgressSet 96, "Запись ошибок..."
    WriteErrorsToWorkbook wbTarget, errorsCol
    
    gStep = "Сохранение результата"
    ProgressSet 98, "Сохранение..."
    If targetPath = finalPath Then
        wbTarget.Save
    Else
        wbTarget.SaveAs fileName:=finalPath, FileFormat:=xlOpenXMLWorkbook
    End If
    
    On Error Resume Next
    DeleteFileIfExists currentPath
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось удалить currentPath. Ошибка: " & Err.description
        Err.Clear
    End If
    wbTarget.SaveCopyAs currentPath
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось сохранить копию currentPath. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo UpdateFail
    
    wbTarget.Close SaveChanges:=False: Set wbTarget = Nothing
    
    gStep = "Перемещение выгрузки в архив"
    ProgressSet 100, "Перемещение выгрузки в архив..."
    Dim archiveFolder As String, archiveSourcePath As String, moveOk As Boolean
    archiveFolder = CreateTimestampFolder(rootFolder & "Archive_export")
    archiveSourcePath = GetUniqueFilePath(archiveFolder, "Выгрузка проектов_export.xlsx")
    moveOk = False
    On Error Resume Next
    Name sourcePath As archiveSourcePath
    If Err.Number = 0 Then moveOk = True
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось переместить файл выгрузки в архив. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo UpdateFail
    
    ProgressHide
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True: Application.EnableEvents = True: Application.DisplayAlerts = True
    
    Dim resultMessage As String
    resultMessage = "Обработка завершена." & vbCrLf & vbCrLf
    resultMessage = resultMessage & "Обновлено: " & updatedCount & vbCrLf
    resultMessage = resultMessage & "Добавлено: " & addedCount & vbCrLf
    resultMessage = resultMessage & "Ошибок: " & errorsCol.count & vbCrLf & vbCrLf
    If moveOk Then
        resultMessage = resultMessage & "Файл выгрузки перемещен в:" & vbCrLf & archiveFolder
    Else
        resultMessage = resultMessage & "Внимание: файл выгрузки не перемещен."
    End If
    MsgBox resultMessage, vbInformation
    
    Exit Sub
    
UpdateFail:
    LogError gStep, Err.Number, Err.description
    EndLogging
    ProgressHide
    
    ' Корректное освобождение объектов даже при ошибке
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось закрыть wbSource. Ошибка: " & Err.description
        Err.Clear
    End If
    If Not wbTarget Is Nothing Then wbTarget.Close SaveChanges:=False
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось закрыть wbTarget. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo UpdateFail
    
    ' Явное освобождение ссылок на объекты
    Set wbSource = Nothing
    Set wbTarget = Nothing
    
    ' Восстановление настроек приложения
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    
    MsgBox "Ошибка обновления." & vbCrLf & "Шаг: " & gStep & vbCrLf & "Код: " & Err.Number & vbCrLf & Err.description, vbCritical
End Sub

'===============================================================
' ОБНОВЛЕНИЕ ЛОГИКИ ПРОЕКТОВ
'===============================================================
Public Sub UpgradeProjectsLogic(ByVal rootFolder As String)
    LogStep "UpgradeProjectsLogic: начало"
    
    Dim finalPath As String: finalPath = rootFolder & "Проекты РЦ АСКОН_Волга в Архив.xlsx"
    Dim wb As Workbook
    Set wb = Workbooks.Open(fileName:=finalPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
    
    Dim backupFolder As String
    backupFolder = CreateTimestampFolder(rootFolder & "Archive_backup")
    wb.SaveCopyAs GetUniqueFilePath(backupFolder, wb.name)
    
    Dim wsProjects As Worksheet: Set wsProjects = wb.Worksheets("Проекты")
    Dim wsLegend As Worksheet: Set wsLegend = wb.Worksheets("Легенда")
    Dim loProjects As ListObject: Set loProjects = wsProjects.ListObjects("тблПроекты")
    
    ' Удаление старых столбцов
    RemoveOldColumns loProjects
    
    ' Добавление/проверка столбцов
    EnsureColumnAfter loProjects, "Код проекта в 1С:УПП", "Код в базе 1С:УПП"
    EnsureColumnAfter loProjects, "Группа ПГС", "Продукт"
    EnsureColumnAfter loProjects, "Группа PLM", "Группа ПГС"
    EnsureColumnAfter loProjects, "Тип проекта", "Состояние"
    EnsureColumnAfter loProjects, "Проверка Код проекта в 1С:УПП", "Проверка Код"
    EnsureColumnAfter loProjects, "Число ошибок", "Ошибка обязательных полей"
    EnsureColumnAfter loProjects, "Проект закрыт", "Число ошибок"
    EnsureColumnAfter loProjects, "Есть в эталоне", "Дата создания"
    
    ' Проверка "Дубль ссылки"
    Dim lcLink As ListColumn
    Set lcLink = Nothing
    On Error Resume Next
    Set lcLink = loProjects.ListColumns("Дубль ссылки")
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось проверить наличие столбца 'Дубль ссылки'. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
    If lcLink Is Nothing Then
        Dim posLink As Long: posLink = loProjects.ListColumns("Дубль кода").Index + 1
        Set lcLink = loProjects.ListColumns.Add(Position:=posLink)
        lcLink.name = "Дубль ссылки"
    End If
    
    loProjects.ListColumns("Код в базе 1С:УПП").Range.NumberFormatLocal = "@"
    
    ' Оформление
    ApplyProjectsFormatting wsProjects, wsLegend
    
    ' Формулы
    If loProjects.DataBodyRange Is Nothing Then
        Dim tmpRow As ListRow: Set tmpRow = loProjects.ListRows.Add
        SetAllProjectFormulas loProjects
        tmpRow.Delete
    Else
        SetAllProjectFormulas loProjects
    End If
    
    ' === ИЗМЕНЕНИЕ: Удалены вызовы BuildProductMapping и FillGroupColumns ===
    ' Столбцы "Группа ПГС" и "Группа PLM" теперь заполняются ТОЛЬКО формулами через SetAllProjectFormulas
    ' Механизм Calculated Column автоматически распространяет формулы на все строки таблицы
    
    ' Статистика
    Dim wsStat As Worksheet
    Set wsStat = Nothing
    On Error Resume Next
    Set wsStat = wb.Worksheets("Статистика")
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось получить доступ к листу 'Статистика'. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
    If wsStat Is Nothing Then
        Set wsStat = wb.Worksheets.Add(After:=wsProjects)
        wsStat.name = "Статистика"
    End If
    LogStep "Вызов обновления листа Статистика..."
    CreateStatisticsSheet wb, wsStat
    
    ' Сохранение цветов
    wb.Save
    Dim legSt() As String, legDs() As String, legCl() As String
    If ReadLegendFromWorkbook(wb, legSt, legDs, legCl) Then
        Dim cnt As Long: cnt = 0
        Dim i2 As Long
        For i2 = LBound(legSt) To UBound(legSt)
            If Len(legSt(i2)) > 0 Then cnt = cnt + 1
        Next i2
        SaveColorsToSettings legSt, legCl, cnt
    End If
    
    Dim currentPath As String: currentPath = rootFolder & "Выгрузка проектов_current.xlsx"
    On Error Resume Next
    DeleteFileIfExists currentPath
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось удалить currentPath в UpgradeProjectsLogic. Ошибка: " & Err.description
        Err.Clear
    End If
    wb.SaveCopyAs currentPath
    If Err.Number <> 0 Then
        LogStep "Предупреждение: не удалось сохранить копию currentPath в UpgradeProjectsLogic. Ошибка: " & Err.description
        Err.Clear
    End If
    On Error GoTo 0
    
    wb.Close SaveChanges:=False: Set wb = Nothing
    LogStep "UpgradeProjectsLogic: завершено"
End Sub


