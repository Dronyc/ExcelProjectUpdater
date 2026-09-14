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
    LogStep "CreateProjectFiles: начало"
    
    Dim finalFilePath As String, currentFilePath As String
    finalFilePath = rootFolder & "Проекты РЦ АСКОН_Волга в Архив.xlsx"
    currentFilePath = rootFolder & "Выгрузка проектов_current.xlsx"
    
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
    
    ' Создаём листы в нужном порядке
    Dim wsProjects As Worksheet, wsStat As Worksheet, wsAnalytics As Worksheet
    Dim wsLegend As Worksheet, wsManualRef As Worksheet, wsAllowedRef As Worksheet
    Dim wsProductRef As Worksheet, wsErrors As Worksheet, wsSource As Worksheet
    
    Set wsProjects = wb.Sheets(1): wsProjects.name = "Проекты"
    Set wsStat = wb.Sheets.Add(After:=wsProjects): wsStat.name = "Статистика"
    Set wsAnalytics = wb.Sheets.Add(After:=wsStat): wsAnalytics.name = "Аналитика"
    Set wsLegend = wb.Sheets.Add(After:=wsAnalytics): wsLegend.name = "Легенда"
    Set wsManualRef = wb.Sheets.Add(After:=wsLegend): wsManualRef.name = "СправочникСтатусов"
    Set wsAllowedRef = wb.Sheets.Add(After:=wsManualRef): wsAllowedRef.name = "СправочникСостояний"
    Set wsProductRef = wb.Sheets.Add(After:=wsAllowedRef): wsProductRef.name = "СправочникПродуктов"
    Set wsErrors = wb.Sheets.Add(After:=wsProductRef): wsErrors.name = "Ошибки"
    Set wsSource = wb.Sheets.Add(After:=wsErrors): wsSource.name = "Выгрузка"
    
    ' Заполняем листы
    CreateLegendSheet wsLegend, wb, finalFilePath, currentFilePath
    CreateManualReferenceSheet wsManualRef, wb
    CreateAllowedReferenceSheet wsAllowedRef, wb
    CreateProjectsSheet wsProjects, wsLegend
    CreateErrorsSheet wsErrors
    CreateSourceSheet wsSource
    CreateStatisticsSheet wb, wsStat
    CreateProductReferenceSheet wb
    
    ' Применяем стили ко всем таблицам
    Dim wsAny As Worksheet, loAny As ListObject
    For Each wsAny In wb.Worksheets
        For Each loAny In wsAny.ListObjects
            loAny.tableStyle = "TableStyleLight13"
        Next loAny
    Next wsAny
    
    wsProjects.ListObjects("тблПроекты").ListColumns("Дата создания").Range.NumberFormatLocal = "дд.мм.гггг чч:мм:сс"
    
    If Not DeleteFileIfExists(finalFilePath) Then
        Err.Raise vbObjectError + 3, , "Не удалось удалить: " & finalFilePath
    End If
    If Not DeleteFileIfExists(currentFilePath) Then
        Err.Raise vbObjectError + 4, , "Не удалось удалить: " & currentFilePath
    End If
    
    wb.SaveAs fileName:=finalFilePath, FileFormat:=xlOpenXMLWorkbook
    wb.SaveCopyAs currentFilePath
    wb.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    LogStep "CreateProjectFiles: завершено"
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
    
    For r = 1 To rowCount
        If r >= nextUpdate Or r = rowCount Then
            ProgressSmooth 25 + (CDbl(r) / rowCount) * 20, "Анализ строк (" & r & " из " & rowCount & ")..."
            nextUpdate = r + stepSize
        End If
        
        srcValid(r) = False: srcProcessed(r) = False
        If IsEmptySourceRow(srcValues, r, colIdx) Then GoTo NextSourceRow
        
        codeText = GetCodeText(srcValues(r, colIdx(3)))
        nameText = GetCleanText(srcValues(r, colIdx(5)))
        parsedDate = ParseRuDateTime(srcValues(r, colIdx(2)))
        
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
    
    For r = 1 To rowCount
        If r >= nextUpdate Or r = rowCount Then
            ProgressSmooth 70 + (CDbl(r) / rowCount) * 10, "Обновление (" & r & " из " & rowCount & ")..."
            nextUpdate = r + stepSize
        End If
        
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
    For r = 1 To rowCount
        If r >= nextUpdate Or r = rowCount Then
            ProgressSmooth 85 + (CDbl(r) / rowCount) * 7, "Добавление (" & r & " из " & rowCount & ")..."
            nextUpdate = r + stepSize
        End If
        
        If srcValid(r) And Not srcProcessed(r) Then
            Set lr = loProjects.ListRows.Add
            WriteSourceRowToListRow loProjects, lr, srcValues, r, colIdx, srcDate(r), srcCode(r)
            srcProcessed(r) = True: addedCount = addedCount + 1
        End If
    Next r
    
    gStep = "Обновление групп ПГС/PLM"
    ProgressSet 93, "Обновление групп ПГС/PLM..."
    BuildProductMapping loProjects
    FillGroupColumns loProjects
    
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
    wbTarget.SaveCopyAs currentPath
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
    If Not wbSource Is Nothing Then On Error Resume Next: wbSource.Close SaveChanges:=False
    If Not wbTarget Is Nothing Then On Error Resume Next: wbTarget.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True: Application.EnableEvents = True: Application.DisplayAlerts = True
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
    
    ' Группы
    BuildProductMapping loProjects
    FillGroupColumns loProjects
    
    ' Статистика
    Dim wsStat As Worksheet
    Set wsStat = Nothing
    On Error Resume Next
    Set wsStat = wb.Worksheets("Статистика")
    On Error GoTo 0
    If wsStat Is Nothing Then
        Set wsStat = wb.Worksheets.Add(After:=wsProjects)
        wsStat.name = "Статистика"
    End If
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
    wb.SaveCopyAs currentPath
    On Error GoTo 0
    
    wb.Close SaveChanges:=False: Set wb = Nothing
    LogStep "UpgradeProjectsLogic: завершено"
End Sub


