Attribute VB_Name = "ModuleUI"
'===============================================================
' МОДУЛЬ ИНТЕРФЕЙСА (ModuleUI) v2
'===============================================================
Option Explicit

Private Const BTN_WIDTH As Long = 260
Private Const BTN_HEIGHT As Long = 40
Private Const BTN_LEFT As Long = 20
Private Const BTN_GAP As Long = 55
Private COLOR_BTN_MAIN As Long
Private COLOR_BG_HEADER As Long

Private Sub InitColors()
    COLOR_BTN_MAIN = RGB(91, 155, 213)
    COLOR_BG_HEADER = RGB(221, 235, 247)
End Sub

'===============================================================
' ГЛАВНАЯ ПРОЦЕДУРА: создание листа "Управление"
'===============================================================
Public Sub CreateControlButtons()
    InitColors
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("Управление")
    
    ' Очистка листа
    ws.Cells.Clear
    Dim i As Long
    For i = ws.Shapes.count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    ' Настройка геометрии
    SetupSheetGeometry ws
    
    ' Заголовок
    ws.Range("A1").value = "Запуск макросов"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    
    ' Создание кнопок и легенды
    CreateMainButtons ws
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub SetupSheetGeometry(ByVal ws As Worksheet)
    ws.Columns("A:C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 28
    ws.Columns("E").ColumnWidth = 70
End Sub

'===============================================================
' СОЗДАНИЕ КНОПОК И ЛЕГЕНДЫ
'===============================================================
Private Sub CreateMainButtons(ByVal ws As Worksheet)
    '--- Заголовки таблицы пояснений (строка 2) ---
    ws.Range("D2").value = "Кнопка"
    ws.Range("E2").value = "Что делает и когда нажимать"
    ws.Range("D2").Font.Bold = True
    ws.Range("E2").Font.Bold = True
    
    '--- Пояснения (строки 3-6) ---
    ws.Range("D3").value = "Создать всё с нуля"
    ws.Range("E3").value = "Создаёт файлы шаблонов И структуру аналитики. ВНИМАНИЕ: все введенные данные стираются! Нажимать ТОЛЬКО по указанию администратора."
    
    ws.Range("D4").value = "Обновить проекты"
    ws.Range("E4").value = "Обновляет реестр по файлу ""Выгрузка проектов_export.xlsx"": добавляет новые и обновляет существующие проекты, пишет ошибки на лист ""Ошибки"", убирает выгрузку в архив. Нажимать после каждой новой выгрузки."
    
    ws.Range("D5").value = "Пересчитать аналитику"
    ws.Range("E5").value = "Заполняет листы ""Эталон_Данные"" и ""Аналитика"" данными из эталонного файла (папка Reference). Нажимать после смены эталонного файла."
    
    ws.Range("D6").value = "Обновить логику и оформление"
    ws.Range("E6").value = "Применяет изменения формул, оформления, структуры и цветов в файле ""Проекты РЦ АСКОН_Волга в Архив.xlsx"" без потери данных. Нажимать после смены цветов легенды или по указанию администратора."
    
    '--- Форматирование пояснений ---
    ws.Range("D3:D6").Font.Bold = True
    ws.Range("E3:E6").WrapText = True
    ws.Rows("3:6").rowHeight = 60
    
    '--- Границы таблицы пояснений ---
    ApplyBordersToRange ws.Range("D2:E6")
    
    '--- Создание кнопок ---
    CreateButton ws, "btnCreateAll", "Создать всё с нуля", _
        "CreateAllFromScratch", ButtonTopForRow(ws, 3)
    
    CreateButton ws, "btnUpdateProjects", "Обновить проекты", _
        "UpdateProjectsMain", ButtonTopForRow(ws, 4)
    
    CreateButton ws, "btnRefreshAnalytics", "Пересчитать аналитику", _
        "RefreshAnalyticsMain", ButtonTopForRow(ws, 5)
    
    CreateButton ws, "btnUpgradeAll", "Обновить логику и оформление", _
        "UpgradeAllLogicAndFormatting", ButtonTopForRow(ws, 6)
End Sub

Private Function ButtonTopForRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
    Dim rowTop As Double
    Dim rowHeight As Double
    rowTop = ws.Rows(rowNum).Top
    rowHeight = ws.Rows(rowNum).Height
    ButtonTopForRow = CLng(rowTop + (rowHeight - BTN_HEIGHT) / 2)
End Function

Private Sub CreateButton( _
    ByVal ws As Worksheet, _
    ByVal btnName As String, _
    ByVal btnCaption As String, _
    ByVal macroName As String, _
    ByVal topPos As Long)
    Dim btn As Shape
    Set btn = ws.Shapes.AddFormControl( _
        xlButtonControl, BTN_LEFT, topPos, BTN_WIDTH, BTN_HEIGHT)
    btn.Placement = xlFreeFloating
    btn.name = btnName
    btn.OnAction = ThisWorkbook.name & "!" & macroName
    With btn.TextFrame
        .Characters.Text = btnCaption
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    btn.Fill.ForeColor.RGB = COLOR_BTN_MAIN
    btn.Line.ForeColor.RGB = RGB(0, 0, 0)
    btn.Line.Weight = 1
End Sub

Private Sub ApplyBordersToRange(ByVal rng As Range)
    Dim e As Long
    For e = 7 To 12
        With rng.Borders(e)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .color = RGB(0, 0, 0)
        End With
    Next e
End Sub

