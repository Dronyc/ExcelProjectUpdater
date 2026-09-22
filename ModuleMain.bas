Attribute VB_Name = "ModuleMain"
'===============================================================
' ГЛАВНЫЙ МОДУЛЬ УПРАВЛЕНИЯ
'===============================================================
Option Explicit

'===============================================================
' МАКРОС 1: Создать всё с нуля
'===============================================================
Public Sub CreateAllFromScratch()
    InitLogging "CreateAll_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
    LogStep "=== CreateAllFromScratch: НАЧАЛО ==="
    
    Dim rootFolder As String
    rootFolder = ThisWorkbook.path & "\"
    
    ProgressShow
    ProgressSet 0, "Создание файлов решения..."
    
    ' 1. Создаём файлы решения
    LogStep "Шаг 1: Создание файлов решения"
    ProgressSet 5, "Создание файлов решения..."
    CreateProjectFiles rootFolder
    
    ' 2. Создаём структуру аналитики
    LogStep "Шаг 2: Создание структуры аналитики"
    ProgressSet 50, "Создание структуры аналитики..."
    CreateAnalyticsStructure rootFolder
    
    ' 3. Упорядочивание листов во всех книгах
    LogStep "Шаг 3: Упорядочивание листов"
    ProgressSet 90, "Упорядочивание листов..."
    ReorderAllSheets
       
    ProgressSet 100, "Готово!"
    Application.Calculate
    
    ProgressHide
    
    EndLogging
    MsgBox "Структура проекта создана успешно!" & vbCrLf & vbCrLf & _
           "Созданы файлы и листы:" & vbCrLf & _
           "  [OK] Проекты РЦ АСКОН_Волга в Архив.xlsx" & vbCrLf & _
           "  [OK] Все листы аналитики (включая Эталон_Данные)" & vbCrLf & _
           "  [OK] СправочникПродуктов" & vbCrLf & vbCrLf & _
           "Дальнейшие действия:" & vbCrLf & _
           "1. Для обновления проектов используйте кнопку:" & vbCrLf & _
           "   > 'Обновить проекты'" & vbCrLf & vbCrLf & _
           "2. Для получения аналитики используйте кнопку:" & vbCrLf & _
           "   > 'Пересчитать аналитику'", vbInformation
End Sub

'===============================================================
' МАКРОС 2: Обновить проекты
'===============================================================
Public Sub UpdateProjectsMain()
    On Error GoTo ErrorHandler
    
    InitLogging "UpdateProjects_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
    LogStep "=== UpdateProjectsMain: НАЧАЛО ==="
    
    Dim rootFolder As String
    rootFolder = ThisWorkbook.path & "\"
    
    ProgressShow
    ProgressSet 0, "Обновление проектов..."
    
    ' Вызываем обновление из ModuleProjects
    UpdateProjectsFromExport rootFolder
    
    ProgressHide
    
    EndLogging
    Exit Sub
    
ErrorHandler:
    LogError "UpdateProjectsMain", Err.Number, Err.description
    EndLogging
    ProgressHide
    
    MsgBox "Ошибка при обновлении проектов." & vbCrLf & _
           "Шаг: " & gStep & vbCrLf & _
           "Код ошибки: " & Err.Number & vbCrLf & _
           "Описание: " & Err.description, vbCritical
    
    On Error Resume Next
    Set wbSource = Nothing
    Set wbTarget = Nothing
    On Error GoTo 0
End Sub

'===============================================================
' МАКРОС 3: Обновить логику и оформление
'===============================================================
Public Sub UpgradeAllLogicAndFormatting()
    InitLogging "UpgradeAll_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
    LogStep "=== UpgradeAllLogicAndFormatting: НАЧАЛО ==="
    
    Dim rootFolder As String
    rootFolder = ThisWorkbook.path & "\"
    
    ProgressShow
    ProgressSet 0, "Обновление логики и оформления..."
    
    ' 1. Обновляем логику проектов
    LogStep "Шаг 1: Обновление логики проектов"
    ProgressSet 10, "Обновление логики проектов..."
    UpgradeProjectsLogic rootFolder
    
    ' 2. Обновляем структуру аналитики
    LogStep "Шаг 2: Обновление структуры аналитики"
    ProgressSet 60, "Обновление структуры аналитики..."
    UpgradeAnalyticsStructure rootFolder
    
    ' 3. Упорядочивание листов во всех книгах
    LogStep "Шаг 3: Упорядочивание листов"
    ProgressSet 90, "Упорядочивание листов..."
    ReorderAllSheets
    
    ProgressSet 100, "Готово!"
    Application.Calculate
    
    ProgressHide
    
    EndLogging
    MsgBox "Логика и оформление обновлены!" & vbCrLf & _
           "Данные сохранены.", vbInformation
End Sub

'===============================================================
' Упорядочивание всех листов
' Эта процедура должна быть Public, так как вызывается из ModuleProjects.bas
'===============================================================
Public Sub ReorderAllSheets()
    Dim targetOrder As Variant
    targetOrder = Array("Проекты", "Статистика", "Аналитика", "Качество заполнения по людям", "Эталон_Данные", _
                       "Легенда", "СправочникСтатусов", "СправочникСостояний", _
                       "СправочникПродуктов", "Ошибки", "Выгрузка")
    Dim wb As Workbook
    Set wb = ThisWorkbook
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
            On Error Resume Next
            ws.Move After:=wb.Worksheets(i + 1)
            On Error GoTo 0
        End If
    Next i
End Sub

'===============================================================
' МАКРОС 4: Пересчитать аналитику
'===============================================================
Public Sub RefreshAnalyticsMain()
    InitLogging "RefreshAnalytics_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".txt"
    LogStep "=== RefreshAnalyticsMain: НАЧАЛО ==="
    
    Dim rootFolder As String
    rootFolder = ThisWorkbook.path & "\"
    
    ProgressShow
    ProgressSet 0, "Пересчёт аналитики..."
    
    ' Вызываем пересчёт из ModuleAnalytics
    RefreshAnalytics rootFolder
    
    ProgressHide
    EndLogging
End Sub


