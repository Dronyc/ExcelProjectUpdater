Attribute VB_Name = "ModuleLogging"
'===============================================================
' МОДУЛЬ ЛОГИРОВАНИЯ
'===============================================================
Option Explicit

Private mLogFilePath As String
Private mLogEnabled As Boolean
Private mStartTime As Double
Private mLogFolder As String

'===============================================================
' Инициализация логирования
'===============================================================
Public Sub InitLogging(Optional ByVal logFileName As String = "")
    Dim rootFolder As String
    
    If Len(ThisWorkbook.path) = 0 Then
        MsgBox "Сначала сохраните файл в рабочей папке.", vbExclamation
        Exit Sub
    End If
    
    rootFolder = ThisWorkbook.path & "\"
    mLogFolder = rootFolder & "Logs\"
    
    ' Создаем папку для логов если не существует
    If Dir(mLogFolder, vbDirectory) = "" Then
        MkDir mLogFolder
    End If
    
    ' Формируем имя файла
    If Len(logFileName) = 0 Then
        logFileName = Format(Now, "yyyy-mm-dd_hh-mm-ss") & "_Log.txt"
    End If
    
    mLogFilePath = mLogFolder & logFileName
    mLogEnabled = True
    mStartTime = Timer
    
    ' Записываем заголовок лога
    WriteLog "=== НАЧАЛО СЕАНСА ==="
    WriteLog "Время начала: " & Format(Now, "dd.mm.yyyy hh:nn:ss")
    WriteLog "Путь к файлу: " & ThisWorkbook.FullName
    WriteLog "Версия Excel: " & Application.version
    WriteLog "====================="
End Sub

'===============================================================
' Запись сообщения в лог
'===============================================================
Public Sub WriteLog(ByVal message As String)
    If Not mLogEnabled Then Exit Sub
    
    Dim fileNum As Integer
    Dim timestamp As String
    Dim elapsed As Double
    
    fileNum = FreeFile
    timestamp = Format(Now, "hh:nn:ss")
    elapsed = Timer - mStartTime
    
    On Error Resume Next
    Open mLogFilePath For Append As #fileNum
    If Err.Number = 0 Then
        Print #fileNum, "[" & timestamp & " +" & Format(elapsed, "0.000") & "s] " & message
        Close #fileNum
    Else
        Close #fileNum
        Err.Clear
    End If
    On Error GoTo 0
End Sub

'===============================================================
' Запись шага в лог
'===============================================================
Public Sub LogStep(ByVal stepName As String, Optional ByVal details As String = "")
    Dim msg As String
    msg = "ШАГ: " & stepName
    If Len(details) > 0 Then
        msg = msg & " | " & details
    End If
    WriteLog msg
End Sub

'===============================================================
' Запись ошибки в лог
'===============================================================
Public Sub LogError(ByVal stepName As String, ByVal errNumber As Long, ByVal errDescription As String)
    WriteLog "!!! ОШИБКА !!!"
    WriteLog "  Шаг: " & stepName
    WriteLog "  Код: " & errNumber
    WriteLog "  Описание: " & errDescription
    WriteLog "!!!!!!!!!!!!!!"
End Sub

'===============================================================
' Завершение логирования
'===============================================================
Public Sub EndLogging()
    If Not mLogEnabled Then Exit Sub
    
    WriteLog "=== ЗАВЕРШЕНИЕ СЕАНСА ==="
    WriteLog "Время окончания: " & Format(Now, "dd.mm.yyyy hh:nn:ss")
    WriteLog "Общее время: " & Format(Timer - mStartTime, "0.000") & " сек"
    WriteLog "Путь к логу: " & mLogFilePath
    WriteLog "========================="
    
    mLogEnabled = False
End Sub

'===============================================================
' Показ критической ошибки с записью в лог
'===============================================================
Public Sub CriticalError(ByVal stepName As String, ByVal errNumber As Long, ByVal errDescription As String)
    LogError stepName, errNumber, errDescription
    EndLogging
    
    Dim msg As String
    msg = "Ошибка при создании структуры аналитики." & vbCrLf & _
          "Шаг: " & stepName & vbCrLf & _
          "Код ошибки: " & errNumber & vbCrLf & _
          "Описание: " & errDescription & vbCrLf & vbCrLf & _
          "Лог сохранен в: " & vbCrLf & mLogFilePath
    
    MsgBox msg, vbCritical, "Ошибка"
End Sub

'===============================================================
' Получение пути к последнему логу
'===============================================================
Public Function GetLastLogPath() As String
    GetLastLogPath = mLogFilePath
End Function

'===============================================================
' Получение папки логов
'===============================================================
Public Function GetLogFolder() As String
    GetLogFolder = mLogFolder
End Function

'===============================================================
' Открыть папку с логами
'===============================================================
Public Sub OpenLogFolder()
    If Len(mLogFolder) = 0 Then
        mLogFolder = ThisWorkbook.path & "\Logs\"
    End If
    
    If Dir(mLogFolder, vbDirectory) = "" Then
        MsgBox "Папка с логами не найдена: " & mLogFolder, vbInformation
        Exit Sub
    End If
    
    Shell "explorer.exe """ & mLogFolder & """", vbNormalFocus
End Sub

'===============================================================
' Показать последний лог
'===============================================================
Public Sub ShowLastLog()
    If Len(mLogFilePath) = 0 Or Dir(mLogFilePath) = "" Then
        ' Ищем последний лог в папке
        Dim logFile As String
        logFile = Dir(mLogFolder & "Log_*.txt")
        If Len(logFile) = 0 Then
            MsgBox "Логи не найдены в папке: " & mLogFolder, vbInformation
            Exit Sub
        End If
        ' Берем последний по имени (сортировка по дате в имени)
        Dim lastFile As String
        lastFile = ""
        Do While Len(logFile) > 0
            lastFile = logFile
            logFile = Dir()
        Loop
        mLogFilePath = mLogFolder & lastFile
    End If
    
    If Len(mLogFilePath) > 0 And Dir(mLogFilePath) <> "" Then
        Shell "notepad.exe """ & mLogFilePath & """", vbNormalFocus
    Else
        MsgBox "Файл лога не найден", vbInformation
    End If
End Sub

