VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lblStatus As MSForms.label
Private lblBack As MSForms.label
Private lblFill As MSForms.label
Private lblPct As MSForms.label

' WinAPI declaration for setting window always on top
#If VBA7 Then
    Private Declare PtrSafe Sub SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Private Const HWND_TOPMOST As LongPtr = -1
    Private Const SWP_NOMOVE As Long = &H2
    Private Const SWP_NOSIZE As Long = &H1
#Else
    Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    Private Const HWND_TOPMOST As Long = -1
    Private Const SWP_NOMOVE As Long = &H2
    Private Const SWP_NOSIZE As Long = &H1
#End If

Private Sub UserForm_Initialize()
    Me.Caption = "Обновление проектов"
    Me.Width = 420
    Me.Height = 130
    Me.StartUpPosition = 1
    
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    lblStatus.Left = 12: lblStatus.Top = 12
    lblStatus.Width = 380: lblStatus.Height = 30
    lblStatus.WordWrap = True
    lblStatus.Caption = "Подготовка..."
    
    Set lblBack = Me.Controls.Add("Forms.Label.1", "lblBack")
    lblBack.Left = 12: lblBack.Top = 50
    lblBack.Width = 380: lblBack.Height = 22
    lblBack.BackColor = RGB(220, 220, 220)
    
    Set lblFill = Me.Controls.Add("Forms.Label.1", "lblFill")
    lblFill.Left = 14: lblFill.Top = 52
    lblFill.Width = 0: lblFill.Height = 18
    lblFill.BackColor = RGB(47, 117, 181)
    
    Set lblPct = Me.Controls.Add("Forms.Label.1", "lblPct")
    lblPct.Left = 12: lblPct.Top = 78
    lblPct.Width = 380: lblPct.Height = 16
    lblPct.Caption = "0%"
    
    ' Устанавливаем форму поверх всех окон (TopMost)
    On Error Resume Next
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    On Error GoTo 0
End Sub

Private Sub UserForm_Terminate()
    ' Снимаем флаг TopMost при закрытии формы
    On Error Resume Next
    SetWindowPos Me.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    On Error GoTo 0
End Sub

Public Sub SetProgress(ByVal pct As Double, ByVal comment As String)
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    
    lblStatus.Caption = comment
    lblFill.Width = 376 * pct
    lblPct.Caption = Format(pct, "0%")
    
    ' DoEvents вызываем только при значительном изменении
    DoEvents
End Sub
