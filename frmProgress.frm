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

' API declaration for setting window on top
Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hWndInsertAfter As LongPtr, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST As LongPtr = -1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

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
    
    ' Make the form topmost so it stays above all other windows
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
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
