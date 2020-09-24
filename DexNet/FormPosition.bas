Attribute VB_Name = "FormPosition"
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Sub FormOnBottom(frm As Form)
On Error Resume Next
Dim DeskH As Long
DeskH = GethWndByWinTitle("Program Manager")
Call SetParent(frm.hwnd, DeskH)
End Sub


Sub FormNormal(frm As Form)
On Error Resume Next
Dim DeskH As Long
DeskH = GethWndByWinTitle("Form1")
Call SetParent(frm.hwnd, DeskH)
End Sub

Sub hwndNormal(hwnd As Long)
On Error Resume Next
Dim DeskH As Long
DeskH = GethWndByWinTitle("Form1")
Call SetParent(hwnd, DeskH)
End Sub

Sub FormOnTop(frm As Form)
On Error Resume Next
Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Public Function GethWndByWinTitle(winTitle As String) As Long
On Error Resume Next
    Dim retval As Long
    GethWndByWinTitle = FindWindow(vbNullString, winTitle)
End Function

