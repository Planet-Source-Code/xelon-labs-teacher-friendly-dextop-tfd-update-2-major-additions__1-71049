Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Animation As Boolean
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Declare Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, _
        ByVal bScan As Byte, _
        ByVal dwFlags As Long, _
        ByVal dwExtraInfo As Long)


Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Boolean
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'Private Const HWND_TOPMOST = -1
'Private Const HWND_NOTOPMOST = -2
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type NewForm2
        Height As Long
        Width As Long
        Left As Long
        Top As Long
End Type

Public Sub MakeNormal(hwnd As Long)
On Error Resume Next
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub SetFGWindow(ByVal hwnd As Long, Show As Boolean)
On Error Resume Next
If Show Then
If IsIconic(hwnd) Then
ShowWindow hwnd, SW_RESTORE
Else
BringWindowToTop hwnd
End If
Else
ShowWindow hwnd, SW_MINIMIZE
End If
End Sub

Public Sub SetDesktop(Whwnd As Long, WindowHwnd As Form)
On Error Resume Next
    SetWindowPos Whwnd, HWND_BOTTOM, WindowHwnd.Top / Screen.TwipsPerPixelX, WindowHwnd.Left / Screen.TwipsPerPixelY, WindowHwnd.Width / Screen.TwipsPerPixelX, WindowHwnd.Height / Screen.TwipsPerPixelY, 0
End Sub

Public Function ActivateWindow(hwnd As Long)
On Error Resume Next
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOZORDER
ShowWindow hwnd, 9
End Function

Sub FormOnBottom(frm As Long)
On Error Resume Next
Dim DeskH As Long
DeskH = GethWndByWinTitle("Program Manager")
Call SetParent(frm, DeskH)
End Sub

Sub FormNormal(frm As Long)
On Error Resume Next
Dim DeskH As Long
DeskH = GethWndByWinTitle("Form1")
Call SetParent(frm, DeskH)
End Sub

Sub FormOnTop(frm As Long)
On Error Resume Next
Call SetWindowPos(frm, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub

Public Function GethWndByWinTitle(winTitle As String) As Long
On Error Resume Next
    Dim retval As Long
    GethWndByWinTitle = FindWindow(vbNullString, winTitle)
End Function
Public Function maximize(hwnd As Long)
On Error Resume Next
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOZORDER
ShowWindow hwnd, 3
End Function
Public Function minimize(hwnd As Long)
On Error Resume Next
ShowWindow hwnd, 2
End Function

Public Function Normalize(hwnd As Long)
On Error Resume Next
    SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOREDRAW Or SWP_NOSIZE Or SWP_NOREPOSITION Or SWP_NOZORDER
ShowWindow hwnd, 1
End Function

Public Function AppHide(hwnd As Long)
On Error Resume Next
ShowWindow hwnd, 0
End Function

