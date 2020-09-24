Attribute VB_Name = "Module1"
Declare Function Shell_NotifyIconA Lib "shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Type NOTIFYICONDATA
     cbze As Long
     hWnd As Long
     exid As Long
     Flags As Long
     backmessage As Long
     icon As Long
     tip As String * 64
End Type

Global Const IcD_ADD = &H0&
Global Const IcD_MODIFY = &H1
Global Const IcD_DELETE = &H2
Global Const IcD_BACKMESSAGE = &H1
Global Const IcD_ICON = &H2
Global Const IcD_TIP = &H4
Global Const WM_MOUSEMOVE = &H200
Global IcD As NOTIFYICONDATA

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hDCSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Public shl As New shell
Public fso As New FileSystemObject

Public Function isTransparent(ByVal hWnd As Long) As Boolean
On Error Resume Next
Dim Msg As Long
Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
  isTransparent = True
Else
  isTransparent = False
End If
If err Then
  isTransparent = False
End If
End Function

Public Function MakeTransparent(ByVal hWnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hWnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hWnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If err Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hWnd As Long) As Long
Dim Msg As Long
On Error Resume Next
Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
Msg = Msg And Not WS_EX_LAYERED
SetWindowLong hWnd, GWL_EXSTYLE, Msg
SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If err Then
  MakeOpaque = 2
End If
End Function

Sub cpl(Dir As String)
On Error Resume Next
shl.ControlPanelItem Dir
End Sub

Sub snapshot(pth As String)
frmCam.snapshot pth
End Sub
