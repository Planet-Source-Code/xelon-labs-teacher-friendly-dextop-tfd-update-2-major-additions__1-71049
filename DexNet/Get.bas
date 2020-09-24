Attribute VB_Name = "Get"
Option Explicit
Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&

Private Const Red As Integer = 3
Private Const Green As Integer = 2
Private Const Blue As Integer = 1

Private Type BITMAPINFOHEADER '40 bytes
      biSize As Long
      biWidth As Long
      biHeight As Long
      biPlanes As Integer
      biBitCount As Integer
      biCompression As Long
      biSizeImage As Long
      biXPelsPerMeter As Long
      biYPelsPerMeter As Long
      biClrUsed As Long
      biClrImportant As Long
End Type

Private Type BITMAPINFO
      bmiHeader As BITMAPINFOHEADER
     'bmiColors As RGBQUAD
End Type

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = -20
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_STYLE = (-16)
Public Const WS_VISIBLE = &H10000000
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const LWA_ALPHA = &H2
'''''''''''

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Type POINTAPI
    x As Long
    Y As Long
End Type
    

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


