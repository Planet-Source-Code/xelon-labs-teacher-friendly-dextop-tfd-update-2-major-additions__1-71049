VERSION 5.00
Begin VB.Form frmCam 
   Caption         =   "frmcam"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   LinkTopic       =   "Form7"
   ScaleHeight     =   8085
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstDevices 
      Height          =   450
      Left            =   0
      TabIndex        =   1
      Top             =   7320
      Width           =   3015
   End
   Begin VB.PictureBox piccapture 
      Height          =   7200
      Left            =   0
      ScaleHeight     =   7140
      ScaleWidth      =   9540
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "frmCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_CAP As Integer = &H400
Private Const WM_CAP_DRIVER_CONNECT As Long = WM_CAP + 10
Private Const WM_CAP_DRIVER_DISCONNECT As Long = WM_CAP + 11
Private Const WM_CAP_EDIT_COPY As Long = WM_CAP + 30

Private Const WM_CAP_SET_PREVIEW As Long = WM_CAP + 50
Private Const WM_CAP_SET_PREVIEWRATE As Long = WM_CAP + 52
Private Const WM_CAP_SET_SCALE As Long = WM_CAP + 53
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Integer = 1
Private Const SWP_NOZORDER As Integer = &H4
Private Const HWND_BOTTOM As Integer = 1

Dim iDevice As Long
Dim hHwnd As Long
Private Declare Function SendMessageS Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Private Declare Function capCreateCaptureWindowA Lib "avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Integer, ByVal hWndParent As Long, ByVal nID As Long) As Long
Private Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Long, ByVal lpszName As String, ByVal cbName As Long, ByVal lpszVer As String, ByVal cbVer As Long) As Boolean
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Sub snapshot(pth As String)
iDevice = 0
    hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640, 480, piccapture.hWnd, 0)

    If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
        SendMessage hHwnd, WM_CAP_SET_SCALE, True, 0
        SendMessage hHwnd, WM_CAP_SET_PREVIEWRATE, 3, 0
        SendMessage hHwnd, WM_CAP_SET_PREVIEW, True, 0
    
   SendMessage hHwnd, WM_CAP_EDIT_COPY, 0, 0
   SavePicture Clipboard.GetData, pth

    SendMessage hHwnd, WM_CAP_DRIVER_DISCONNECT, iDevice, 0
    DestroyWindow hHwnd
    Unload Me
    Else
        DestroyWindow hHwnd
    End If
 End Sub

