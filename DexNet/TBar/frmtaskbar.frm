VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmtaskbar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   600
   End
   Begin VB.ListBox lstapps 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   3120
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox lstnames 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   3840
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   6240
      Picture         =   "frmtaskbar.frx":0000
      ScaleHeight     =   525
      ScaleWidth      =   1095
      TabIndex        =   5
      Top             =   0
      Width           =   1095
      Begin VB.Label tym 
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "hh:mm AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00F8E6E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   1320
      Picture         =   "frmtaskbar.frx":3E10
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   480
      Top             =   600
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   960
      Top             =   600
   End
   Begin VB.ListBox lsttrans 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox ptmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   480
      ScaleHeight     =   825
      ScaleWidth      =   825
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer blnk 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   600
   End
   Begin PrjTskbr.PictureButton PictureButton1 
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1500
      _extentx        =   2646
      _extenty        =   820
      picture         =   "frmtaskbar.frx":5290
      picturehover    =   "frmtaskbar.frx":773A
      picturedown     =   "frmtaskbar.frx":9BE2
   End
   Begin MSComctlLib.Slider translider 
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      ToolTipText     =   "Slide to change transparency effects"
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Max             =   255
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   295
      Left            =   4560
      TabIndex        =   9
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393217
      BackColor       =   16578547
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"frmtaskbar.frx":C08C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   0
      Picture         =   "frmtaskbar.frx":C113
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   360
      Width           =   15
   End
   Begin VB.Image imgico 
      Height          =   330
      Left            =   1680
      Picture         =   "frmtaskbar.frx":D593
      Top             =   1440
      Width           =   435
   End
End
Attribute VB_Name = "frmtaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public X2 As Integer, Y2 As Integer
Public imenu As Integer
Dim NewX
Dim NewY
Dim blk As Integer
Dim hlt As Boolean
Dim thwnd As Long


Private Sub blnk_Timer()
On Error Resume Next
If blk = 0 Then
blk = 1
        Picture4(1).Picture = Picture4(0).Picture
        Picture4(1).BackColor = vbBlack
        Call DrawIcon(Picture4(1).hDC, lstapps.list(0), 1, 1)
ElseIf blk = 1 Then
blk = 2
        Picture4(1).Picture = imgico.Picture
        Picture4(1).BackColor = vbHighlight
        Call DrawIcon(Picture4(1).hDC, lstapps.list(0), 1, 1)
ElseIf blk = 2 Then
        Picture4(1).Picture = Picture4(0).Picture
        Picture4(1).BackColor = vbBlack
        Call DrawIcon(Picture4(1).hDC, lstapps.list(0), 1, 1)
        blnk = False
        blk = 0
End If
End Sub

Private Sub Form_GotFocus()
On Error Resume Next
Clear
InitButtons
frmmenu.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 112 And Shift = 1 Then
PictureButton1_Click
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
FormOnTop Me.hwnd
fEnumWindows Me.lstapps
DoEvents
InitButtons
DoEvents
DockForm Me, DockTop
End Sub

Private Sub Form_Resize()
On Error Resume Next
Top = 0
Left = 0
Width = Screen.Width
Image1.Width = Width
Picture1.Left = Width - Picture1.Width
text1.Left = Picture1.Left - translider.Width - 30
translider.Left = text1.Left - 75 - translider.Width
Height = 460
End Sub

Private Sub Form_Terminate()
On Error Resume Next
UnDockForm Me
restore
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
restore
UnDockForm Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu frmpop.hideo
ElseIf Button = 1 Then
If Timer3 = True Then
Timer3 = False
Else
Timer3 = True
End If
End If
frmmenu.Visible = False

End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Picture4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then
Dim rct As RECT
GetWindowRect Picture4(Index).Tag, rct
If rct.Top < Me.Height / 15 And rct.Top > -frmtaskbar.Height Then
SetWindowPos Picture4(Index).Tag, 0, rct.Left, Me.Height / 15, rct.Right - rct.Left, rct.Bottom - rct.Top, 0
End If
ActivateWindow Picture4(Index).Tag
ElseIf Button = 2 Then
imenu = Index
PopupMenu frmpop.trans, , Picture4(Index).Left, Me.Height - 40
End If
End Sub

Private Sub PictureButton1_Click()
On Error Resume Next
If frmmenu.Visible = False Then
frmmenu.Visible = True
frmmenu.Top = frmtaskbar.Top + frmtaskbar.Height
fade frmmenu.hwnd, 20
DoEvents
Else
MakeTransparent frmmenu.hwnd, 0
frmmenu.Hide
End If
End Sub

Private Sub PictureButton1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Frmtip2.Screentip "Start Button: Click to Begin", 1125, Top + Height
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
PopupMenu frmpop.ser, , text1.Left, text1.Top + text1.Height
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim x As Integer
fEnumWindows Me.lstapps
DoEvents
For x = 0 To Picture4.UBound
    If Picture4.UBound <> lstapps.ListCount Then
        Clear
        InitButtons
        If lstapps.ListCount > 0 Or Picture4(x).ToolTipText <> lstnames.list(x - 1) Then
        Picture4(1).Picture = imgico.Picture
            Picture4(1).BackColor = vbHighlight
            Call DrawIcon(Picture4(1).hDC, lstapps.list(0), 1, 1)
            blnk = True
            blk = 0
        End If
        Exit For
    End If
Next
For x = 1 To Picture4.UBound
Picture4(x).BorderStyle = 0
Next
End Sub

Function InitButtons()
On Error Resume Next
lsttrans.Clear
For X2 = 0 To lstapps.ListCount - 1
        load Picture4(Picture4.UBound + 1)
        With Picture4(Picture4.UBound)
                .Left = Picture4(Picture4.UBound - 1).Left + 320
                .AutoRedraw = True
                .Visible = True
                .ZOrder 0
                .ToolTipText = lstnames.list(X2)
                .Tag = lstapps.list(X2)
                 Call DrawIcon(Picture4(Picture4.UBound).hDC, lstapps.list(X2), 1, 1)
                 lsttrans.AddItem "255"
        End With
'DoEvents
Next
Timer1.Enabled = True
Frmulti.Form_Load
End Function


Public Sub DrawIcon(hDC As Long, hwnd As Long, x As Integer, y As Integer)
On Error Resume Next
ico = GetIcon(hwnd)
DrawIconEx hDC, x, y, ico, 16, 16, 0, 0, &H3
End Sub

Public Function GetIcon(hwnd As Long) As Long
On Error Resume Next
Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICONSM)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICON)
If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

Function CheckCaption(text) As String
On Error Resume Next
Dim TextLen
TextLen = 30
If Len(text) > TextLen Then
text = Left(text, TextLen) & "..."
End If
CheckCaption = text
End Function

Public Sub Clear()
On Error Resume Next
If Picture4.UBound >= 1 Then
For x = 1 To Picture4.UBound
Unload Picture4(x)
Next
End If
Me.Width = Screen.Width
End Sub

Private Function CheckMouseOver(idx As Integer) As Boolean
On Error Resume Next
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.x, pt.y) = Picture4(idx).hwnd)
End Function

Private Sub Timer2_Timer()
On Error Resume Next
    Dim pt As POINTAPI
    Dim x
    Dim r, rw As Integer, rh As Integer
    Dim re As RECT
    r = -1
    GetCursorPos pt
For x = 1 To Picture4.UBound
If WindowFromPoint(pt.x, pt.y) = Picture4(x).hwnd Then
r = x
Exit For
End If
Next
If r <> -1 Then
If Frmtip2.Visible = False Or Frmtip2.lbl <> Picture4(r).ToolTipText Then
    GetWindowRect Picture4(r).Tag, re
    rw = re.Right - re.Left
    rh = re.Bottom - re.Top
Picture4(r).BorderStyle = 1
Frmtip2.Screentip Picture4(r).ToolTipText, 1125, Top + Height
Set dsp.Image1.Picture = Nothing
Set dsp.Image1 = CaptureForm(Picture4(r).Tag, rw, rh)
dsp.Show
dsp.Left = 1125
dsp.Top = Top + Height + Frmtip2.Height
dsp.Image1.Height = (dsp.Image1.Picture.Height * ((1851 / Screen.Height) / 15)) / 1.5
dsp.Image1.Width = (dsp.Image1.Picture.Width * ((1851 / Screen.Width) / 15)) / 1.5
dsp.Height = dsp.Image1.Height * 15
dsp.Width = dsp.Image1.Width * 15
End If
ElseIf r = -1 Then
Frmtip2.Visible = False
dsp.Visible = False
Set dsp.Image1.Picture = Nothing
For x = 1 To Picture4.UBound
Picture4(x).BorderStyle = 0
Next
End If
tym = Time
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Dim y As Integer
    Dim pt As POINTAPI
    GetCursorPos pt
If pt.y >= (Me.Height) / 15 Then
If Visible = True Then
Visible = False
Timer2 = False
Timer3 = True
Frmtip2.Hide
dsp.Hide
Clear
InitButtons
End If
Else
If pt.y < 7 Then
If Visible = False Then
Visible = True
Timer2 = True
ZOrder 0
End If
End If
End If
frmmenu.Top = frmtaskbar.Top + frmtaskbar.Height
End Sub

Private Sub translider_Change()
On Error Resume Next
Dim mhwnd As Long
mhwnd = GetParent(thwnd)
MakeTransparent thwnd, translider.Value
MakeTransparent mhwnd, translider.Value
End Sub

Private Sub translider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 2 Then
translider.Visible = False
End If
End Sub

Private Sub translider_Scroll()
On Error Resume Next
MakeTransparent thwnd, translider.Value
End Sub

Private Sub tym_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then
Screen.MousePointer = 2
ElseIf Button = 2 Then
PopupMenu frmpop.hideo
End If
End Sub

Private Sub tym_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Frmtip2.Screentip "Drag the clock to any window to set transparency", Me.Width - Frmtip2.Width, Top + Height
End Sub

Public Sub restore()
On Error Resume Next
AppHide Me.hwnd
For x = 1 To frmpop.hides.UBound
On Error Resume Next
ActivateWindow frmpop.hides(x).Tag
MakeOpaque frmpop.hides(x).Tag
Unload frmpop.hides(x)
Next
Unload Me
End Sub

Private Sub tym_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim pt As POINTAPI
If Button = 1 Then
Screen.MousePointer = 0
GetCursorPos pt
thwnd = WindowFromPoint(pt.x, pt.y)
translider.Visible = True
End If
End Sub

