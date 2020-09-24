VERSION 5.00
Begin VB.UserControl title 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   ScaleHeight     =   3600
   ScaleWidth      =   6165
   Begin VB.Timer tmrfl 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   1560
   End
   Begin Project1.PictureButton macbutton1 
      Height          =   225
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   397
      Picture         =   "title.ctx":0000
      PictureHover    =   "title.ctx":07D4
      PictureDown     =   "title.ctx":0FA8
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   3615
   End
   Begin Project1.aicAlphaImage flare1 
      Height          =   315
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3916
      _ExtentY        =   3916
      Image           =   "title.ctx":177C
      Scaler          =   4
      Opacity         =   70
      Props           =   5
      ScaleCx         =   147
      ScaleCy         =   21
   End
   Begin Project1.aicAlphaImage flare 
      Height          =   315
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3916
      _ExtentY        =   3916
      Image           =   "title.ctx":32A1
      Scaler          =   4
      Opacity         =   70
      Props           =   5
      ScaleCx         =   147
      ScaleCy         =   21
   End
   Begin VB.Image icon 
      Height          =   240
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   6135
   End
   Begin VB.Image Title 
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   19200
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   0
      Picture         =   "title.ctx":4DC4
      Top             =   0
      Width           =   18930
   End
End
Attribute VB_Name = "title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Public fom As Object
Dim over As Boolean
Public cross As Boolean
Public obj As Boolean
Public unfold As Boolean
Sub unfolded()
On Error Resume Next
MacButton1_Click
End Sub

Private Sub MacButton1_Click()
On Error Resume Next
Un_blink
End Sub
Sub reload()
On Error Resume Next
Timer1 = True
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Timer1 = False
'If over = True Then
blink
'over = False
'Else
'Un_blink
'over = True
'End If
End Sub

Private Sub tmrfl_Timer()
On Error Resume Next
If tmrfl.Tag = "blink" Then
flare.Visible = True
If flare.Left < UserControl.Width Then
flare.Left = flare.Left + 230
Else
tmrfl = False
flare.Visible = False
End If
Else
flare1.Visible = True
If flare1.Left > -flare1.Width Then
flare1.Left = flare1.Left - 230
Else
tmrfl = False
flare1.Visible = False
If unfold = False Then
fom.Hide
Else
Unload fom
End If
End If
End If
End Sub


Sub Skin(path As String)
On Error Resume Next
On Error Resume Next
Image1 = LoadPicture(path & "\title.bmp")
MacButton1.loadpics path, "Close"
flare.LoadImage_FromFile path & "\Flare.png"
flare1.LoadImage_FromFile path & "\Flare1.png"
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
MacButton1.Left = UserControl.Width - MacButton1.Width
MacButton1.Top = 0
Title.Width = Width
Image2.Width = Width
Label1.AutoSize = True
Image2.ZOrder 0
End Sub
Private Sub image2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    On Error Resume Next
    If Button = 1 Then
    If obj <> True Then
    NewX = x
    NewY = Y
    Else
    NewX = x / 15
    NewY = Y / 15
    End If
    MakeTransparent fom.hWnd, 150
    Else
    If fom.Height > Title.Height Then
    Title.Tag = fom.Height
    fom.Height = Title.Height
    Else
    fom.Height = Title.Tag
    End If
    End If
End Sub

Private Sub image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
    fom.Move fom.Left + x - NewX, fom.Top + Y - NewY
    Image1.Left = -fom.Left
    End If
End Sub

Private Sub image2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    MakeTransparent fom.hWnd, 255
    If Shift = 2 Then
        fom.Left = Round(fom.Left / fom.Width) * fom.Width
        fom.Top = Round(fom.Top / fom.Height) * fom.Height
    End If
End Sub
Public Sub sett(fym As Object)
On Error Resume Next
Set fom = fym
    Image1.Left = -fym.Left
    icon.Picture = fom.icon
    Label1 = fym.Caption
    DrawBorder fym
    Timer1 = True
End Sub
Sub blink()
On Error Resume Next
Dim i
tmrfl.Tag = "blink"
tmrfl = True
flare.Left = -flare.Width
End Sub
Sub Un_blink()
On Error Resume Next
tmrfl.Tag = "un"
tmrfl = True
flare1.Left = UserControl.Width
End Sub

Sub DrawBorder(frm As Form)
With frm
.AutoRedraw = True
'.DrawMode = 16
frm.Line (0, 0)-(0, .Height), vbBlack
frm.Line (15, 0)-(15, .Height - 15), &HC0C0C0
frm.Line (30, 0)-(30, .Height - 30), &HE0E0E0
frm.Line (0, 0)-(.Width, 0), vbBlack
frm.Line (0, 15)-(.Width - 15, 15), &HC0C0C0
frm.Line (0, 30)-(.Width - 30, 30), &HE0E0E0
frm.Line (0, .Height - 15)-(.Width, .Height - 15), &H808080
frm.Line (0, .Height - 30)-(.Width, .Height - 30), vbBlack
frm.Line (0, .Height - 45)-(.Width, .Height - 45), &HE0E0E0
frm.Line (.Width - 15, 0)-(.Width - 15, .Height), vbBlack
frm.Line (.Width - 30, 0)-(.Width - 30, .Height), &H808080
frm.Line (.Width - 45, 0)-(.Width - 45, .Height), &HE0E0E0
'.DrawMode = 13
.AutoRedraw = False
End With
End Sub
