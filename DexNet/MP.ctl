VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.UserControl MP 
   BackStyle       =   0  'Transparent
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   EditAtDesignTime=   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5985
   ScaleWidth      =   6795
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3480
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3480
      Top             =   480
   End
   Begin VB.Timer movetmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   0
   End
   Begin VB.PictureBox pannel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   770
      Left            =   0
      Picture         =   "MP.ctx":0000
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   278
      TabIndex        =   1
      Top             =   2280
      Width           =   4170
      Begin VB.PictureBox bctrl 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   750
         Picture         =   "MP.ctx":A6D0
         ScaleHeight     =   255
         ScaleWidth      =   1260
         TabIndex        =   13
         Top             =   330
         Visible         =   0   'False
         Width           =   1260
         Begin Project1.Slider BSlider 
            Height          =   185
            Left            =   70
            TabIndex        =   14
            Top             =   30
            Width           =   1095
            _ExtentX        =   318
            _ExtentY        =   1931
            PictureBack     =   "MP.ctx":B7CE
            PictureProgress =   "MP.ctx":C2D2
            Bar             =   "MP.ctx":CDD6
            BarOver         =   "MP.ctx":CFFE
            BarDown         =   "MP.ctx":D226
            BackColor       =   0
            Value           =   50
            Position        =   1
         End
      End
      Begin VB.PictureBox vctrl 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         Picture         =   "MP.ctx":D44E
         ScaleHeight     =   255
         ScaleWidth      =   1260
         TabIndex        =   8
         Top             =   330
         Visible         =   0   'False
         Width           =   1260
         Begin Project1.Slider vSlider 
            Height          =   185
            Left            =   70
            TabIndex        =   9
            Top             =   30
            Width           =   1095
            _ExtentX        =   318
            _ExtentY        =   1931
            PictureBack     =   "MP.ctx":E54C
            PictureProgress =   "MP.ctx":F050
            Bar             =   "MP.ctx":FB54
            BarOver         =   "MP.ctx":FD7C
            BarDown         =   "MP.ctx":FFA4
            BackColor       =   0
            Value           =   100
            Position        =   1
         End
      End
      Begin Project1.Slider Slider1 
         Height          =   105
         Left            =   165
         TabIndex        =   7
         Top             =   15
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   185
         PictureBack     =   "MP.ctx":101CC
         PictureProgress =   "MP.ctx":122E4
         Bar             =   "MP.ctx":143FC
         BarOver         =   "MP.ctx":14690
         BarDown         =   "MP.ctx":14924
         BackColor       =   0
         Position        =   1
      End
      Begin Project1.PictureButton PictureButton1 
         Height          =   630
         Left            =   2775
         TabIndex        =   2
         Top             =   135
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1111
         Picture         =   "MP.ctx":14BB8
         PictureHover    =   "MP.ctx":1610C
         PictureDown     =   "MP.ctx":17660
      End
      Begin Project1.PictureButton PictureButton2 
         Height          =   390
         Left            =   1995
         TabIndex        =   3
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   688
         Picture         =   "MP.ctx":18BB4
         PictureHover    =   "MP.ctx":19C48
         PictureDown     =   "MP.ctx":1ACDC
      End
      Begin Project1.PictureButton PictureButton3 
         Height          =   390
         Left            =   3360
         TabIndex        =   4
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   688
         Picture         =   "MP.ctx":1BD70
         PictureHover    =   "MP.ctx":1CE04
         PictureDown     =   "MP.ctx":1DE98
      End
      Begin Project1.PictureButton PictureButton4 
         Height          =   405
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   714
         Picture         =   "MP.ctx":1EF2C
         PictureHover    =   "MP.ctx":1F180
         PictureDown     =   "MP.ctx":1F63C
      End
      Begin Project1.PictureButton PictureButton5 
         Height          =   405
         Left            =   480
         TabIndex        =   10
         Top             =   300
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   714
         Picture         =   "MP.ctx":1FAF0
         PictureHover    =   "MP.ctx":1FD44
         PictureDown     =   "MP.ctx":20200
      End
      Begin Project1.PictureButton PictureButton6 
         Height          =   405
         Left            =   840
         TabIndex        =   11
         Top             =   300
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   714
         Picture         =   "MP.ctx":206B4
         PictureHover    =   "MP.ctx":20908
         PictureDown     =   "MP.ctx":20DC4
      End
      Begin Project1.PictureButton PictureButton7 
         Height          =   405
         Left            =   1200
         TabIndex        =   12
         Top             =   300
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   714
         Picture         =   "MP.ctx":21278
         PictureHover    =   "MP.ctx":214CC
         PictureDown     =   "MP.ctx":21988
      End
   End
   Begin VB.Label Title 
      BackColor       =   &H00000000&
      Caption         =   "Title"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2030
      Width           =   4170
   End
   Begin WMPLibCtl.WindowsMediaPlayer Media 
      Height          =   2025
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Media Screen"
      Top             =   0
      Width           =   4170
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7355
      _cy             =   3572
   End
End
Attribute VB_Name = "MP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Event Previous()
Event Forward()
Event DClick()
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ty As Integer
Event InRange()
Event OutRange()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim alr As Boolean
Private Sub BSlider_Change(Value As Long)
On Error Resume Next
Media.settings.balance = Value
End Sub

Private Sub BSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Timer1 = False
End Sub

Private Sub BSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
vctrl.Visible = False
End Sub

Private Sub BSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Timer1 = True
End Sub

Private Sub Media_DoubleClick(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
On Error Resume Next
RaiseEvent DClick
End Sub

Private Sub Media_KeyPress(ByVal nKeyAscii As Integer)
On Error Resume Next
If nKeyAscii = 27 Then
Media.fullscreen = False
End If
End Sub

Private Sub Media_OpenStateChange(ByVal NewState As Long)
On Error Resume Next
movetmr = True
End Sub

Private Sub movetmr_Timer()
On Error Resume Next
Slider1.Value = Media.Controls.currentPosition
Slider1.Max = Media.currentMedia.duration
Title = Media.currentMedia.name
Media.settings.autoStart = True
End Sub

Private Sub pannel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub pannel_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Picture1_Click()
On Error Resume Next

End Sub

Private Sub PictureButton1_Click()
On Error Resume Next
If Media.playState = wmppsPlaying Then
Media.Controls.pause
Else
Media.Controls.Play
Timer3 = True
End If
End Sub

Private Sub PictureButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PictureButton2_Click()
On Error Resume Next
RaiseEvent Previous
End Sub

Private Sub PictureButton2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PictureButton3_Click()
On Error Resume Next
RaiseEvent Forward
End Sub

Private Sub PictureButton3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PictureButton4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
vctrl.Visible = True
bctrl.Visible = False
BSlider.Visible = False
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub PictureButton5_Click()
On Error Resume Next
On Error Resume Next
Media.fullscreen = True
End Sub

Private Sub PictureButton6_Click()
On Error Resume Next
mpStop
End Sub

Private Sub PictureButton7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
bctrl.Visible = True
vctrl.Visible = False
BSlider.Visible = True
RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Media.Controls.currentPosition = Slider1.Value
End If
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
vctrl.Visible = False
bctrl.Visible = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If CheckMouseOver = True Then
If alr = True Then
RaiseEvent InRange
alr = False
End If
Else
If alr = False Then
RaiseEvent OutRange
alr = True
End If
End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
If Media.playState = wmppsStopped Then
RaiseEvent Forward
End If
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
Slider1.ToolTipText = "Time Line Bar" & vbCrLf & _
"Drag to slide the player time"
PictureButton1.ToolTipText = "Play Button " & vbCrLf & _
"Click to swap Pause or Play state"
PictureButton2.ToolTipText = "Back Button " & vbCrLf & _
"Click to Go Back one level in album"
PictureButton3.ToolTipText = "Forward Button " & vbCrLf & _
"Click to Go next one level in album"
PictureButton5.ToolTipText = "Full Screen " & vbCrLf & _
"Click to have full screen view of media"
PictureButton6.ToolTipText = "Stop Button " & vbCrLf & _
"Click to Stop current media"
vSlider.ToolTipText = "Volume Bar " & vbCrLf & _
"Drag to set volume"
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
On Error Resume Next
If nKeyAscii = 27 Then
Media.fullscreen = False
End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = 4170
Media.Width = 4170
pannel.Width = 4170
Title.Width = 4170
pannel.Top = UserControl.Height - pannel.Height
Title.Top = pannel.Top - Title.Height
Media.Height = Title.Top
End Sub

Public Sub SetURL(URL As String)
On Error Resume Next
Media.URL = URL
End Sub

Public Sub Play(URL As String)
On Error Resume Next
Media.URL = URL
Media.Controls.Play
End Sub

Private Sub vctrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub vSlider_Change(Value As Long)
On Error Resume Next
vctrl.Visible = True
Media.settings.volume = Value
End Sub

Private Sub vSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Timer1 = False
End Sub

Private Sub vSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Public Sub M_visible(choice As Boolean)
On Error Resume Next
Media.Visible = choice
End Sub

Private Function CheckMouseOver() As Boolean
On Error Resume Next
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.X, pt.Y) = UserControl.hwnd)
End Function

Private Sub vSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Timer1 = True
End Sub

Public Sub mpStop()
On Error Resume Next
Timer3 = False
Media.Controls.stop
End Sub

Function fullscreen() As Boolean
fullscreen = Media.fullscreen
End Function
