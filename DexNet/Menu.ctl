VERSION 5.00
Begin VB.UserControl UserControl1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   3435
   ScaleWidth      =   2715
   Begin VB.Timer anit 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   2280
   End
   Begin VB.PictureBox tall 
      Appearance      =   0  'Flat
      BackColor       =   &H00E2C9C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1695
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   0
      Width           =   375
      Begin VB.PictureBox pic 
         BackColor       =   &H00E2C9C0&
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   0
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   1
         Top             =   50
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.Image roller 
      Height          =   315
      Left            =   360
      Picture         =   "Menu.ctx":0000
      Top             =   1920
      Width           =   3750
   End
   Begin VB.Image img 
      Height          =   255
      Left            =   600
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   2175
   End
   Begin Project1.aicAlphaImage shine 
      Height          =   2145
      Left            =   360
      Top             =   0
      Width           =   2325
      _ExtentX        =   5583
      _ExtentY        =   5583
      Image           =   "Menu.ctx":3DF4
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Event Click(str As String)
Event MouseMove()
Event DClick()
Event Add()
Event Remove()
Public MItem As String
Public color As UserControl1

Public Sub Additem(Caption As String, Image As Image)
On Error Resume Next
Dim x As Integer
If Label(0).Visible = True Then
load pic(pic.UBound + 1)
load Label(Label.UBound + 1)
End If
pic(pic.UBound).Top = pic.UBound * (pic(0).Height + 50) + 50
pic(pic.UBound).Picture = Image.Picture
Label(Label.UBound).Caption = Caption
pic(pic.UBound).Visible = True
Label(Label.UBound).Visible = True
Label(Label.UBound).ZOrder 0
pic(pic.UBound).ZOrder 0
tall.Visible = True
UserControl.Width = 2700
pic(pic.UBound).Visible = True
roller.ZOrder 0
If Label.UBound = 0 Then
Label(Label.UBound).Top = 60
Else
Label(Label.UBound).Top = pic(pic.UBound).Top + 50
End If
UserControl.Height = Label(Label.UBound).Height + Label(Label.UBound).Top + 180
tall.Height = UserControl.Height
RaiseEvent Add
End Sub

Private Sub anit_Timer()
On Error Resume Next
If img.Left > Width Then GoTo B
img.Left = img.Left + 140
img.ZOrder 0
Exit Sub
B:
anit = False
img.Visible = False
End Sub

Private Sub Label_Click(Index As Integer)
On Error Resume Next
MItem = Label(Index).Caption
RaiseEvent Click(Label(Index).Caption)
End Sub

Private Sub Label_DblClick(Index As Integer)
On Error Resume Next
RaiseEvent DClick
End Sub

Private Sub Label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

roller.Visible = True
If roller.Top <> Label(Index).Top Then
anit.Tag = Index
img.Picture = pic(Index)
img.Visible = True
img.Left = 0
img.Top = Label(Index).Top
anit = True
End If
roller.Top = Label(Index).Top
roller.Tag = Index
Label(Index).ZOrder 0
End Sub

Private Sub Label1_Click()
On Error Resume Next

End Sub

Private Sub pic_Click(Index As Integer)
On Error Resume Next
MItem = Label(Index).Caption
RaiseEvent Click(Label(Index).Caption)
End Sub

Private Sub pic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
roller.Visible = True
roller.Top = Label(Index).Top
End Sub

Private Sub Roller_Click()
On Error Resume Next
On Error Resume Next
MItem = Label(roller.Tag).Caption
RaiseEvent Click(Label(roller.Tag).Caption)
End Sub

Private Sub UserControl_Click()
On Error Resume Next
RaiseEvent Click("")
End Sub
Public Sub Set_BackColor(color As Long)
On Error Resume Next
UserControl.BackColor = color
End Sub
Public Sub Set_ForeColor(color As Long)
On Error Resume Next
For x = 0 To Label.UBound
Label(x).ForeColor = color
Next
End Sub
Sub fd()
On Error Resume Next
shine.Opacity = 0
shine.FadeInOut 100
End Sub

Private Sub UserControl_Initialize()
shine.LoadImage_FromFile App.path & "\Images\Shine.png"
shine.AutoSize = True
End Sub
