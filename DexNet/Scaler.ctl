VERSION 5.00
Begin VB.UserControl Scale 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin Project1.MacButton MacButton1 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "<+>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin Project1.MacButton Command2 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "\/"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin Project1.MacButton Command1 
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   4
      TX              =   ">"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      FCOL            =   0
   End
   Begin VB.PictureBox WE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   720
      ScaleHeight     =   0
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.PictureBox UD 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   360
      ScaleHeight     =   825
      ScaleWidth      =   0
      TabIndex        =   2
      Top             =   720
      Width           =   15
   End
   Begin VB.TextBox v1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "Value1"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox v2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "Value2"
      Top             =   1125
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   720
      Y1              =   720
      Y2              =   840
   End
End
Attribute VB_Name = "Scale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim NewX As Integer
Dim NewY As Integer

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
NewX = 1560
Command1.Left = 1560
WE.Width = 855
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
NewX = X
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Command2.Top = Command2.Top + Y - NewY
UD.Height = Command2.Top - UD.Top
v2 = Command2.Top
End If
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
NewY = 1560
Command2.Top = 1560
UD.Height = 855
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
NewX = X
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Command1.Left = Command1.Left + X - NewX
WE.Width = Command1.Left - WE.Left
v1 = Command1.Left
End If
End Sub
Public Function Val1()
Set Val1 = v2
End Function
Public Function Val2()
Set Val2 = v1
End Function

