VERSION 5.00
Begin VB.Form frmlk 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Lock Icon"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form5"
   ScaleHeight     =   2520
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MacButton MacButton4 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Lock"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   14862784
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton3 
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Cancel"
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
      COLTYPE         =   3
      FOCUSR          =   -1  'True
      BCOL            =   14862784
      FCOL            =   0
   End
   Begin VB.PictureBox access 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      ScaleHeight     =   975
      ScaleWidth      =   4095
      TabIndex        =   5
      Top             =   720
      Width           =   4095
      Begin Project1.LabelText LabelText3 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         Caption         =   "ID"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Application is already locked"
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.PictureBox pass 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      BorderStyle     =   0  'None
      ForeColor       =   &H0059341C&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1095
      ScaleWidth      =   4095
      TabIndex        =   2
      Top             =   600
      Width           =   4095
      Begin Project1.LabelText LabelText1 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         Caption         =   "Pass"
      End
      Begin VB.Label Prompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter Password For Locking"
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   529
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0059341C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F2E2D9&
      Height          =   2055
      Left            =   120
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmlk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Locked As Boolean
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long

Public Password As String
Private Sub Form_Load()
On Error Resume Next
titlebar.sett Me
LabelText1.Caption = "Password"
LabelText1.Apply
LabelText3.Caption = "Password"
LabelText3.Apply
LabelText3.pword "*"
LabelText1.pword "*"
If Locked = True Then
access.Visible = True
pass.Visible = False
Else
pass.Visible = True
access.Visible = False
End If
End Sub


Private Sub Form_GotFocus()
On Error Resume Next
Title.blink
LabelText3.pword "!"
LabelText1.pword "!"
End Sub

Private Sub Form_LostFocus()
On Error Resume Next
Title.unblink
End Sub

Private Sub LabelText3_Browsed()

End Sub

Private Sub LabelText3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
MacButton4_Click
End If
End Sub

Private Sub MacButton3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub MacButton4_Click()
On Error Resume Next
If Locked = True Then
If LabelText3.text = Password Then
pass.Visible = True
access.Visible = False
Locked = False
Else
MsgBox "Invalid Password", vbCritical, "Error"
End If
ElseIf Locked = False Then
DuCr.I2S App.path & "\Keys.bmp", App.path & "\Keys.ini"
WriteIni "Main", Form1.imgicon(Me.Tag).ToolTipText, LabelText1.text, App.path & "\Keys.ini"
DuCr.S2I App.path & "\Keys.ini", App.path & "\Keys.bmp"
'Kill App.path & "\keys.ini"
DoEvents
MsgBox "Password Successfully Entered To Registry"
Form1.LoadDesktop
Unload Me
End If
End Sub
