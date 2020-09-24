VERSION 5.00
Begin VB.Form Frminput 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Input Data"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   LinkTopic       =   "Form7"
   ScaleHeight     =   2415
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5160
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "-------------------------"
      Top             =   1200
      Width           =   2295
   End
   Begin Project1.MacButton MacButton2 
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton1 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "OK"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.title titlebar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "------------------"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1220
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00F2E2D9&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0059341C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F2E2D9&
      Height          =   1935
      Left            =   120
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "Frminput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
titlebar.sett Me
Text1.SetFocus
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
MacButton1_Click
End If
End Sub

Private Sub MacButton1_Click()
On Error Resume Next
Hide
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
Text1.text = ""
Hide
End Sub

Private Sub Text1_GotFocus()
Label2.BackColor = 8421504
End Sub

Private Sub Text1_LostFocus()
Label2.BackColor = &HFFFFFF
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Text1.SetFocus
End Sub
