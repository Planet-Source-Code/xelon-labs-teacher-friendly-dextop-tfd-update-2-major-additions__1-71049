VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Make New Icon"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   LinkTopic       =   "Form2"
   ScaleHeight     =   5160
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00782F1D&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2760
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   11
      ToolTipText     =   "Position pad, Drag mouse to change position"
      Top             =   3480
      Width           =   975
      Begin VB.Shape pt 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   360
         Shape           =   2  'Oval
         Top             =   360
         Width           =   135
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.LabelText lft 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Caption         =   "Left"
      Text            =   "PreSet"
   End
   Begin Project1.LabelText tp 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   503
      Caption         =   "Top"
      Text            =   "PreSet"
   End
   Begin Project1.MacButton MacButton2 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Create"
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
      BCOL            =   15128530
      FCOL            =   0
   End
   Begin Project1.LabelText LabelText3 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      Caption         =   "Icon"
      Text            =   "Set Icon Path"
   End
   Begin Project1.LabelText LabelText2 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      Caption         =   "Path"
      Text            =   "Focus Here to enter Path"
   End
   Begin Project1.LabelText LabelText1 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      Caption         =   "Name"
      Text            =   "Enter Name"
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   529
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Position of the Icon on desktop"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00F2E2D9&
      X1              =   360
      X2              =   3720
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Icon to be viewed"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00F2E2D9&
      X1              =   360
      X2              =   3720
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00F2E2D9&
      X1              =   360
      X2              =   3720
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Path which is to be executed"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H0059341C&
      Caption         =   "Enter Data For Making new icon"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   2295
   End
   Begin Project1.aicAlphaImage aimage 
      Height          =   735
      Left            =   360
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Scaler          =   1
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00F2E2D9&
      Height          =   4215
      Left            =   240
      Top             =   600
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0059341C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F2E2D9&
      Height          =   4575
      Left            =   120
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Dim Dir As Boolean

Private Sub Aimage_Click()
On Error Resume Next
frm = 2
Form3.Show
Form3.Load_Click
End Sub

Private Sub Form_GotFocus()
On Error Resume Next
Title.blink
End Sub

Private Sub Form_LostFocus()
On Error Resume Next
Title.unblink
End Sub

Private Sub Form_Load()
On Error Resume Next
LabelText3.Set_Browse
LabelText2.Set_Browse
titlebar.sett Me
aimage.LoadImage_FromFile App.path & "\icons\address book.ico"
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
aimage.ClearImage
End Sub

Private Sub LabelText2_Browsed()
On Error Resume Next
CD.ShowOpen
If CD.filename <> "" Then
LabelText2.text = CD.filename
LabelText2.Apply
End If
End Sub

Private Sub LabelText2_Change()
On Error GoTo z
Dir = True
Dir1 = LabelText2.text
aimage.LoadImage_FromFile App.path & "\icons\Dir.ico"
Exit Sub
z:
On Error Resume Next
Dir = False
aimage.LoadImage_FromFile App.path & "\icons\File.ico"
End Sub

Private Sub LabelText2_GotFocus()
On Error Resume Next
LabelText2.Set_Browse
End Sub

Private Sub LabelText3_Browsed()
On Error Resume Next
CD.ShowSave
LabelText3.text = CD.filename
LabelText3.Apply
aimage.LoadImage_FromFile CD.filename
End Sub

Private Sub MacButton1_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
MakeINI
Me.Hide
Form1.LoadDesktop
End Sub

Private Sub MakeINI()
On Error Resume Next
Dim strsave As String
strsave = App.path & "\links\" & LabelText1.text & ".lnk"
Call WriteIni("Main", "Path", LabelText2.text, strsave)
Call WriteIni("Main", "Caption", LabelText1.text, strsave)
Call WriteIni("Main", "Marker", lft.text & "," & tp.text, strsave)
If LabelText3.text = "" Then
If Dir = True Then
Call WriteIni("Main", "Picture", App.path & "\dir.ico", strsave)
Else
Call WriteIni("Main", "Picture", App.path & "\prog.ico", strsave)
End If
Else
Call WriteIni("Main", "Picture", LabelText3.text, strsave)
End If
Call WriteIni("Main", "Key", "", strsave)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
pt.Move X - 67, Y - 67
lft.text = Fix((X * (Screen.Width) / 975) / 15)
tp.text = Fix((Y * (Screen.Height) / 975) / 15)
End If
End Sub
