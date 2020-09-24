VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Dextop Clean Wizard"
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form5"
   ScaleHeight     =   2865
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tcmp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   2280
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.MacButton MacButton3 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Do Tasks"
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
      BCOL            =   14993340
      FCOL            =   0
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0059341C&
      Caption         =   "Transfer"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Transfer all icons to.."
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0059341C&
      Caption         =   "Remove"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Remove all icons"
      Top             =   1320
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   480
      Pattern         =   "*.lnk*"
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   529
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00BB8A77&
      FillStyle       =   7  'Diagonal Cross
      Height          =   255
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop Clean Wizard Helps you to Remove All the Icons of your Desktop or to Copy the icons to specified path"
      ForeColor       =   &H00F2E2D9&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0059341C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F2E2D9&
      Height          =   2415
      Left            =   120
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Private Sub Form_Load()
On Error Resume Next
titlebar.sett Me
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
Title.blink
End Sub

Private Sub Form_LostFocus()
On Error Resume Next
Title.unblink
End Sub

Private Sub MacButton3_Click()
On Error Resume Next
Shape1.Width = 15
Shape1.Visible = True
On Error Resume Next
Dim i As Integer
Dim sey As String
file1.path = App.path & "\links\"
file1.selected(0) = True
If Option1.Value = True Then
tcmp = True
For i = 0 To file1.ListCount
sey = GetFromIni("Main", "Key", file1.path & "\" & file1.filename)
If sey = "" Then
Kill App.path & "\links\" & file1.filename
file1.selected(i) = True
End If
Next
ElseIf Option2.Value = True Then
CD.ShowSave
If CD.FileTitle <> "" Then
MkDir CD.filename
tcmp = True
For i = 0 To file1.ListCount
sey = GetFromIni("Main", "Key", file1.path & "\" & file1.filename)
If sey = "" Then
FileCopy file1.path & "\" & file1.filename, CD.filename & "\" & file1.filename
file1.selected(i) = True
End If
Next
End If
End If

End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyAscii = 13 Then
MacButton3_Click
End If
End Sub

Private Sub Option2_Click()

End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyAscii = 13 Then
MacButton3_Click
End If
End Sub

Private Sub tcmp_Timer()
On Error Resume Next
If Shape1.Width > 3255 Then
tcmp.Enabled = False
Form1.LoadDesktop
Else
Shape1.Width = Shape1.Width + 40
End If
End Sub
