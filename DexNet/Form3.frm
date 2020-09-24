VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Select Icon"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   LinkTopic       =   "Form3"
   ScaleHeight     =   5805
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MacButton Set 
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      ToolTipText     =   "Set the icon"
      Top             =   5400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Set Image Icon"
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
      BCOL            =   10314831
      FCOL            =   0
   End
   Begin Project1.MacButton load 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "reload the list"
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Reload"
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
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   529
   End
   Begin Project1.List List1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   310
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9340
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Public already As Boolean

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

Private Sub List1_DClick()
On Error Resume Next
Set_Click
End Sub

Public Sub Load_Click()
On Error Resume Next
List1.Clear
List1.AddPath App.path & "\icons", True
List1.BackColor &HFFFFFF
List1.Refresh
End Sub


Private Sub Set_Click()
On Error Resume Next
If frm = 2 Then
Form2.LabelText3.text = List1.head & " <AppPath>"
Form2.LabelText3.Apply
Form2.aimage.LoadImage_FromFile App.path & "\icons\" & List1.head
Else
frmcln.Aicon.LoadImage_FromFile App.path & "\icons\" & List1.head
frmcln.ico.text = List1.head & " <AppPath>"
End If
Me.Hide
End Sub
