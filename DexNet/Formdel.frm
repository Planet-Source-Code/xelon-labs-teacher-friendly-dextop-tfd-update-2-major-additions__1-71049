VERSION 5.00
Begin VB.Form Formdel 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Delete Selected Icon"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MacButton Delete 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Delete"
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
   Begin Project1.MacButton Cancel 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
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
      BCOL            =   14993340
      FCOL            =   0
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   529
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right-click on delete button to show properties"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Prompt 
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to Delete"
      ForeColor       =   &H00F2E2D9&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0059341C&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   120
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "Formdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long

Private Sub Cancel_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
Title.blink
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Delete_Click
End If
End Sub

Private Sub Form_LostFocus()
On Error Resume Next
Title.unblink
End Sub

Public Sub Delete_Click()
On Error Resume Next
Kill App.path & "\links\" & Form1.imgicon(Me.Tag).ToolTipText
Form1.LoadDesktop
Unload Me
End Sub
Public Sub Delete2_Click()
On Error Resume Next
Kill GetFromIni("Main", "Path", Form1.imgicon(Me.Tag).Tag)
End Sub
Private Sub Delete_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu Form4.rt, 1, Delete.Left, Delete.Top + Delete.Height
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
titlebar.sett Me
Prompt.Caption = "Are you sure you want to Delete this Shortcut"
End Sub

