VERSION 5.00
Begin VB.Form frmtip1 
   BackColor       =   &H00B38279&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image img 
      Height          =   240
      Left            =   0
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmtip1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
FormOnTop Me.hwnd
End Sub
Public Sub Screentip(txt As String, x As Integer, y As Integer)
On Error Resume Next
Left = x
Top = y
lbl.AutoSize = True
lbl = txt
Frmtip2.Hide
frmtip1.Show
Width = lbl.Width + 16
img.Width = Width
End Sub

