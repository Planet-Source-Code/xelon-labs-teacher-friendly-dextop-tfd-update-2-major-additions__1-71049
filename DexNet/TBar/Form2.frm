VERSION 5.00
Begin VB.Form Frmtip2 
   BackColor       =   &H00B38279&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   LinkTopic       =   "Form2"
   ScaleHeight     =   255
   ScaleWidth      =   3660
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
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Frmtip2"
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
Frmtip2.Left = x
Frmtip2.Top = y
Frmtip2.lbl = txt
Frmtip2.lbl.Width = Me.TextWidth(Frmtip2.lbl) + 16
Frmtip2.Width = Frmtip2.lbl.Width + 16
frmtip1.Visible = False
Frmtip2.Visible = True
img.Width = Width
End Sub

