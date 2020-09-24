VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmico 
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   LinkTopic       =   "Form6"
   ScaleHeight     =   4305
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   840
      Top             =   0
   End
   Begin VB.PictureBox ico 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin MSComctlLib.ImageList il1 
      Left            =   2160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Me.Width = 480
Me.Height = 480
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Timer1 = False
fade Me.hwnd, 1
Me.Hide
MakeOpaque Me.hwnd
End Sub
Public Sub load(file As String)
On Error Resume Next
il1.ListImages.Clear
Show
Call ExtractIcon(file, il1, ico, 32)
frmico.Top = frmmenu.UserControl11.iY + frmmenu.UserControl11.Top + frmmenu.Top
frmico.Left = frmmenu.UserControl11.iX + frmmenu.UserControl11.Left + frmmenu.Left
Timer1 = True
End Sub
