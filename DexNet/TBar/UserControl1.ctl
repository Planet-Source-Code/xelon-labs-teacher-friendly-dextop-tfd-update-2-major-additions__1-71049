VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl UserControl1 
   BackColor       =   &H00BCA49A&
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4545
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   4545
   Begin VB.PictureBox page 
      BackColor       =   &H00BCA49A&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.PictureBox icon 
         Appearance      =   0  'Flat
         BackColor       =   &H00BCA49A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   -500
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   1680
      Top             =   1560
   End
   Begin MSComctlLib.ImageList il1 
      Left            =   3720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public current As String

Private Type POINTAPI
        x As Long
        y As Long
End Type

Public iX As Integer
Public iY As Integer
Event MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Event InBound()
Event OutBound()
Event Click()

Public Sub Add(str As String)
On Error Resume Next
load icon(icon.UBound + 1)
With icon(icon.UBound)
.ToolTipText = str
Call ExtractIcon(str, il1, icon(icon.UBound), 32)
.Top = (icon.UBound * 500)
.Visible = True
If icon(icon.UBound - 1).Left = 0 Then
.Left = 840
.Top = .Top
ElseIf icon(icon.UBound - 1).Left = 840 Then
.Left = 0
.Top = .Top - 500
End If
page.Height = (icon(icon.UBound).Top + icon(icon.UBound).Height) + 500
End With
End Sub
Private Sub icon_Click(Index As Integer)
On Error Resume Next
SaveDC icon(Index).hDC
RaiseEvent Click
End Sub

Private Sub icon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
RaiseEvent MouseDown(Index, Button, Shift, x, y)
End Sub

Private Sub icon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
current = icon(Index).ToolTipText
iX = icon(Index).Left
iY = icon(Index).Top
RaiseEvent MouseMove(Index, Button, Shift, x, y)
End Sub

Private Sub icon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
RaiseEvent MouseUp(Index, Button, Shift, x, y)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim x As Integer
il1.ListImages.Clear
For x = 1 To icon.UBound
If Checkicon(x) = True Then
If icon(x).BackColor <> vbHighlight Then
icon(x).BackColor = vbHighlight
Call ExtractIcon(icon(x).ToolTipText, il1, icon(x), 32)
RaiseEvent InBound
End If
Else
If icon(x).BackColor <> &HBCA49A Then
icon(x).BackColor = &HBCA49A
Call ExtractIcon(icon(x).ToolTipText, il1, icon(x), 32)
RaiseEvent OutBound
End If
End If
Next
End Sub

Private Function Checkicon(idx As Integer) As Boolean
On Error Resume Next
    Dim pt As POINTAPI
    GetCursorPos pt
    Checkicon = (WindowFromPoint(pt.x, pt.y) = icon(idx).hwnd)
End Function

Public Property Get list(idx As Integer)
On Error Resume Next
list = icon(idx).ToolTipText
End Property

