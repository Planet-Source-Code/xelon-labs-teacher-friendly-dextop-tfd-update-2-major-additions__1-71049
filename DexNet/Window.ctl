VERSION 5.00
Begin VB.UserControl UserControl1 
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2985
   DrawStyle       =   14642  'Solid
   DrawWidth       =   50
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
   Begin VB.PictureBox ico 
      DrawStyle       =   16  'Solid
      DrawWidth       =   -3556
      Height          =   375
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim NewX As Single
Dim NewY As Single
Dim q As Integer

Private Sub ico_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
NewX = X
NewY = Y
q = 1
End Sub

Private Sub ico_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If q = 1 Then
ico.Move ico.Left + X - NewX, ico.Top + Y - NewY
End If
End If
End Sub

Private Sub ico_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
q = 0
End Sub
