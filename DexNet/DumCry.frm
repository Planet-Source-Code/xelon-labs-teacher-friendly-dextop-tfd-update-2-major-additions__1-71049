VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DuCr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Encrypt/Decrypt"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2970
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox text2 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"DumCry.frx":0000
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1215
      Left            =   600
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"DumCry.frx":008B
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3000
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "DuCr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rand As Integer


Sub lode(file As String)
On Error Resume Next
Set Picture1.Picture = LoadPicture(file)
Image1 = Picture1
text2 = DeCrypt(Picture1, 1, Image1.Width + 150, -1)
End Sub


Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Text1_Change()
On Error Resume Next
Randomize 1
Image1 = Picture1
rand = Fix(Rnd * 100)
encrypt Picture1, 1, text1, rand
text2 = DeCrypt(Picture1, 1, Image1.Width + 15, rand)
If text2.text <> text1.text Then
MsgBox "Wrong"
End If
End Sub

Sub save(file As String)
On Error Resume Next
Dim rand As Integer
rand = Rnd() * 100
encrypt Picture1, 1, text2, rand
text2.SaveFile file, 1
End Sub

Sub I2S(op As String, Sv As String)
On Error Resume Next
lode op
save Sv
End Sub

Sub S2I(op As String, Sv As String)
On Error Resume Next
Dim rand As Integer
Dim pct As IPictureDisp
text2.LoadFile op, 1
Randomize 1
Image1 = Picture1
rand = Fix(Rnd * 100)
encrypt Picture1, 1, text2.text, Fix(rand)
SavePicture Picture1.Image, Sv
End Sub
