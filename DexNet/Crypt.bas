Attribute VB_Name = "Crypt"
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Sub encrypt(pict As PictureBox, Height As Long, strng As String, rd As Integer)
Dim i As Integer, n As Integer, str As String, X As Integer
On Error Resume Next
Set pict.Picture = Nothing
pict.Width = Val(Len(strng) * 15) + 190
pict.Height = 1
SetPixel pict.hdc, 0, 0, rd
For X = 0 To Len(strng) - 1
str = Left(strng, X + 1)
str = Right(str, 1)
For i = 0 To 255
If Chr(i) = str Then n = i
Next
For H = 0 To Height
SetPixel pict.hdc, X + 1, H, (n * rd) + 1
SetPixel pict.hdc, X + 1, n, &HF2E2D9
Next
Next
End Sub

Function DeCrypt(pict As PictureBox, start As Integer, wdth As Long, pass As Integer) As String
On Error GoTo z
Dim str As String, X As Integer, str2 As String, w2 As Long
str2 = ""
X = start
w2 = pict.Width
pict.Width = wdth
If pass = -1 Then pass = GetPixel(pict.hdc, 0, 0)
While GetPixel(pict.hdc, X, 0) <> 13160660
str2 = str2 & Chr((GetPixel(pict.hdc, X, 0) - 1) / pass)
X = X + 1
Wend
GoTo X
z:
MsgBox "Fatal Error, Cannot retrieve critical Data", vbCritical, "System Overflow"
Shut.quit
X:
pict.Width = w2
DeCrypt = str2
End Function
