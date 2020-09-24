VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl TimeTable 
   BackColor       =   &H00592D2B&
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8850
   ScaleHeight     =   6405
   ScaleWidth      =   8850
   Begin MSFlexGridLib.MSFlexGrid Flex 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   9
      Cols            =   7
      BackColor       =   16776701
      BackColorFixed  =   14603987
      BackColorBkg    =   4725277
      ScrollBars      =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "TimeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event PropHit(x As Integer, y As Integer)
Public handle As Long

Private Sub Flex_DblClick()
Dim str As String
str = InputBox("Please Enter The New Period", , Flex.TextMatrix(Flex.Row, Flex.col))
Flex.TextMatrix(Flex.Row, Flex.col) = str
End Sub

Private Sub Flex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then RaiseEvent PropHit(CInt(x), CInt(y))
End Sub

Private Sub UserControl_Initialize()
Dim x As Integer
Dim y As Integer
Dim day(6) As String
day(1) = "Monday"
day(2) = "Tuesday"
day(3) = "Wednesday"
day(4) = "Thursday"
day(5) = "Friday"
day(6) = "Saturday"

Flex.TextMatrix(0, 0) = "Time Table"
For x = 1 To Flex.Rows - 1
Flex.TextMatrix(x, 0) = CStr(x)
Next
For x = 1 To Flex.Cols - 1
Flex.TextMatrix(0, x) = day(x)
Next
handle = Flex.hwnd
End Sub

Sub LoadTable(pth As String)
For x = 1 To Flex.Rows - 1
For y = 1 To Flex.Cols - 1
Flex.TextMatrix(x, y) = GetFromIni("Data", "Period " & CStr(x) & "," & CStr(y), pth)
Next
Next
End Sub

Sub SaveTable(pth As String)
For x = 1 To Flex.Rows - 1
For y = 1 To Flex.Cols - 1
WriteIni "Data", "Period " & CStr(x) & "," & CStr(y), Flex.TextMatrix(x, y), pth
Next
Next
End Sub

Private Sub UserControl_Resize()
Width = 6735
Height = 2175
End Sub
