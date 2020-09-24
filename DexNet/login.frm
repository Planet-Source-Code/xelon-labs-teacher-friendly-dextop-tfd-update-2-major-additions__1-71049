VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00A21000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   9225
   ClientLeft      =   4545
   ClientTop       =   3660
   ClientWidth     =   12195
   LinkTopic       =   "Form7"
   ScaleHeight     =   9225
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox from 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      Picture         =   "login.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrmove 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   840
   End
   Begin Project1.LabelText Text2 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "Code"
      Text            =   "pass"
   End
   Begin Project1.LabelText Text1 
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      Caption         =   "User-ID"
      Text            =   "user"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      Picture         =   "login.frx":458E0
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin Project1.PictureButton Accessor 
      Height          =   450
      Left            =   7200
      TabIndex        =   0
      ToolTipText     =   "Access to Dextop"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   794
      Picture         =   "login.frx":7FB99
      PictureHover    =   "login.frx":821E5
      PictureDown     =   "login.frx":84831
   End
   Begin Project1.aicAlphaImage Logo 
      Height          =   4875
      Left            =   600
      Top             =   0
      Visible         =   0   'False
      Width           =   9300
      _ExtentX        =   18521
      _ExtentY        =   18521
      Image           =   "login.frx":86E7D
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xdone  As Boolean
Dim R As Integer
Private Sub Accessor_Click()
On Error Resume Next
DuCr.I2S App.path & "\User.bmp", App.path & "\User.ini"
If Text1.text = GetFromIni("Main", "UserName", App.path & "\User.ini") Then
If Text2.text = GetFromIni("Main", "Password", App.path & "\User.ini") Then
Dim shell As New shell
shell.MinimizeAll
Form1.Show
Set Me.Picture1 = Nothing
list1.Clear
Unload Me
Else
GoTo x
End If
Kill App.path & "User.ini"
Else
x:
Picture = Picture1
MsgBox "You have Entered Wrong Password", vbCritical, "Access Denied"
Picture = from
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Accessor_Click
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim ang As Long, avg As Long
If App.PrevInstance = True Then GoTo x
Text2.pword "!"
Me.Left = 0
Top = 0
Width = Screen.Width
Height = Screen.Height
xdone = True
tmrmove = True
R = 0
Logo.LoadImage_FromFile App.path & "\Images\Accessor.png"
LoadBkg from
LoadBkg Picture1
Picture = from
Exit Sub
x:
MsgBox "Your Workstation is not having the capabilities to create multiple environments", vbDefaultButton1 + vbSystemModal, "Fatal Error"
End
End Sub

Sub LoadBkg(std As PictureBox)
Dim c32 As New c32bppDIB
std.AutoRedraw = True
std.Width = Screen.Width + 75
std.Height = Screen.Height + 75
c32.InitializeDIB Screen.Width / 15, Screen.Height / 15
c32.LoadPicture_StdPicture std.Picture
Set std.Picture = Nothing
c32.Render std.hdc, 0, 0, Screen.Width / 15, Screen.Height / 15
std.Picture = std.Image
std.Refresh
std.AutoRedraw = False
Set c32 = Nothing
End Sub

Sub DragLogin()
On Error Resume Next
On Error Resume Next
Logo.Left = (Width / 2) - (Logo.Width / 2)
With Logo
.Top = -.Height
.Visible = True
Dim x As Integer
Dim i As Integer
i = -.Height
For x = -.Height To 0 Step 20
.Top = x
.Opacity = i / 32
i = i - 1
Next
.FadeInOut 100
End With
merge
xdone = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Accessor_Click
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Accessor_Click
End If
End Sub

Private Sub tmrmove_Timer()
On Error Resume Next
If xdone = False Then
tmrmove.Interval = 1
Dim i  As Integer
 If Text1.Left > Logo.Left + 2200 Then GoTo G
Text2.Left = Text2.Left + 120
Text1.Left = Text1.Left + 120
Else
DragLogin
End If
Exit Sub
G:
xdone = True
tmrmove = False
moveacc
End Sub
Sub merge()
On Error Resume Next
Text1.Left = -Text1.Width
Text2.Left = -Text2.Width
Text1.Visible = True
Text2.Visible = True
End Sub
Sub moveacc()
On Error Resume Next
Accessor.ZOrder 0
Accessor.Top = 3480
Accessor.Visible = True
Dim x As Integer
Dim i As Integer
For x = 3480 To 4320
Accessor.Top = x
Accessor.Left = Text1.Left + Text1.Width + 300
Next
x = 4320
For i = 3480 To 4320
Accessor.Top = x
x = x - 1
Next
tmrphr = True
End Sub

