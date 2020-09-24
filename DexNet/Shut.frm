VERSION 5.00
Begin VB.Form Shut 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Shut Down"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1200
      ScaleHeight     =   1095
      ScaleWidth      =   3135
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "!"
         TabIndex        =   10
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Processing .."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter the password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Timer tdes 
      Interval        =   200
      Left            =   5040
      Top             =   1800
   End
   Begin Project1.API API 
      CausesValidation=   0   'False
      Height          =   480
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   4560
      Top             =   1800
   End
   Begin VB.PictureBox Cont 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   2640
      Picture         =   "Shut.frx":0000
      ScaleHeight     =   2985
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   3000
      Width           =   5415
      Begin VB.PictureBox ppass 
         Appearance      =   0  'Flat
         BackColor       =   &H0059341C&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   1785
         TabIndex        =   6
         Top             =   2160
         Visible         =   0   'False
         Width           =   1815
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H0059341C&
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   0
            PasswordChar    =   "!"
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password needed :-"
            ForeColor       =   &H00F2E2D9&
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1455
         End
      End
      Begin Project1.MacButton MacButton2 
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   2520
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
         BCOL            =   13160660
         FCOL            =   0
      End
      Begin Project1.MacButton MacButton1 
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Do It"
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
         BCOL            =   13160660
         FCOL            =   0
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "Shut.frx":2B3F
         Left            =   240
         List            =   "Shut.frx":2B52
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Desc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00F2E2D9&
         Height          =   1215
         Left            =   2280
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Shut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim mX As Integer
Dim mY As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
If Cont.Visible = False Then
pic.Visible = Not pic.Visible
text2.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
n = 10
MakeTransparent Me.hWnd, 10
list1.SetFocus
FormOnTop Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Cont.Visible = False Then
pic.Visible = Not pic.Visible
text2.SetFocus
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Me.Height = Screen.Height
Me.Width = Screen.Width

pic.Left = (Width / 2) - (pic.Width / 2)
pic.Top = (Height / 2) - (pic.Height / 2)
Cont.Left = (Width / 2) - (Cont.Width / 2)
Cont.Top = (Height / 2) - (Cont.Height / 2)
End Sub



Private Sub List1_Click()
On Error Resume Next
If list1.text = "Shut Down" Then
Desc = "Shuts down all the programs and makes computer ready for turn off "
ElseIf list1.text = "Restart" Then
Desc = "Shuts down all the programs and makes computer ready for turn off and then again restart"
ElseIf list1.text = "Stand By" Then
Desc = "Leaves computer in low power mode and blanks the screen while work is done in background"
ElseIf list1.text = "Log off" Then
Desc = "Exits all programs and logs off from current user to switch to other user"
ElseIf list1.text = "Lock Screen" Then
Desc = "Lock the screen and make it to transparent mode. You can resume after entering access password"
ElseIf list1.text = "Exit 8-X Dextop" Then
Desc = "Exits the 8-X Dextop and switches to normal windows explorer"
End If
If list1.text = "Lock Screen" Then
ppass.Visible = True
Text1.SetFocus
Else
ppass.Visible = False
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 27 Then
MacButton2_Click
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
MacButton1_Click
End If
End Sub

Private Sub MacButton1_Click()
On Error Resume Next
If list1.text <> "" Then

If list1.text = "Shut Down" Then
shell "shutdown.exe"
quit
ElseIf list1.text = "Restart" Then
ExitWindowsEx (1 Or 4 Or 2), &HFFFF
quit
ElseIf list1.text = "Lock Screen" Then
Dim str As String
str = Text1
DuCr.I2S App.path & "\User.bmp", App.path & "\User.ini"
If str = GetFromIni("Main", "Password", App.path & "\User.ini") Then
Cont.Visible = False
MakeTransparent hWnd, 185
Text1 = ""
Else
Text1.ForeColor = vbRed
End If
Kill App.path & "\User.ini"
ElseIf list1.text = "Log off" Then
shell "logoff.exe"
quit
ElseIf list1.text = "Exit 8-X Dextop" Then
quit
End If
End If
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub tdes_Timer()
On Error Resume Next
Dim inte As Integer
Dim str As Long, str2 As String
str = GethWndByWinTitle("Windows Task Manager")
str2 = GethWndByWinTitle("Start Menu")
If str <> 0 Then
hwndNormal str
Me.SetFocus
End If
If str2 <> 0 Then
SendMessage str2, &H10, 0, 0
End If
If Cont.Visible = False Then
For inte = 0 To 255
If GetKeyState(CLng(inte)) < 0 Then
text2.SetFocus
pic.Visible = True
End If
Next
End If
End Sub

Private Sub Text1_Change()
On Error Resume Next
Text1.ForeColor = vbWhite

End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 38 Then
list1.ListIndex = list1.ListIndex - 1
list1.SetFocus
ElseIf KeyCode = 40 Then
list1.ListIndex = list1.ListIndex + 1
list1.SetFocus
ElseIf KeyCode = 27 Then
MacButton2_Click
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
MacButton1_Click
End If
End Sub

Private Sub Text2_Change()
On Error Resume Next
text2.ForeColor = vbBlack
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Res
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
n = n + 7
If n > 205 Then
Timer1 = False
Else
MakeTransparent hWnd, n
End If
End Sub

Sub Res()
On Error Resume Next
Dim str As String
str = text2
DuCr.I2S App.path & "\User.bmp", App.path & "\User.ini"
If str = GetFromIni("Main", "Password", App.path & "\User.ini") Then
Cont.Visible = True
MakeTransparent Me.hWnd, 205
pic.Visible = False
text2 = ""
list1.SetFocus
Else
text2.ForeColor = vbRed
text2.Visible = False
DoEvents
Randomize
snapshot App.path & "\Resource\Pic " & CStr(Rnd) & " .bmp"
DoEvents
text2.Visible = True
text2.SetFocus
End If
Kill App.path & "\User.ini"
End Sub

Sub quit()
On Error Resume Next
Form1.tbar.endit
API.TaskBarShow
End
End Sub
