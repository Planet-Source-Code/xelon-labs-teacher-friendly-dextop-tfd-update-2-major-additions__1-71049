VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "About Dextop"
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14280
   LinkTopic       =   "Form7"
   ScaleHeight     =   5880
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MacButton MacButton3 
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   360
      Width           =   1215
      _extentx        =   2143
      _extenty        =   450
      btype           =   4
      tx              =   "User Accounts"
      enab            =   -1
      font            =   "frmabout.frx":0000
      coltype         =   3
      focusr          =   -1
      bcol            =   13160660
      fcol            =   0
   End
   Begin Project1.MacButton about 
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      _extentx        =   1931
      _extenty        =   450
      btype           =   4
      tx              =   "About Author"
      enab            =   -1
      font            =   "frmabout.frx":002C
      coltype         =   3
      focusr          =   -1
      bcol            =   13160660
      fcol            =   0
   End
   Begin Project1.MacButton MacButton2 
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   5160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      btype           =   4
      tx              =   "System Info"
      enab            =   -1
      font            =   "frmabout.frx":0058
      coltype         =   3
      focusr          =   -1
      bcol            =   13160660
      fcol            =   0
   End
   Begin Project1.MacButton MacButton1 
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   4680
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      btype           =   4
      tx              =   "OK"
      enab            =   -1
      font            =   "frmabout.frx":0084
      coltype         =   3
      focusr          =   -1
      bcol            =   14994622
      fcol            =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0059341C&
      Caption         =   "User Accounts"
      ForeColor       =   &H00F2E2D9&
      Height          =   5415
      Left            =   9480
      TabIndex        =   15
      Top             =   360
      Width           =   4455
      Begin VB.PictureBox picuser 
         Appearance      =   0  'Flat
         BackColor       =   &H0069332C&
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   360
         ScaleHeight     =   2745
         ScaleWidth      =   3945
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   3975
         Begin Project1.MacButton MacButton5 
            Height          =   255
            Left            =   2520
            TabIndex        =   23
            Top             =   2160
            Width           =   975
            _extentx        =   1720
            _extenty        =   450
            btype           =   4
            tx              =   "Set"
            enab            =   -1
            font            =   "frmabout.frx":00B0
            coltype         =   3
            focusr          =   -1
            bcol            =   13160660
            fcol            =   0
         End
         Begin Project1.LabelText LabelText2 
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1560
            Width           =   3135
            _extentx        =   5530
            _extenty        =   450
            caption         =   "Pass"
         End
         Begin Project1.LabelText LabelText1 
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   1200
            Width           =   3135
            _extentx        =   5530
            _extenty        =   450
            caption         =   "User-ID"
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Now you have got full authority to change user name or password"
            ForeColor       =   &H00F2E2D9&
            Height          =   615
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   3135
         End
      End
      Begin Project1.LabelText access 
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   2280
         Width           =   3255
         _extentx        =   5741
         _extenty        =   450
         caption         =   "Pass"
      End
      Begin Project1.MacButton MacButton4 
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   2880
         Width           =   975
         _extentx        =   1720
         _extenty        =   661
         btype           =   4
         tx              =   "Access"
         enab            =   -1
         font            =   "frmabout.frx":00DC
         coltype         =   3
         focusr          =   -1
         bcol            =   13160660
         fcol            =   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "You must enter correct password to proceed"
         ForeColor       =   &H00F2E2D9&
         Height          =   495
         Left            =   600
         TabIndex        =   16
         Top             =   1560
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0059341C&
      Caption         =   "About Author"
      ForeColor       =   &H00F2E2D9&
      Height          =   5415
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   4455
      Begin VB.Image Image4 
         Height          =   3225
         Left            =   75
         Picture         =   "frmabout.frx":0108
         Top             =   360
         Width           =   4305
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ali Ashraf"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "ali_ashraf1129@hotmail.com"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Website is comming soon"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4680
         Width           =   2415
      End
   End
   Begin Project1.title title1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _extentx        =   8281
      _extenty        =   450
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0059341C&
      Caption         =   "About 8-X Dextop"
      ForeColor       =   &H00F2E2D9&
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   1455
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmabout.frx":1059D
         Top             =   2640
         Width           =   3735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "© 8-X Corp. 2008"
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   360
         Picture         =   "frmabout.frx":106C3
         Top             =   4320
         Width           =   720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Update Versions         :   0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Built                            :   Expert"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Version                       :   1.0.00"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   1335
         Left            =   80
         Picture         =   "frmabout.frx":10EFF
         Top             =   240
         Width           =   4305
      End
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A(1)

Private Sub access_Browsed()
On Error Resume Next
If KeyCode = 13 Then
MacButton4_Click
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Width = 9360 / 2
title1.sett Me
A(0) = Frame1.Left
A(1) = Frame2.Left
access.pword "!"
End Sub

Private Sub LabelText1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
MacButton5_Click
End If
End Sub

Private Sub LabelText2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
MacButton5_Click
End If
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
cpl "C:\WINDOWS\system32\sysdm.cpl"
End Sub

Private Sub about_Click()
On Error Resume Next
If about.Caption = "About Author" Then
Frame2.Left = A(0)
Frame1.Left = A(1)
about.Caption = "About Dextop"
Else
Frame2.Left = A(1)
Frame1.Left = A(0)
about.Caption = "About Author"
End If
End Sub

Private Sub MacButton3_Click()
On Error Resume Next
If Frame3.Left = 9480 Then
Frame3.Left = Frame1.Left
Else
Frame3.Left = 9480
End If
End Sub

Private Sub MacButton4_Click()
On Error Resume Next
DuCr.I2S App.path & "\User.bmp", App.path & "\User.ini"
If access.text = GetFromIni("Main", "Password", App.path & "\User.ini") Then
picuser.Visible = True
Else
MsgBox "You have entered wrong password", vbCritical, "Critical Error"
End If
Kill App.path & "\User.ini"
End Sub

Private Sub MacButton5_Click()
On Error Resume Next
If LabelText1.text = "" Or LabelText2.text = "" Then
MsgBox "You cannot leave blank spaces", vbCritical, "Critical Error"
Else
WriteIni "Main", "Password", LabelText2.text, App.path & "\User.ini"
WriteIni "Main", "UserName", LabelText1.text, App.path & "\User.ini"
DuCr.S2I App.path & "\User.ini", App.path & "\User.bmp"
picuser.Visible = False
Kill App.path & "\User.ini"
End If
End Sub
