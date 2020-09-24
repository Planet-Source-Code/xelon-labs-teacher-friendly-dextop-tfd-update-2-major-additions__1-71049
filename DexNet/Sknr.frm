VERSION 5.00
Begin VB.Form Sknr 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Skin Dextop"
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form7"
   ScaleHeight     =   10200
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   4185
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   4215
      Begin VB.DirListBox Dir1 
         Height          =   2340
         Left            =   240
         TabIndex        =   19
         Top             =   555
         Width           =   3735
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin Project1.MacButton MacButton1 
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         ToolTipText     =   "Add the selected skin to list"
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BTYPE           =   4
         TX              =   "Add Skin ^"
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
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin Project1.MacButton Skin 
         Height          =   255
         Left            =   2160
         TabIndex        =   17
         ToolTipText     =   "Skin-up the dextop"
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   4
         TX              =   "Skin"
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
         BCOL            =   12632256
         FCOL            =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Explore Skins"
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   4185
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ListBox SknLst 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin Project1.MacButton MacButton6 
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Save the Selected Skin"
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         BTYPE           =   4
         TX              =   "Save"
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
         BCOL            =   15592683
         FCOL            =   0
      End
      Begin Project1.MacButton MacButton5 
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         BTYPE           =   2
         TX              =   "Ref"
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
         BCOL            =   15592683
         FCOL            =   0
      End
      Begin Project1.MacButton MacButton4 
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         ToolTipText     =   "Remove"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         BTYPE           =   2
         TX              =   "[X]"
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
         BCOL            =   15592683
         FCOL            =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Skin List"
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
   End
   Begin Project1.MacButton MacButton8 
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   ">"
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
   Begin Project1.MacButton MacButton9 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   ">"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2985
      ScaleWidth      =   4185
      TabIndex        =   5
      Top             =   360
      Width           =   4215
      Begin Project1.PictureButton PictureButton2 
         Height          =   225
         Left            =   3480
         TabIndex        =   6
         Top             =   1440
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   397
         Picture         =   "Sknr.frx":0000
         PictureHover    =   "Sknr.frx":07D4
         PictureDown     =   "Sknr.frx":0FA8
      End
      Begin Project1.PictureButton PictureButton1 
         Height          =   465
         Left            =   360
         TabIndex        =   7
         Top             =   315
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   820
         Picture         =   "Sknr.frx":177C
         PictureHover    =   "Sknr.frx":3C24
         PictureDown     =   "Sknr.frx":60CC
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Skin Preview"
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   1800
         Picture         =   "Sknr.frx":8574
         Stretch         =   -1  'True
         Top             =   330
         Width           =   2295
      End
      Begin VB.Image Image3 
         Height          =   315
         Left            =   1200
         Picture         =   "Sknr.frx":99F2
         Top             =   1440
         Width           =   19200
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   360
         Stretch         =   -1  'True
         Top             =   315
         Width           =   3735
      End
   End
   Begin Project1.MacButton MacButton3 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "Toggle to next skin in list"
      Top             =   9720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Toggle"
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
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton2 
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   9720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Done"
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
      BCOL            =   12632256
      FCOL            =   0
   End
   Begin Project1.title titlebar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
   End
End
Attribute VB_Name = "Sknr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer

Private Sub Dir1_Change()
On Error Resume Next
PictureButton1.loadpics Dir1, "Start"
Image2 = LoadPicture(Dir1 & "\Bar.bmp")
Image1 = LoadPicture(Dir1 & "\WP.jpg")
Image3 = LoadPicture(Dir1 & "\Title.bmp")
PictureButton2.loadpics Dir1, "Close"
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1 = Drive1
End Sub

Private Sub Form_Load()
On Error Resume Next
titlebar.sett Me
Image1.Picture = Form1.Picture
n = 1
Height = 3925
Get_INIList App.path & "\items.cfg", SknLst
'SknLst.Additem "Default <SknDir>"
End Sub

Private Sub Form_Resize()
On Error Resume Next
MacButton2.Top = Height - 105 - 375
MacButton3.Top = Height - 105 - 375
End Sub

Private Sub MacButton1_Click()
On Error Resume Next
SknLst.Additem Dir1
Kill App.path & "\items.cfg"
Write_INIList App.path & "\items.cfg", SknLst
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
Hide
End Sub

Private Sub MacButton3_Click()
On Error Resume Next
toggle
End Sub

Private Sub MacButton4_Click()
On Error Resume Next
SknLst.RemoveItem SknLst.ListIndex
Write_INIList App.path & "\items.cfg", SknLst
End Sub

Private Sub MacButton5_Click()
On Error Resume Next
SknLst.Clear
Get_INIList App.path & "\items.cfg", SknLst
End Sub

Private Sub MacButton6_Click()
On Error Resume Next
Write_INIList App.path & "\items.cfg", SknLst
End Sub


Private Sub MacButton8_Click()
On Error Resume Next
If Picture3.Visible = True Then
Picture3.Visible = False
n = n - 1
Else
Picture3.Visible = True
n = n + 1
End If
arr
End Sub
Private Sub MacButton9_Click()
On Error Resume Next
If Picture2.Visible = True Then
Picture2.Visible = False
Picture3.Top = Picture2.Top
n = n - 1
Else
Picture2.Visible = True
n = n + 1
Picture3.Top = 6600
End If
arr
End Sub
Private Sub Skin_Click()
On Error Resume Next
sknup Dir1
End Sub
Sub sknup(path As String)
On Error Resume Next
Form1.tbar.Skin path
Form1.LoadBkg path & "\WP.jpg"
Form2.titlebar.Skin path
Form3.titlebar.Skin path
Form5.titlebar.Skin path
Form6.titlebar.Skin path
frmTT.titlebar.Skin path
frmSAS.titlebar.Skin path
Formdel.titlebar.Skin path
Frminput.titlebar.Skin path
frmcln.titlebar.Skin path
frmcalc.titlebar.Skin path
frmcst.titlebar.Skin path
frmlk.titlebar.Skin path
Sknr.titlebar.Skin path
frmabout.title1.Skin path

End Sub

Sub toggle()
On Error GoTo W
SknLst.ListIndex = SknLst.ListIndex + 1
sknup SknLst.text
Exit Sub
W:
On Error GoTo t
sknup SknLst.text
Exit Sub
t:
SknLst.ListIndex = 0
sknup SknLst.text
End Sub

Private Sub SknLst_DblClick()
On Error Resume Next
Dim str As String
                    If Right$(SknLst.text, 9) = " <SknDir>" Then
                    str = Left$(SknLst.text, Len(SknLst.text) - 9)
                    sknup App.path & "\Skins\" & str & "\"
                    Else
                    If fso.FolderExists(SknLst.text) = True Then
                    sknup SknLst.text
                    Else
                    MsgBox "Skin '" & SknLst.text & "' not found, Refer to HTML Help file", vbCritical, "Error"
                    End If
                    End If
End Sub

Sub arr()
On Error Resume Next
Height = (n * Picture1.Height) + 255 + 360 + 120 + 375
End Sub

