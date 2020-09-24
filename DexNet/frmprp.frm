VERSION 5.00
Begin VB.Form frmcln 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Icon Properties"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form5"
   ScaleHeight     =   3900
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox indexlist 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox itemlist 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   529
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0059341C&
      Caption         =   "Properties for selected Icon"
      ForeColor       =   &H00F2E2D9&
      Height          =   3495
      Left            =   1700
      TabIndex        =   3
      Top             =   360
      Width           =   3975
      Begin Project1.MacButton MacButton1 
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   1800
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         BTYPE           =   4
         TX              =   "PreSet"
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
         BackColor       =   &H00F2E2D9&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   2880
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   9
         ToolTipText     =   "Position pad, Drag mouse to change position"
         Top             =   1920
         Width           =   975
         Begin VB.Shape pt 
            BackColor       =   &H00782F1D&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00782F1D&
            FillColor       =   &H00782F1D&
            FillStyle       =   0  'Solid
            Height          =   135
            Left            =   360
            Shape           =   2  'Oval
            Top             =   360
            Width           =   135
         End
      End
      Begin Project1.LabelText tp 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "Top"
         Text            =   "Top"
      End
      Begin Project1.LabelText lft 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "Left"
         Text            =   "Left"
      End
      Begin Project1.LabelText ico 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "Icon"
         Text            =   "Icon"
      End
      Begin Project1.LabelText pth 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "Path"
         Text            =   "Path"
      End
      Begin Project1.LabelText cpt 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         Caption         =   "Caption"
         Text            =   "Caption"
      End
      Begin Project1.MacButton MacButton3 
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   3000
         Width           =   1095
         _ExtentX        =   1931
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
         BCOL            =   15582658
         FCOL            =   0
      End
      Begin Project1.MacButton Apply 
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         ToolTipText     =   "Save Setting"
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BTYPE           =   4
         TX              =   "Apply"
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
         BCOL            =   15582658
         FCOL            =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00F2E2D9&
         X1              =   240
         X2              =   2520
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00F2E2D9&
         X1              =   240
         X2              =   2520
         Y1              =   600
         Y2              =   600
      End
      Begin Project1.aicAlphaImage Aicon 
         Height          =   975
         Left            =   2880
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1720
         Scaler          =   1
      End
   End
End
Attribute VB_Name = "frmcln"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long, Change As Boolean, dwn As Boolean

Private Sub Aicon_Click()
frm = 3
Form3.Show
Form3.Load_Click
End Sub

Private Sub Apply_Click()
On Error Resume Next
Dim strfile As String, str As String
str = Form1.imgicon(indexlist.text).ToolTipText
strfile = App.path & "\links\" & str
WriteIni "Main", "Caption", cpt.text, strfile
WriteIni "Main", "Path", pth.text, strfile
WriteIni "Main", "Picture", ico.text, strfile
WriteIni "Main", "Marker", lft.text & "," & tp.text, strfile
Form1.LoadDesktop
Change = False
End Sub

Private Sub cpt_change()
Change = True
End Sub

Public Sub Form_Load()
On Error Resume Next
dwn = False
titlebar.sett Me
Dim X As Integer
itemlist.Clear
ListIndex.Clear
For X = 1 To Form1.imgicon.UBound
If Form1.lblcaption(X).Tag = "" Then
Call indexlist.Additem(X)
Call itemlist.Additem(Form1.lblcaption(X).Caption)
End If
Next
On Error Resume Next
itemlist.ListIndex = 0
Change = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Change = True Then
If MsgBox("Save the settings", vbYesNo) = vbYes Then
Apply_Click
End If
End If
Aicon.ClearImage
Set Aicon = Nothing
End Sub

Private Sub ico_change()
Change = True
End Sub

Private Sub Itemlist_Click()
On Error Resume Next
Dim strfile As String, str As String
indexlist.ListIndex = itemlist.ListIndex
If Change = True Then
If MsgBox("Save the settings", vbYesNo) = vbYes Then
Apply_Click
End If
End If
str = Form1.imgicon(indexlist.text).ToolTipText
strfile = App.path & "\links\" & str
cpt.text = ""
pth.text = ""
ico.text = ""
lft.text = ""
tp.text = ""

cpt.text = Form1.lblcaption(indexlist.text).Caption
pth.text = GetFromIni("Main", "Path", strfile)
ico.text = GetFromIni("Main", "Picture", strfile)
                    If Right$(ico.text, 10) = " <AppPath>" Then
                    i = Left$(ico.text, Len(ico.text) - 10)
                    Aicon.LoadImage_FromFile (App.path & "\icons\" & i)
                    Else
                    Aicon.LoadImage_FromFile (ico.text)
                    End If
Marker = GetFromIni("Main", "Marker", strfile)
                        Dim x3, y3
                        x3 = Left(Marker, InStr(1, Marker, ",") - 1)
                        y3 = Right(Marker, InStr(1, Marker, ",") - 1)
                        lft.text = x3
                        If Left(y3, 1) = "," Then
                        y3 = Right(y3, Len(y3) - 1)
                        End If
                        tp.text = y3
                        Change = False
End Sub

Private Sub lft_change()
On Error Resume Next
Change = True
If dwn = False Then
pt.Left = Fix(Val(lft.text) / (((Screen.Width) / 975) / 15)) - 67
End If
End Sub

Private Sub MacButton1_Click()
lft.text = "PreSet"
tp.text = "PreSet"
End Sub

Private Sub MacButton3_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
dwn = True
Picture1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
pt.Move X - 67, Y - 67
check X, Y
lft.text = Fix(((X + 67) * (Screen.Width) / 975) / 15)
tp.text = Fix(((Y + 67) * (Screen.Height) / 975) / 15)
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
dwn = False
End Sub

Private Sub pth_change()
On Error Resume Next
Change = True
End Sub

Private Sub tp_change()
On Error Resume Next
Change = True
If dwn = False Then
pt.Top = Fix(Val(tp.text) / (((Screen.Height) / 975) / 15)) - 67
End If
End Sub

Sub check(X As Variant, Y As Variant)
On Error Resume Next
If X > Picture1.Width + 67 Then X = Picture1.Width + 67
If X < -67 Then X = -67
If Y > Picture1.Height + 67 Then Y = Picture1.Height + 67
If Y < -67 Then Y = -67
End Sub
