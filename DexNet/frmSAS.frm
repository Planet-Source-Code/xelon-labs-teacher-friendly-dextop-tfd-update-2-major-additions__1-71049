VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSAS 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Student Attendance System"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   LinkTopic       =   "Form7"
   ScaleHeight     =   7650
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.MacButton MacButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Save List"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15128530
      FCOL            =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0059341C&
      Caption         =   "Student Attendance System"
      ForeColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6015
      Begin MSComDlg.CommonDialog cd 
         Left            =   5400
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ComctlLib.ListView lv1 
         Height          =   5655
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   9975
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Presence"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Marks"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Use Keys, Z for Present, X for Absent, C for Leave, S for Marks Entry and A for Name Chenge, D for New Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   5295
      End
   End
   Begin Project1.title titlebar 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
   End
   Begin Project1.MacButton MacButton2 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Open List"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15128530
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton3 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "New Name"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15128530
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton4 
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Clear List"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15128530
      FCOL            =   0
   End
End
Attribute VB_Name = "frmSAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
titlebar.sett Me
End Sub

Private Sub lv1_DblClick()
Dim str As String
str = InputFrm("Enter Name of Student", "Name", "Name", "")
If str <> "" And str <> "Name" Then
lv1.selectedItem.text = str
End If
End Sub

Private Sub lv1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim str As String
Dim idx As Integer
Dim x As Integer
idx = lv1.selectedItem.Index
If KeyCode = 90 Then ' Z
lv1.ListItems(idx).SubItems(1) = "Present"
ElseIf KeyCode = 88 Then ' X
lv1.ListItems(idx).SubItems(1) = "Absent"
ElseIf KeyCode = 67 Then ' C
lv1.ListItems(idx).SubItems(1) = "Leave"
ElseIf KeyCode = 83 Then ' S
str = InputFrm("Enter Marks for Student", "Marks", "00", "")
If Val(str) <> 0 Then
lv1.ListItems(idx).SubItems(2) = str
End If
ElseIf KeyCode = 65 Then ' A
str = InputFrm("Enter Name of Student", "Name", lv1.selectedItem.text, "")
If str <> "" Then
lv1.ListItems(idx).text = str
End If
ElseIf KeyCode = 68 Then ' D
NewName
ElseIf KeyCode = 46 Then ' Del
For x = lv1.ListItems.count To 1 Step -1
If lv1.ListItems(x).selected = True Then
lv1.ListItems.Remove x
End If
Next
End If
lv1.SetFocus
For x = 1 To lv1.ListItems.count
lv1.ListItems(x).selected = False
Next
lv1.ListItems(idx).selected = True
End Sub

Private Sub MacButton1_Click()
On Error Resume Next
Dim x As Integer
cd.InitDir = App.path & "\Roll"
'cd.filename = Date & ".Roll"
cd.ShowSave
If cd.filename = "" Then Exit Sub
If LCase(Right(cd.filename, 5)) <> ".roll" Then cd.filename = cd.filename & ".Roll"
WriteIni "Data", "Students", lv1.ListItems.count, cd.filename
For x = 1 To lv1.ListItems.count
WriteIni "Data", "Name " & CStr(x), lv1.ListItems(x).text, cd.filename
WriteIni "Data", "Presence " & CStr(x), lv1.ListItems(x).SubItems(1), cd.filename
WriteIni "Data", "Marks " & CStr(x), lv1.ListItems(x).SubItems(2), cd.filename
Next
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
Dim x As Integer
Dim count As Integer
Dim dat(3) As String
cd.filename = ""
cd.InitDir = App.path & "\Roll"
cd.ShowOpen
If cd.filename = "" Then Exit Sub
lv1.ListItems.Clear
count = GetFromIni("Data", "Students", cd.filename)
For x = 1 To count
lv1.ListItems.Add , , GetFromIni("Data", "Name " & CStr(x), cd.filename)
lv1.ListItems(x).SubItems(1) = GetFromIni("Data", "Presence " & CStr(x), cd.filename)
lv1.ListItems(x).SubItems(2) = GetFromIni("Data", "Marks " & CStr(x), cd.filename)
Next
End Sub

Private Sub MacButton3_Click()
NewName
End Sub

Sub NewName()
lv1.ListItems.Add , , "Name"
lv1.ListItems(lv1.ListItems.count).SubItems(1) = "---"
lv1.ListItems(lv1.ListItems.count).SubItems(2) = "000"
End Sub

Private Sub MacButton4_Click()
Dim Res As VbMsgBoxResult
Res = MsgBox("Are you sure you want to Clear the list", vbYesNoCancel + vbQuestion, "Roll Save")
If Res = vbYes Then
lv1.ListItems.Clear
End If
End Sub
