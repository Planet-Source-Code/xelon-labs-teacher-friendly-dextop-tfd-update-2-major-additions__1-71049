VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmenu 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   Picture         =   "frmmenu.frx":0000
   ScaleHeight     =   6150
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin PrjTskbr.StartMenu StartMenu1 
      Height          =   3105
      Left            =   2085
      TabIndex        =   10
      Top             =   120
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   5477
   End
   Begin VB.DriveListBox Drive2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin PrjTskbr.Search ser 
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1680
      Top             =   5640
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Text            =   "Keyword"
      ToolTipText     =   "Hit Enter to proceed"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton start 
      Caption         =   ">"
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00F2E2D9&
      Caption         =   "."
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   135
   End
   Begin PrjTskbr.UserControl1 UserControl11 
      Height          =   5175
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   9128
   End
   Begin MSComctlLib.ListView List1 
      Height          =   1455
      Left            =   2160
      TabIndex        =   6
      Top             =   3960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2566
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15917785
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D1B7AD&
      Caption         =   "Search list /\"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2175
      TabIndex        =   8
      Top             =   3690
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C1ABA2&
      BackStyle       =   0  'Transparent
      Caption         =   "|| Search ||"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   2055
      Top             =   3810
      Width           =   3375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      Height          =   1665
      Left            =   2070
      Top             =   3825
      Width           =   3345
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dragged As Boolean
Public sel As Integer
Dim idx As String

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'sel = List1.ListItems.item(List1.SelectedItem.Index).Index
On Error Resume Next
PopupMenu frmpop.mnu, , Command1.Left + 135, (Command1.Top + Top)
End Sub

Private Sub ExSearch_found(item As String)
UserControl11.Add item
End Sub


Private Sub Form_Load()
On Error Resume Next
Dim x As Integer
For x = 0 To Drive1.ListCount - 1
UserControl11.Add Left(Drive1.list(x), 2) & "\"
Next
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim z As Integer
For z = 1 To Data.Files.Count
Call UserControl11.Add(Data.Files(z))
Next
End Sub

Private Sub Label1_Click()
On Error Resume Next
If Label1 = "Search list \/" Then
Label1 = "Search list /\"
Shape1.Height = Shape1.Height + 1400
Shape2.Height = Shape2.Height + 1400
List1.Visible = True
Else
Label1 = "Search list \/"
Shape1.Height = Shape1.Height - 1400
Shape2.Height = Shape2.Height - 1400
List1.Visible = False
End If
End Sub

Private Sub Label2_Click()
On Error Resume Next
If Drive1.Visible = True Then
Drive1.Visible = False
text1.Visible = False
start.Visible = False
Else
Drive1.Visible = True
text1.Visible = True
start.Visible = True
End If
End Sub


Private Sub List1_DblClick()
On Error Resume Next
ShellFile List1.SelectedItem.text
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Dragged = True Then
Dragged = False
Timer1 = False
List1.ListItems.Add , , idx
End If
End Sub

Private Sub ser_found(item As String)
On Error Resume Next
List1.ListItems.Add , , item
End Sub


Private Sub start_Click()
On Error Resume Next
If start.Caption = ">" Then
List1.ListItems.Clear
 start.Caption = "X"
 ser.start Drive1, text1
 ElseIf start.Caption = "X" Then
   start.Caption = ">"
 ser.Stopp
End If
End Sub

Private Sub StartMenu1_btnDown(pth As String, x As Single, y As Single, Button As Integer, Shift As Integer)
Dim pt As POINTAPI
GetCursorPos pt
If Button = 2 Then PopupMenu frmpop.AProg, , pt.x * 15, (pt.y * 15) - 460
End Sub

Private Sub StartMenu1_NodeClick(pth As String)
If Left(pth, 2) <> "F:" And Left(pth, 3) <> "SF:" And Left(pth, 4) <> "SSF:" Then
Hide
ShellFile pth
End If
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
 If KeyAscii = 13 Then
 ser.start Drive1, text1
 End If
End Sub

Private Sub UserControl11_Click()
On Error Resume Next
Dim x As Integer
Dim y As Integer
frmico.load UserControl11.current
ShellFile UserControl11.current
Unload Me
End Sub

Private Sub UserControl11_InBound()
On Error Resume Next
frmtip1.Screentip UserControl11.current, 1125, frmtaskbar.Top + frmtaskbar.Height
frmtaskbar.SetFocus
End Sub

Private Sub UserControl11_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dragged = True
Timer1 = True
idx = UserControl11.current
End Sub

Private Sub UserControl11_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dragged = False
Timer1 = False
If CheckOnList = True Then
Dragged = False
Timer1 = False
List1.ListItems.Add , , idx
End If
End Sub

Private Sub UserControl11_OutBound()
On Error Resume Next
frmtip1.Hide
End Sub

Private Function CheckOnList() As Boolean
On Error Resume Next
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckOnList = (WindowFromPoint(pt.x, pt.y) = List1.hwnd)
End Function
