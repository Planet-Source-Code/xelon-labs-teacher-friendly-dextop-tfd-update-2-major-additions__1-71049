VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Search 
   ClientHeight    =   7920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   ScaleHeight     =   7920
   ScaleWidth      =   8775
   Begin VB.CommandButton Command3 
      Caption         =   "Searcher"
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   8
      Text            =   "*.*"
      Top             =   360
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   285
      Left            =   7740
      TabIndex        =   7
      Top             =   360
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   360
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   285
      Left            =   7050
      TabIndex        =   5
      Top             =   360
      Width           =   555
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3150
      Hidden          =   -1  'True
      Left            =   120
      System          =   -1  'True
      TabIndex        =   4
      Top             =   3840
      Width           =   3675
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   120
      TabIndex        =   3
      Top             =   7080
      Width           =   3675
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3675
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3705
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Searcher.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Searcher.ctx":005E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7305
      Left            =   3960
      TabIndex        =   0
      Top             =   720
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   12885
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Path"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim tago As String
Event Complete()
Dim direct As String
Public running As Boolean
Event found(item As String)

Private Sub Command1_Click()
On Error GoTo bere
Dim num As Long
running = True
tago = "Start"

ListView1.ListItems.Clear
List1.Clear



List1.AddItem Dir1.path
List1.ListIndex = 0
'==============================================


ere:
If tago = "Stop" Then GoTo bere


Dir1.path = List1.text
For a = 0 To Dir1.ListCount - 1
List1.AddItem Dir1.list(a)
If InStr(1, UCase(GetLast(Dir1.list(a))), UCase(text1.text)) > 0 Then
Set lv = ListView1.ListItems.Add(, , GetLast(Dir1.list(a)), 1, 1)
RaiseEvent found(Dir1.path & "\" & GetLast(Dir1.list(a)))
lv.ListSubItems.Add , , Dir1.path
End If
DoEvents
Next a
For i = 0 To File1.ListCount - 1
If InStr(1, UCase(File1.list(i)), UCase(text1.text)) > 0 Then
Set lv = ListView1.ListItems.Add(, , File1.list(i), 2, 2)
RaiseEvent found(Dir1.path & "\" & File1.list(i))
lv.ListSubItems.Add , , Dir1.path
End If
DoEvents
Next i

List1.ListIndex = List1.ListIndex + 1
GoTo ere

bere:
RaiseEvent Complete
running = False

End Sub

Private Sub Command2_Click()
On Error Resume Next
tago = "Stop"
End Sub

Private Sub Dir1_Change()
On Error Resume Next
File1.path = Dir1.path
Dir1.Refresh
File1.Refresh
End Sub

Private Sub Text2_Change()
On Error Resume Next
On Error Resume Next
File1.Pattern = Text2.text
End Sub


Function GetLast(dir As String) As String
On Error Resume Next
txt = Split(dir, "\")
GetLast = txt(UBound(txt))
End Function

Private Sub Drive1_Change()
On Error Resume Next
Dir1.path = Drive1.Drive
Dir1.Refresh
File1.Refresh
End Sub

Sub start(dir As String, str As String)
On Error Resume Next
text1 = str
Dir1 = dir
Command1_Click
End Sub
Sub Stopp()
On Error Resume Next
Command2_Click
End Sub
Public Function gett(idx As Integer)
On Error Resume Next
gett = ListView1.ListItems(idx).text
End Function

Public Function many(idx As Integer)
On Error Resume Next
many = ListView1.ListItems.Count
End Function

Private Sub UserControl_Resize()
On Error Resume Next
Width = 1335
Height = 735
End Sub

