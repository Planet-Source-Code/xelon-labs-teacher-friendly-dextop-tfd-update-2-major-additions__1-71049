VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl StartMenu 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   ScaleHeight     =   2175
   ScaleWidth      =   2730
   Begin MSComctlLib.ImageList diril 
      Left            =   2040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Start.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Start.ctx":0354
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Start.ctx":06A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox list3 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir3 
      Height          =   315
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   2040
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox list1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.DirListBox Dir2 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.FileListBox list2 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1931
      _Version        =   393217
      Indentation     =   353
      LineStyle       =   1
      Style           =   7
      ImageList       =   "diril"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "StartMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event DblClk(pth As String)
Event Click(pth As String)
Event NodeClick(pth As String)
Event btnDown(pth As String, x As Single, y As Single, Button As Integer, Shift As Integer)
Dim pth(500) As String
Public file As String

Private Sub tv_Click()
RaiseEvent Click(pth(tv.SelectedItem.Index))
End Sub

Private Sub tv_DblClick()
RaiseEvent DblClk(pth(tv.SelectedItem.Index))
End Sub

Private Sub tv_Expand(ByVal Node As MSComctlLib.Node)
Dim i As Integer
Dim chld As Node
For i = 1 To Node.Children
Set chld = tv.Nodes(Node.Index + i)
If chld.Image = 0 Then
ExtractIcon pth(chld.Index), diril, Picture1, 34
chld.Image = diril.ListImages.Count
DoEvents
End If
Next
End Sub

Private Sub tv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent btnDown(tv.SelectedItem.Key, x, y, Button, Shift)
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
RaiseEvent NodeClick(pth(Node.Index))
file = pth(Node.Index)
End Sub

Private Sub tv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim i As Integer
For i = 1 To Data.Files.Count
FileCopy Data.Files(i), "C:\Documents and Settings\All Users\Start Menu\Programs\" & GetFilename(Data.Files(i))
Next
LoadItems
End Sub

Private Function GetFilename(ByVal strPath As String) As String
On Error Resume Next
    If InStrRev(strPath, "\") > 0 Then
        GetFilename = Mid$(strPath, InStrRev(strPath, "\") + 1)
    Else
        GetFilename = strPath
    End If
End Function

Private Sub UserControl_Initialize()
LoadItems
End Sub

Private Sub UserControl_Resize()
tv.Width = Width
tv.Height = Height
End Sub

Function getname(str As String)
getname = Right(str, Len(str) - InStrRev(str, "\"))
End Function

Function WoExt(str As String) As String
Dim pos As Integer
pos = InStrRev(str, ".")
If pos <> 0 Then
WoExt = Left(str, pos - 1)
Else
WoExt = str
End If
End Function


Sub LoadItems()
Dir1 = "C:\Documents and Settings\All Users\Start Menu\Programs\"
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim n As Integer
Dim Fil As Integer
Dim gn As String
Dim mNode As Node
tv.Nodes.Clear
List1.path = Dir1
For j = 0 To List1.ListCount - 1
ExtractIcon List1.path & "\" & List1.list(j), diril, Picture1, 34
Set mNode = tv.Nodes.Add(, , , WoExt(List1.list(j)), diril.ListImages.Count)
pth(mNode.Index) = List1.path & "\" & List1.list(j)
Next
For i = 0 To Dir1.ListCount - 1
gn = getname(Dir1.list(i))
Set mNode = tv.Nodes.Add(, , "F:" & Dir1.list(i), gn, 1)
pth(mNode.Index) = "F:" & Dir1.list(i)
List1.path = Dir1.list(i)
DoEvents
For j = 0 To List1.ListCount - 1
Set mNode = tv.Nodes.Add("F:" & Dir1.list(i), tvwChild, , WoExt(List1.list(j)))
pth(mNode.Index) = List1.path & "\" & List1.list(j)
DoEvents
Next
Dir2.path = Dir1.list(i)
For k = 0 To Dir2.ListCount - 1
gn = getname(Dir2.list(k))
Set mNode = tv.Nodes.Add("F:" & Dir1.list(i), tvwChild, "SF:" & Dir2.list(k), gn, 1)
pth(mNode.Index) = "SF:" & Dir2.list(k)
DoEvents
List2.path = Dir2.list(k)
For l = 0 To List2.ListCount - 1
Set mNode = tv.Nodes.Add("SF:" & Dir2.list(k), tvwChild, , WoExt(List2.list(l)))
pth(mNode.Index) = List2.path & "\" & List2.list(l)
DoEvents
Next
Dir3.path = Dir2.list(k)
For m = 0 To Dir3.ListCount - 1
gn = getname(Dir3.list(m))
Set mNode = tv.Nodes.Add("SF:" & Dir2.list(k), tvwChild, "SSF:" & Dir3.list(m), gn, 1)
pth(mNode.Index) = "SSF:" & Dir3.list(m)
DoEvents
List3.path = Dir3.list(m)
For n = 0 To List3.ListCount - 1
Set mNode = tv.Nodes.Add("SSF:" & Dir3.list(m), tvwChild, , WoExt(List3.list(l)))
pth(mNode.Index) = List3.path & "\" & List3.list(n)
DoEvents
Next
Next
Next
DoEvents
Next
DoEvents
Imagize
End Sub

Sub Expand(Optional mode As Boolean = True)
Dim x As Integer
For x = 1 To tv.Nodes.Count
tv.Nodes(x).Expanded = mode
Next
End Sub

Private Sub Imagize()
On Error Resume Next
Dim x As Integer
For x = 1 To tv.Nodes.Count
If Left(tv.Nodes(x).Key, 2) = "F:" Or Left(tv.Nodes(x).Key, 3) = "SF:" Or Left(tv.Nodes(x).Key, 4) = "SSF:" Then
tv.Nodes(x).SelectedImage = 2
tv.Nodes(x).ExpandedImage = 3
tv.Nodes(x).Checked = True
End If
Next
End Sub

