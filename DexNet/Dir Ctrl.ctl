VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl List_Box 
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   ScaleHeight     =   5730
   ScaleWidth      =   6165
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   5040
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin ComctlLib.ListView lv1 
      Height          =   3615
      Left            =   1080
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "il1"
      SmallIcons      =   "il1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox picTemp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   3000
      ScaleHeight     =   360
      ScaleWidth      =   1560
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin ComctlLib.ImageList il2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList il1 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "List_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim path As String

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As typSHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal Flags&) As Long
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

Private Type typSHFILEINFO
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * 260
  szTypeName As String * 80
End Type
Public count As Integer
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const ILD_TRANSPARENT = &H1
Private Const Flags = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private FileInfo As typSHFILEINFO
Event Click()
Event DClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)




Public Sub Additem(nam As String, str As String)
On Error Resume Next
Call lv1.ListItems.Add(, str, nam, ExtractIcon(str, il1, picTemp, 32), ExtractIcon(str, il2, picTemp, 16))
End Sub
Public Property Get Dir() As String
On Error Resume Next
    Dir = path
End Property

Public Property Let Dir(ByVal New_Dir As String)
On Error Resume Next
    Navigate New_Dir
    path = New_Dir
End Property
Public Sub Navigate(str As String)
On Error Resume Next
Dim dbx As DirListBox
lv1.ListItems.Clear
file1 = str
Dim i As Integer
For X = 0 To file1.ListCount - 1
Call Add(GetFilename(file1.List(X)), file1.path & "\" & file1.List(X), file1, i)
i = i + 1
Next
    path = str
    count = lv1.ListItems.count
    PropertyChanged "Dir"
End Sub

Private Sub Add(nam As String, str As String, fil As FileListBox, idx As Integer)
On Error Resume Next
On Error GoTo X
Dir1.path = GetPath(fil.path & "\" & fil.List(idx))
Call lv1.ListItems.Add(, str, nam, ExtractIcon(fil.path & "\" & fil.List(idx), il1, picTemp, 32), ExtractIcon(fil.path & "\" & fil.List(idx), il2, picTemp, 16))
Exit Sub
X:
On Error Resume Next
Call lv1.ListItems.Add(, str, nam, ExtractIcon(fil.path & "" & fil.List(idx), il1, picTemp, 32), ExtractIcon(fil.path & "" & fil.List(idx), il2, picTemp, 16))
End Sub
Private Function GetFilename(ByVal strPath As String) As String
On Error Resume Next
    If InStrRev(strPath, "\") > 0 Then
        GetFilename = Mid$(strPath, InStrRev(strPath, "\") + 1)
    Else
        GetFilename = strPath
    End If
End Function
Private Function GetPath(ByVal strPath As String) As String
On Error Resume Next
    If InStrRev(strPath, "\") > 0 Then
        GetPath = Mid$(strPath, 1, InStrRev(strPath, "\"))
    Else
        GetPath = strPath
    End If
End Function

Private Sub Form_Load()
On Error Resume Next
End Sub

Public Sub Clear()
On Error Resume Next
lv1.ListItems.Clear
End Sub

Private Function ExtractIcon(filename As String, AddtoImageList As ImageList, PictureBox As PictureBox, PixelsXY As Integer) As Long
On Error Resume Next
    Dim SmallIcon As Long
    Dim NewImage As ListImage
    Dim IconIndex As Integer
    
    If PixelsXY = 16 Then
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_SMALLICON)
    Else
        SmallIcon = SHGetFileInfo(filename, 0&, FileInfo, Len(FileInfo), Flags Or SHGFI_LARGEICON)
    End If
    
    If SmallIcon <> 0 Then
      With PictureBox
        .Height = 15 * PixelsXY
        .Width = 15 * PixelsXY
        .ScaleHeight = 15 * PixelsXY
        .ScaleWidth = 15 * PixelsXY
        .Picture = LoadPicture("")
        .AutoRedraw = True
        
        SmallIcon = ImageList_Draw(SmallIcon, FileInfo.iIcon, PictureBox.hdc, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      
      IconIndex = AddtoImageList.ListImages.count + 1
      Set NewImage = AddtoImageList.ListImages.Add(IconIndex, , PictureBox.Image)
      ExtractIcon = IconIndex
    End If
End Function

Private Sub lv1_Click()
On Error Resume Next
RaiseEvent Click
End Sub

Private Sub lv1_DblClick()
On Error Resume Next
RaiseEvent DClick
End Sub

Private Sub lv1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub lv1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub lv1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
UserControl_Resize
lv1.ListItems.Clear
Navigate Dir
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    path = PropBag.ReadProperty("Dir", Dir)
    Dir = PropBag.ReadProperty("Dir", Dir)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
lv1.Left = 0
lv1.Top = 0
lv1.Width = UserControl.Width
lv1.Height = UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("Dir", path, "C:\")
End Sub
Public Property Get selected(i As Integer) As Boolean
On Error Resume Next
    selected = lv1.ListItems.Item(i).selected
End Property

Public Property Let selected(i As Integer, ByVal selected As Boolean)
On Error Resume Next
    lv1.ListItems.Item(i).selected = selected
End Property
Public Property Get text_of(i As Integer) As String
On Error Resume Next
    text_of = lv1.ListItems.Item(i).text
End Property

Public Property Let text_of(i As Integer, ByVal text_of As String)
On Error Resume Next
    lv1.ListItems.Item(i).text = text_of
End Property
Public Property Get text() As String
On Error Resume Next
    text = lv1.selectedItem.text
End Property

Public Property Let text(ByVal text As String)
On Error Resume Next
    lv1.selectedItem.text = text
End Property
Public Sub List_Size(large As Boolean)
On Error Resume Next
If large = True Then
lv1.SmallIcons = il1
Navigate path
Else
lv1.SmallIcons = il2
Navigate path
End If
End Sub
Public Sub View(Report As Boolean)
On Error Resume Next
If Report = True Then
lv1.View = lvwReport
Else
lv1.View = lvwList
End If
End Sub
