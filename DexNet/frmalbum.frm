VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Media Album"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   LinkTopic       =   "Form6"
   ScaleHeight     =   5655
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.MacButton MacButton4 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "Add from browse"
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "-.-.-"
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
      BCOL            =   15850195
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton3 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "Remove from playlist"
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "X"
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
      BCOL            =   15850195
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton2 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Add to playlist"
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "\/"
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
      BCOL            =   15850195
      FCOL            =   0
   End
   Begin Project1.List_Box file1 
      Height          =   2635
      Left            =   2280
      TabIndex        =   13
      Top             =   315
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4657
      Dir             =   ""
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   315
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   240
      ScaleHeight     =   15
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   3240
      Width           =   375
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   3600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   4080
   End
   Begin VB.TextBox Lstdat 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Drag n' Drop List Data"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox RList 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   3000
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox Lstdir 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   3840
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Project1.MacButton MacButton7 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   5260
      Width           =   1215
      _ExtentX        =   2143
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15850195
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton6 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Save the playlist"
      Top             =   5260
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Save Playlist"
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
      BCOL            =   15850195
      FCOL            =   0
   End
   Begin Project1.MacButton MacButton5 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Load the playlist"
      Top             =   5260
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Load Playlist"
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
      BCOL            =   15850195
      FCOL            =   0
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2520
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox Album 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   0
      MultiSelect     =   2  'Extended
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   2940
      Width           =   6135
   End
   Begin Project1.title titlebar 
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   529
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   0
      Picture         =   "frmalbum.frx":0000
      Stretch         =   -1  'True
      Top             =   5300
      Width           =   6135
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewX As Long, NewY As Long, FormX As Long, FormY As Long
Dim sel As Boolean
Dim over As Boolean

Private Sub Album_DblClick()
On Error Resume Next
Lstdir.ListIndex = Album.ListIndex
Form1.MP1.Play Lstdir.text
Form1.MP1.Tag = Album.ListIndex
End Sub

Private Sub Album_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Timer1 = True Then
Timer1 = False
Lstdat.Visible = False
If over = True Then
MacButton2_Click
over = False
End If
End If
End Sub

Private Sub Album_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
For X = 1 To Data.Files.count
Album.Additem (GetFilename(Data.Files(X)))
Lstdir.Additem Data.Files(X)
Next
Lstdat.Visible = False
End Sub

Private Sub Album_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
On Error Resume Next
Dim ptXY As POINTAPI
    GetCursorPos ptXY
Lstdat.Visible = True
Lstdat.Top = (ptXY.Y) * 15 - Me.Top + 140
Lstdat.Left = (ptXY.X) * 15 - Me.Left + 130
End Sub

Private Sub Dir1_Change()
On Error Resume Next
file1.Dir = Dir1
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1 = Drive1
End Sub


Private Sub File1_DClick()
On Error Resume Next
MacButton2_Click
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
over = False
Timer1 = True
If Button = 2 Then
PopupMenu Form4.Lview
End If
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Timer1 = True Then
Timer1 = False
Lstdat.Visible = False
If over = True Then
MacButton2_Click
over = False
End If
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    NewX = X
    NewY = Y
    MakeTransparent Me.hwnd, 150
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        Me.Move Me.Left + X - NewX, Me.Top + Y - NewY
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
MakeTransparent Me.hwnd, 255
End Sub

Private Sub MacButton1_Click()
On Error Resume Next
Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
titlebar.sett Me
lskn App.path & "\new.cfg"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
End Sub

Private Sub MacButton2_Click()
On Error Resume Next
Dim X As Integer
For X = 1 To file1.count
If file1.selected(X) = True Then
Lstdir.Additem file1.Dir & "\" & file1.text_of(X)
Album.Additem file1.text_of(X)
End If
Next
End Sub
Private Sub Form_GotFocus()
On Error Resume Next
Title.blink
End Sub

Private Sub Form_LostFocus()
On Error Resume Next
Title.unblink
End Sub

Private Sub MacButton3_Click()
On Error Resume Next
Dim X As Integer
For X = 0 To Album.ListCount - 1
If Album.selected(X) = True Then
Album.RemoveItem X
Lstdir.RemoveItem X
X = X + 1
End If
Next
End Sub

Private Sub MacButton4_Click()
On Error Resume Next
CD.ShowOpen
Album.Additem CD.FileTitle
Lstdir.Additem CD.filename
End Sub

Private Sub MacButton5_Click()
On Error Resume Next
Album.Clear
Lstdir.Clear
Dim X As Integer
Dim leg As String
Dim itm As String
Dim nam As String
CD.ShowOpen
leg = GetFromIni("Item Count", "No. of Items", CD.filename)
For X = 0 To leg
itm = GetFromIni("Items Data", "Item No. " & X, CD.filename)
nam = GetFilename(itm)
Lstdir.Additem itm
Album.Additem nam
Next
End Sub
Sub lskn(path As String)
On Error Resume Next
leg = GetFromIni("Item Count", "No. of Items", path)
For X = 0 To leg
itm = GetFromIni("Items Data", "Item No. " & X, path)
nam = GetFilename(itm)
Lstdir.Additem itm
Album.Additem nam
Next
End Sub
Private Sub MacButton6_Click()
On Error Resume Next
Dim X As Integer
CD.ShowSave
WriteIni "Item Count", "No. of Items", Lstdir.ListCount - 1, CD.filename
For X = 0 To Lstdir.ListCount - 1
WriteIni "Items Data", "Item No. " & X, Lstdir.List(X), CD.filename
Next
End Sub

Private Sub MacButton7_Click()
On Error Resume Next
Me.Hide
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim ptXY As POINTAPI
    GetCursorPos ptXY
If CheckMouseOver1 = True Then
Lstdat.Visible = True
Lstdat.Top = (ptXY.Y) * 15 - Me.Top + 140
Lstdat.Left = (ptXY.X) * 15 - Me.Left + 130
over = True
Exit Sub
End If
Lstdat.Visible = False
over = False
End Sub


Private Function GetPath(ByVal strPath As String) As String
On Error Resume Next
    If InStrRev(strPath, "\") > 0 Then
        GetPath = Mid$(strPath, 1, InStrRev(strPath, "\"))
    Else
        GetPath = strPath
    End If
End Function

Private Function GetFilename(ByVal strPath As String) As String
On Error Resume Next
    If InStrRev(strPath, "\") > 0 Then
        GetFilename = Mid$(strPath, InStrRev(strPath, "\") + 1)
    Else
        GetFilename = strPath
    End If
End Function

Private Function CheckMouseOver1() As Boolean
On Error Resume Next
    Dim pt As POINTAPI
    GetCursorPos pt
    CheckMouseOver1 = (WindowFromPoint(pt.X, pt.Y) = Album.hwnd)
End Function

