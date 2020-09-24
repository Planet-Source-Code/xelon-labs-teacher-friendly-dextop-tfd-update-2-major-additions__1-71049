VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl List 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7785
   EditAtDesignTime=   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6990
   ScaleWidth      =   7785
   Begin VB.CommandButton Command2 
      BackColor       =   &H00404040&
      Caption         =   "|"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   840
      MousePointer    =   9  'Size W E
      TabIndex        =   9
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Stopper 
      BackColor       =   &H00404040&
      Caption         =   "|"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   3720
      MousePointer    =   9  'Size W E
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00404040&
      Caption         =   "|"
      Height          =   255
      Left            =   -120
      MousePointer    =   9  'Size W E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Pannel1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pannel 1"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   -120
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Pannel2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pannel 2"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   870
      TabIndex        =   5
      Top             =   0
      Width           =   2895
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4215
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   255
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   4125
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox page 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   281
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin Project1.aicAlphaImage img 
         Height          =   975
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   975
         _extentx        =   1720
         _extenty        =   1720
         image           =   "32-list.ctx":0000
      End
      Begin VB.Shape shpsel 
         BorderColor     =   &H00000000&
         Height          =   975
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image back 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Header 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Footer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   2655
      End
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5160
      Top             =   -120
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4680
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Dim q As Integer
Dim ico As Integer
Public NoPannel As Boolean
Public currenthead As String
Attribute currenthead.VB_VarProcData = "Item_Data"
Public currentfoot As String
Attribute currentfoot.VB_VarProcData = "Item_Data"
Public pic As String
Public head As String
Public foot As String
Public currentpic As String
Public iconwidth As Long
Public labelswidth As Long
Dim i As Integer
Dim NevX As Integer
Dim NevY As Integer
Event Click()
Event Additem()
Event AddPath()
Event DClick()
Event Edit()
Event Readini()
Event WriteIni()
Event Scroll()
Event Resize()
Event Clear()
Event Export()
Event Hover()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


Public Sub Additem(pictur As String, head As String, foot As String)
On Error Resume Next
If img(img.UBound).Visible = False Then
ElseIf img(img.UBound).Visible = True Then
load img(img.UBound + 1)
load Footer(Footer.UBound + 1)
load Header(Header.UBound + 1)
End If
With img(img.UBound)
.Visible = True
.Top = img.UBound * img(img.UBound).Height
.Tag = pictur
.LoadImage_FromFile pictur
End With
With Header(Header.UBound)
.Visible = True
.Top = img.UBound * img(img.UBound).Height + 16
.Caption = head
End With
With Footer(Footer.UBound)
.Visible = True
.Top = img.UBound * img(img.UBound).Height + 40
.Caption = foot
End With
page.Height = img.UBound * img(0).Height * 15.6 + 975
VScroll1.Max = page.Height
RaiseEvent Additem
End Sub

Private Sub anim8()
On Error Resume Next
q = 1
If back.Visible = False Then
back.Visible = True
back.Top = img(ico).Top
Else
If back.Top = img(ico).Top Then
tmr.Enabled = False
Exit Sub
End If
Dim X As Integer
If back.Top = img(ico).Top Then
tmr.Enabled = False
ElseIf img(ico).Top > back.Top Then
X = back.Top
X = X + 13
back.Top = X
ElseIf img(ico).Top < back.Top Then
X = back.Top
X = X - 13
back.Top = X
End If
End If
q = 0
End Sub

Private Sub dt_Timer()
On Error Resume Next
If page.Height > UserControl.Height Then
page.Top = page.Top - 2
End If
End Sub

Private Sub back_Click()
On Error Resume Next
pic = img(ico).Tag
head = Header(ico).Caption
foot = Footer(ico).Caption
shpsel.Top = img(ico).Top
shpsel.Visible = True
RaiseEvent Click
End Sub

Private Sub back_DblClick()
On Error Resume Next
RaiseEvent DClick
End Sub

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Me.currentfoot = Footer(ico).Caption
Me.currenthead = Header(ico).Caption
Me.currentpic = img(ico).Tag
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
NevX = X
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If NoPannel = True Then Exit Sub
If Button = 1 Then
On Error GoTo err
Command2.Left = Command2.Left + X - NevX
Pannel1.Width = Command2.Left
Pannel2.Left = Command2.Left
Pannel2.Width = Command2.Width + Command2.Left - X
Stopper.Left = Pannel2.Left + Pannel2.Width
For i = 0 To img.UBound
Header(i).Left = Command2.Left / 15.6
Footer(i).Left = Command2.Left / 15.6
Header(i).Width = Pannel2.Width / 15.6
Footer(i).Width = Pannel2.Width / 15.6
Next
back.Width = Pannel2.Width / 15.6 + Header(0).Left + 10
End If
Exit Sub
err:
Command2.Left = 275
End Sub

Private Sub Footer_Click(Index As Integer)
On Error Resume Next
ico = img(Index).Index
tmr.Enabled = True

RaiseEvent Click
End Sub

Private Sub Footer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Footer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
ico = Index
tmr.Enabled = True
RaiseEvent Hover
hvr (Index)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Footer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Header_Click(Index As Integer)
On Error Resume Next
ico = img(Index).Index
tmr.Enabled = True

RaiseEvent Click
End Sub

Private Sub Header_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Header_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
ico = Index
tmr.Enabled = True
RaiseEvent Hover
hvr (Index)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Header_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
ico = Index
tmr.Enabled = True
RaiseEvent Hover
RaiseEvent MouseMove(Button, Shift, X, Y)
hvr (Index)
End Sub

Private Sub img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub page_Click()
On Error Resume Next
If q = 0 Then
tmr.Enabled = False
back.Visible = False
End If
RaiseEvent Click
End Sub

Private Sub Stopper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
NevX = X
End Sub

Private Sub Stopper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If NoPannel = True Then Exit Sub
If Button = 1 Then
On Error GoTo err
Stopper.Left = Stopper.Left + X - NevX
Pannel2.Width = Stopper.Left - Pannel2.Left
For i = 0 To img.UBound
Header(i).Left = Command2.Left / 15.6
Footer(i).Left = Command2.Left / 15.6
Header(i).Width = Pannel2.Width / 15.6
Footer(i).Width = Pannel2.Width / 15.6
Next
back.Width = Pannel2.Width / 15.6 + Header(0).Left + 10
End If
Exit Sub
err:
Stopper.Left = Pannel2.Left + 100
End Sub

Private Sub tmr_Timer()
On Error Resume Next
anim8
End Sub

Private Sub UserControl_Click()
On Error Resume Next
page_Click
page.Top = 240
Command1.Top = 0
Command2.Top = 0
Pannel1.Top = 0
Pannel2.Top = 0
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
page.Width = UserControl.Width
VScroll1.Left = page.Width - 255
VScroll1.Top = 0
VScroll1.Height = UserControl.Height
RaiseEvent Resize
End Sub
Public Sub Export(count As TextBox, Picture As TextBox, head As TextBox, foot As TextBox, idx As Integer)
On Error Resume Next
Picture.text = img(idx).Tag
head.text = Header(idx).Caption
foot.text = Footer(idx).Caption
count.text = img(img.UBound).Index
RaiseEvent Export
End Sub
Public Sub Clear()
On Error Resume Next
Dim X As Integer
For X = 1 To img.UBound
Unload img(X)
Unload Header(X)
Unload Footer(X)
Next
img(0).Visible = False
img(0).Tag = ""
Header(0).Visible = False
Footer(0).Visible = False
back.Visible = False
page.Height = img.UBound * img(0).Height + 6
RaiseEvent Clear
End Sub
Public Sub edititem(Picture As String, head As String, foot As String, idx As Integer)
On Error Resume Next
img(idx).Tag = Picture
img(idx).LoadImage_FromFile Picture
Header(idx).Caption = head
Footer(idx).Caption = foot
RaiseEvent Edit
End Sub

Public Sub AddPath(path As String, Clear As Boolean)
On Error Resume Next
Dim i As Integer
file1.Pattern = "*.gif;*.bmp;*.jpg;*.jpeg;*.ico;*.cur;*.wmf;*.emf;*.png"
file1.path = path
file1.ListIndex = 0
If Clear = True Then
Me.Clear
End If
For i = 0 To file1.ListCount - 1
If i = file1.ListCount + 1 Then Exit Sub
Me.Additem file1.path & "\" & file1.filename, file1.filename, file1.path
On Error Resume Next
file1.ListIndex = file1.ListIndex + 1
Next
RaiseEvent AddPath
End Sub

Private Sub UserControl_Show()
On Error Resume Next
UserControl_Resize
End Sub

Public Sub writeini_items(strsave As String)
On Error Resume Next
Dim X As Integer
Call WriteIni("Main", "Item Count", CStr(img.UBound), strsave)
For X = 0 To img.UBound
    Call WriteIni("Item No." & X, "Picture", CStr(img(X).Tag), strsave)
    Call WriteIni("Item No." & X, "Header", CStr(Header(X).Caption), strsave)
    Call WriteIni("Item No." & X, "Footer", CStr(Footer(X).Caption), strsave)
Next
RaiseEvent WriteIni
End Sub

Public Sub Readini_items(strsave As String)
On Error Resume Next
Dim px As String, hd As String, ft As String, X As Integer
Me.Clear
Dim counter As Integer
counter = GetFromIni("Main", "Item Count", strsave)
For X = 0 To counter
img(0).Tag = GetFromIni("Item No.0", "Picture", strsave)
img(0).LoadImage_FromFile img(0).Tag
Header(0).Caption = GetFromIni("Item No.0", "Header", strsave)
Footer(0).Caption = GetFromIni("Item No.0", "Footer", strsave)
px = GetFromIni("Item No." & X, "Picture", strsave)
hd = GetFromIni("Item No." & X, "Header", strsave)
ft = GetFromIni("Item No." & X, "Footer", strsave)
Me.Additem px, hd, ft
Next
page.Height = img.UBound * img(0).Height * 15.6 + 975

RaiseEvent Readini
End Sub

Private Sub VScroll1_Scroll()
On Error Resume Next
VScroll1.Max = page.Height
page.Top = -VScroll1.Value + 255
RaiseEvent Scroll
End Sub
Public Sub BackColor(col As Long)
On Error Resume Next
page.BackColor = col
UserControl.BackColor = col
End Sub
Private Sub hvr(idx As Integer)
On Error Resume Next
Me.currentfoot = Footer(idx).Caption
Me.currenthead = Header(idx).Caption
Me.currentpic = img(idx).Tag
End Sub
Public Sub Refresh()
On Error Resume Next
page.Height = img.UBound * img(0).Height * 15.6 + 975
End Sub
