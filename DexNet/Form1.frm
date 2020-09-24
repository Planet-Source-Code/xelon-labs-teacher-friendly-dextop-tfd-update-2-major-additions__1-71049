VERSION 5.00
Object = "*\ATBar\Project1.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00F2E8E1&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   LinkTopic       =   "Form7"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PrjTskbr.UserControl2 tbar 
      Height          =   1095
      Left            =   3360
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
   End
   Begin Project1.TimeTable TimeTable1 
      Height          =   2175
      Left            =   4440
      TabIndex        =   16
      Top             =   3720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3836
   End
   Begin VB.ComboBox list1 
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7440
      Top             =   6840
   End
   Begin VB.Timer tmranim8 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   120
      Top             =   2880
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000006&
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer grd 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   2400
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1665
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox listsel 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.FileListBox file1 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   6240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Project1.PictureButton clos 
      Height          =   750
      Left            =   240
      TabIndex        =   5
      Top             =   7560
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1323
      Picture         =   "Form1.frx":0000
      PictureHover    =   "Form1.frx":1E04
      PictureDown     =   "Form1.frx":3C08
   End
   Begin Project1.MP MP1 
      Height          =   2775
      Left            =   2400
      TabIndex        =   6
      Top             =   5520
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   4895
   End
   Begin Project1.UserControl1 MENU2 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4683
   End
   Begin Project1.UserControl1 MENU1 
      Height          =   2655
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4683
   End
   Begin Project1.API API 
      Height          =   480
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin Project1.aicAlphaImage aicMenu 
      Height          =   1095
      Left            =   8520
      Top             =   2280
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   1931
      Image           =   "Form1.frx":5A0C
   End
   Begin VB.Label lblPer 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   3720
      Width           =   375
   End
   Begin Project1.aicAlphaImage aicAlphaImage1 
      Height          =   1920
      Left            =   6840
      Top             =   6240
      Width           =   1920
      _ExtentX        =   3413
      _ExtentY        =   3413
      Image           =   "Form1.frx":5A24
   End
   Begin Project1.aicAlphaImage aicSecond 
      Height          =   2040
      Left            =   6720
      ToolTipText     =   "Glass Clock"
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3598
      Image           =   "Form1.frx":5A3C
   End
   Begin VB.Label Bag 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   1920
      MousePointer    =   10  'Up Arrow
      TabIndex        =   15
      Tag             =   "Add New"
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label Cafe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   4080
      MousePointer    =   10  'Up Arrow
      TabIndex        =   14
      Tag             =   "Add New"
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label tasks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   6480
      MousePointer    =   10  'Up Arrow
      TabIndex        =   13
      Tag             =   "Add New"
      Top             =   3000
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      Visible         =   0   'False
      X1              =   0
      X2              =   1
      Y1              =   16
      Y2              =   17
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      Visible         =   0   'False
      X1              =   0
      X2              =   1
      Y1              =   16
      Y2              =   17
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      Visible         =   0   'False
      X1              =   0
      X2              =   1
      Y1              =   16
      Y2              =   17
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      Visible         =   0   'False
      X1              =   0
      X2              =   1
      Y1              =   16
      Y2              =   17
   End
   Begin VB.Label s2 
      BackStyle       =   0  'Transparent
      Height          =   8175
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label s1 
      BackStyle       =   0  'Transparent
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.Shape shpsel 
      BorderStyle     =   2  'Dash
      Height          =   735
      Index           =   0
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin Project1.aicAlphaImage aicHour 
      Height          =   2100
      Left            =   6720
      Top             =   6240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   3704
      Image           =   "Form1.frx":5A54
   End
   Begin Project1.aicAlphaImage aicMinute 
      Height          =   2055
      Left            =   6720
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   3625
      Image           =   "Form1.frx":5A6C
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   240
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin Project1.aicAlphaImage imgicon 
      Height          =   720
      Index           =   0
      Left            =   240
      Top             =   960
      Visible         =   0   'False
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      Image           =   "Form1.frx":5A84
      OLEdrop         =   1
   End
   Begin VB.Label lblcaption 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   1260
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hold As Boolean
Dim NewX As Long
Dim NewY As Long
Dim selected As Boolean
Private Whwnd As Long
Dim ShapePlace(5) As Boolean
Dim ShapeNumber As Integer
Dim mX As Long
Dim mY As Long
Dim SelX As Long
Dim SelY As Long
Dim tm As Boolean
Dim tmptop As Integer
Dim tmplft As Integer
Public imenu As Integer
Private sRecoString As String
Dim pre(50) As Boolean


Private Sub aicAlphaImage1_DblClick()
On Error Resume Next
cpl "C:\WINDOWS\system32\timedate.cpl"
End Sub

Private Sub aicHour_DblClick()
On Error Resume Next
cpl "C:\WINDOWS\system32\timedate.cpl"
End Sub

Private Sub aicMenu_Click()
PopupMenu Form4.TechMenu, , aicMenu.Left, aicMenu.Top
End Sub

Private Sub aicMenu_MouseEnter()
aicMenu.FadeInOut 100
End Sub

Private Sub aicMenu_MouseExit()
aicMenu.FadeInOut 60
End Sub

Private Sub aicMinute_DblClick()
On Error Resume Next
cpl "C:\WINDOWS\system32\timedate.cpl"
End Sub

Private Sub aicSecond_DblClick()
On Error Resume Next
cpl "C:\WINDOWS\system32\timedate.cpl"
End Sub

Private Sub Bag_Click(Index As Integer)
On Error Resume Next
Dim inti As String
If Bag(Index).Tag = "Add New" Then
inti = InputFrm("Enter Address", "New Link", "www.com", "")
Make_Bag inti
ElseIf Bag(Index) = "Functions Calculator" Then
frmcalc.Show
Else
On Error Resume Next
exp Bag(Index).Tag
End If
End Sub
Sub Make_Cafe(inti As String)
On Error Resume Next
load Cafe(Cafe.UBound + 1)
Cafe(Cafe.UBound).Caption = GetFilename(inti)
Cafe(Cafe.UBound).Tag = inti
Cafe(Cafe.UBound).Visible = True
Setcont
End Sub
Sub Make_Bag(inti As String)
On Error Resume Next
load Bag(Bag.UBound + 1)
Bag(Bag.UBound).Caption = GetFilename(inti)
Bag(Bag.UBound).Tag = inti
Bag(Bag.UBound).Visible = True
Setcont
End Sub
Sub Make_Target(inti As String)
On Error Resume Next
load tasks(tasks.UBound + 1)
tasks(tasks.UBound).Tag = inti
tasks(tasks.UBound).Caption = inti
tasks(tasks.UBound).Visible = True
Setcont
End Sub


Private Sub Cafe_Click(Index As Integer)
On Error Resume Next
Dim inti As String
If Cafe(Index).Tag = "Add New" Then
inti = InputFrm("Enter Address", "New Link", "www.com", "")
Make_Cafe inti
Else
On Error Resume Next
ShellFile Cafe(Index).Tag
End If
End Sub

Private Sub clos_Click()
On Error Resume Next
Shut.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
fbout.Show
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF1 Then
Sknr.toggle
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim str As String
Set Form = Form1
mini
FormOnBottom Me
Bag(0).MousePointer = LoadPicture(App.path & "\reserved\harrow.cur")
Me.Top = 0
Me.Left = 0
Me.Width = Screen.Width
Me.Height = Screen.Height
clos.Top = (Me.Height / 15) - clos.Height
Dim cH As New cAppHider
cH.HideApplication
clockpos

tm = False
bott = True
Form3.already = False
load_menus
icon_menu
load_grid
frmcst.GetIni
Set_MP
TimeTable1.LoadTable App.path & "\TimeTable.table"
'If fso.FileExists(App.EXEName & ".exe.manifest") = False Then
'XPVB Enable this to have Xp style controls
'End If But will not run when compiled
tmranim8 = True
API.TaskBarHide
PrepareTag
tbar.Show
LoadDesktop
str = GetFromIni("Main", "Run", App.path & "\Config.ini")
If str <> "" Then
ShellFile str
End If
frmcst.GetIni
End Sub


Public Function LoadDesktop()
On Error Resume Next
Dim i As Long, n As Integer, Num As Integer, n2 As Integer
Dim pth As String
Dim FF As Long
Dim l As Integer
Dim P As Long
Dim path As String, icon As String, Marker As String, mak2(50) As Integer
Dim x As Long, Y As Long
Dim Caption As String
Dim pword As String
Dim z As Integer
pre(0) = True

If imgicon.UBound > 0 Then
    For i = 1 To imgicon.UBound
        Unload imgicon(i)
        Unload lblcaption(i)
    Next i
End If
DuCr.I2S App.path & "\Keys.bmp", App.path & "\Keys.ini"
pth = App.path & "\Links\"
file1.path = "C:"
file1.path = pth
ShapeNumber = 0
Num = 0: n2 = 0
file1.path = pth
For i = 0 To file1.ListCount - 1
    If Right(file1.List(i), 4) = ".lnk" Then
        load imgicon(imgicon.UBound + 1)
        load lblcaption(lblcaption.UBound + 1)

        pword = GetFromIni("Main", file1.List(i), App.path & "\Keys.ini")
        icon = GetFromIni("Main", "Picture", pth & file1.List(i))
        Marker = GetFromIni("Main", "Marker", pth & file1.List(i))
        Caption = GetFromIni("Main", "Caption", pth & file1.List(i))
        lblcaption(imgicon.UBound).Caption = Caption
        If Marker = "PreSet,PreSet" Then
            pre(i) = True
            mak2(n2) = imgicon.UBound
                n2 = n2 + 1
        Else
            pre(i) = False
            stng = InStr(1, Marker, ",")
            With imgicon(imgicon.UBound)
                .Top = Right(Marker, Len(Marker) - Val(stng))
                .Left = Left(Marker, Val(stng) - 1)
            End With
        End If
                    If Right$(icon, 10) = " <AppPath>" Then
                    icon = Left$(icon, Len(icon) - 10)
                    imgicon(imgicon.UBound).LoadImage_FromFile (App.path & "\icons\" & icon)
                    Else
                    imgicon(imgicon.UBound).LoadImage_FromFile (icon)
                    End If
                    Dim ONOff As Boolean
                    With imgicon(imgicon.UBound)
                       .ToolTipText = file1.List(i)
                        .Tag = pth & "\" & file1.List(i)
                    End With
                          With lblcaption(imgicon.UBound)
                               .Visible = False
                               .Top = imgicon(imgicon.UBound).Top + 20
                               .ZOrder 0
                              .Tag = pword
                End With
        DoEvents
        If pre((i)) = False Then
        With lblcaption(lblcaption.UBound) '- 1)
        .Caption = Caption
rewidth:
        .Left = imgicon(imgicon.UBound).Left + imgicon(imgicon.UBound).Width + 8
        lblcaption(imgicon.UBound).Width = Me.TextWidth(Caption) + 4
        End With
    End If
    End If
Next i
    For n = 0 To n2 - 1
        With imgicon(mak2(n))
                    .Top = (58 * n) + 52
            lblcaption(mak2(n)).Left = .Left + .Width + 4
            lblcaption(mak2(n)).Width = Me.TextWidth(lblcaption(mak2(n))) + 4
            lblcaption(mak2(n)).Top = imgicon(mak2(n)).Top + 20
            Order mak2(n), 1
        End With
    Next
        DoEvents
    Kill App.path & "\Keys.ini"
For i = 1 To imgicon.UBound
        imgicon(i).Visible = True
        lblcaption(i).Visible = True
Next
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
listsel.Clear
Set_down x, Y
Dim t As Integer
For t = 1 To shpsel.UBound + 1
Unload shpsel(t)
Next
shpsel(0).Visible = False
selected = False
MENU1.Visible = False
MENU2.Visible = False
Else
MENU1.Visible = True
MENU1.fd
MENU2.Visible = False
MENU1.Top = Y
MENU1.Left = x
If Y + MENU1.Height > Screen.Height Then
MENU1.Top = MENU1.Top - MENU1.Height
End If
If x + MENU1.Width > Screen.Width Then
MENU1.Left = MENU1.Left - MENU1.Width
End If
End If
End Sub

Sub LoadBkg(pth As String)
Dim c32 As New c32bppDIB
AutoRedraw = True
c32.InitializeDIB Screen.Width / 15, Screen.Height / 15
c32.LoadPicture_File pth
Set Picture = Nothing
c32.Render hdc, 0, 0, Screen.Width / 15, Screen.Height / 15
Picture = Image
Refresh
AutoRedraw = False
Set c32 = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
Call Set_move(x, Y)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim s As Integer
Set_up x, Y
listsel.Clear
For s = 1 To imgicon.UBound
If On_Over(s) = True Then
load shpsel(shpsel.UBound + 1)
shpsel(shpsel.UBound).Top = imgicon(s).Top
shpsel(shpsel.UBound).Left = imgicon(s).Left
shpsel(shpsel.UBound).Visible = True
listsel.Additem imgicon(s).Index
selected = True
End If
Next
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim strsave As String
    For i = 1 To Data.Files.count
strsave = App.path & "\links\" & WOExt(Data.Files(i)) & ".lnk"
Call WriteIni("Main", "Path", Data.Files(i), strsave)
Call WriteIni("Main", "Caption", WOExt(Data.Files(i)), strsave)
Call WriteIni("Main", "Marker", "PreSet,PreSet", strsave)
Call WriteIni("Main", "Key", "", strsave)
    If fso.FileExists(Data.Files(i)) = True Then
Call WriteIni("Main", "Picture", App.path & "\icons\Blue Image.png", strsave)
    Else
Call WriteIni("Main", "Picture", App.path & "\icons\Registers.png", strsave)
    End If
Next
LoadDesktop
End Sub

Function WOExt(str As String)
On Error Resume Next
Dim dpth As String
Dim init As Integer
Dim str2 As String
dpth = InStrRev(str, ".")
If dpth <> 0 Then dpth = dpth - Len(Text1)
str2 = Left(str, Len(str) + dpth)
init = Len(str2) - InStrRev(str2, "\")
str2 = Right(str2, init)
If Right(str2, 1) = "." Then str2 = Left(str2, Len(str2) - 1)
WOExt = str2
End Function

Private Sub Form_Terminate()
On Error Resume Next
API.TaskBarShow
tbar.Undock
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
API.TaskBarShow
tbar.Undock
End
End Sub

Private Sub grd_Timer()
On Error Resume Next
If s1(s1.UBound).Top > Screen.Height / 15.5 Then
grd = False
End If
load s2(s2.UBound + 1)
load s1(s1.UBound + 1)
s2(s2.UBound).Left = s2(s2.UBound - 1).Left + 58
s2(s2.UBound).Height = Me.Height / 15.5
s1(s1.UBound).Top = s1(s1.UBound - 1).Top + 58
s1(s1.UBound).Width = Me.Width / 15.5
End Sub

Private Sub imgicon_DblClick(Index As Integer)
On Error GoTo x
Dim Spth As String
Dim ipt As String
If lblcaption(Index).Tag = "" Then
MENU2.Visible = False
Spth = GetFromIni("Main", "Path", imgicon(Index).Tag)
ShellFile (Spth)
Else
MENU2.Visible = False
ipt = InputFrm("Enter Password", "Locked", "Enter Key", "*")
If ipt = lblcaption(Index).Tag Then
Spth = GetFromIni("Main", "Path", imgicon(Index).Tag)
ShellFile (Spth)
Else
MENU2.Visible = False
MsgBox "Invalid Key", vbCritical, "Error"
End If
End If
Exit Sub
x:
MsgBox "Invalid Shortcut : " & Spth, vbCritical + vbApplicationModal, "Error"
End Sub

Private Sub imgicon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

imgicon(Index).ZOrder 0
If Button = 1 Then
MENU1.Visible = False
mX = x
mY = Y
Exit Sub
ElseIf Button = 2 Then
MENU2.Top = imgicon(Index).Top + imgicon(Index).Height
MENU2.Left = imgicon(Index).Left
MENU2.fd
MENU2.Visible = True
MENU1.Visible = False
imenu = imgicon(Index).Index
End If
End Sub

Private Sub imgicon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
If selected = False Then
    imgicon(Index).Left = imgicon(Index).Left + x - mX
    imgicon(Index).Top = imgicon(Index).Top + Y - mY
    lblcaption(Index).Left = imgicon(Index).Left + imgicon(Index).Width + 8
    lblcaption(Index).Top = imgicon(Index).Top + 20
ElseIf selected = True Then
listsel.selected(0) = True
Dim q As Integer
Dim i As Integer
For q = 0 To listsel.ListCount - 1
listsel.selected(q) = True
If imgicon(Index).Index = listsel.text Then
For i = 0 To listsel.ListCount - 1
listsel.selected(i) = True
    imgicon(listsel.text).Left = imgicon(listsel.text).Left + x - mX
    imgicon(listsel.text).Top = imgicon(listsel.text).Top + Y - mY
    lblcaption(listsel.text).Left = imgicon(listsel.text).Left + imgicon(listsel.text).Width + 8
    lblcaption(listsel.text).Top = imgicon(listsel.text).Top + 20
On Error Resume Next
shpsel(i).Visible = False
Next
End If
Next
End If
End If
End Sub

Private Sub imgicon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
If Button = 1 Then
tmrMoveIcon.Enabled = False
    Order Index, Button
End If
Dim q As Integer
Dim i As Integer
For q = 0 To listsel.ListCount
listsel.ListIndex = q
shpsel(listsel.text).Visible = False
Order listsel.text, 1
shpsel(listsel.text).Top = imgicon(listsel.text).Top
shpsel(listsel.text).Left = imgicon(listsel.text).Left
Next
End Sub

Sub quit()
On Error Resume Next
SetCursorPos Screen.ActiveForm.Left / 15.5 + Screen.ActiveForm.Width / 15.5 - 5, Screen.ActiveForm.Top / 15.5 + 15
MouseClick "Left"
End Sub

Private Sub imgicon_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim str As String, z As Integer, sf As String
str = GetFromIni("Main", "Path", imgicon(Index).Tag)

If str = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
    MsgBox "Cannot copy file to My Computer", vbCritical, "Error"
ElseIf str = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}" Then
    MsgBox "Cannot copy file to Network places", vbCritical, "Error"
ElseIf str = "::{450D8FBA-AD25-11D0-98A8-0800361B1103}" Then
If MsgBox("Do you want to copy " & Data.Files.count & " file(s) to My Documents", vbYesNo, "Copy file") = vbYes Then
    For z = 1 To Data.Files.count
        On Error GoTo n
        If fso.FileExists(Data.Files(z)) = True Then
            fso.CopyFile Data.Files(z), "C:\Documents and Settings\Administrator\My Documents\"
        ElseIf fso.FolderExists(Data.Files(z)) = True Then
            fso.CopyFolder Data.Files(z), "C:\Documents and Settings\Administrator\My Documents\"
        Else
            MsgBox "Error while accessing :  " & Data.Files(z), vbCritical, "Error"
        End If
       GoTo P

n:
        MsgBox "Error while accessing :  " & Data.Files(z), vbCritical, "Error"
P:
    Next
End If
ElseIf str = "Reserved\Trash.lnk" Then
If MsgBox("Do you want to delete " & Data.Files.count & " file(s)", vbYesNo, "Move to recyclebin") = vbYes Then
    For z = 1 To Data.Files.count
        On Error GoTo M
    Dim Oper As SHFILEOPSTRUCT
        With Oper
            .wFunc = &H3
            .pFrom = Data.Files(z)
            .fFlags = &H40
        End With
            SHFileOperation Oper
        GoTo o
M:
        MsgBox "Error while accessing :  " & Data.Files(z), vbCritical, "Error"
o:
    Next
    End If
ElseIf fso.FileExists(str) = True Then
    MsgBox "Cannot copy " & Data.Files.count & " file(s) to a file", vbCritical, "Error"
Else
If MsgBox("Do you want to copy " & Data.Files.count & " file(s) to " & GetFilename(str), vbYesNo, "Copy file") = vbYes Then
        On Error GoTo t
    For z = 1 To Data.Files.count
        If fso.FileExists(Data.Files(z)) = True Then
            fso.CopyFile Data.Files(z), str
        ElseIf fso.FolderExists(Data.Files(z)) = True Then
            fso.CopyFolder Data.Files(z), str
        Else
            MsgBox "Error while accessing :  " & Data.Files(z), vbCritical, "Error"
        End If
        GoTo i
t:
        MsgBox "Error while accessing :  " & Data.Files(z), vbCritical, "Error"
        GoTo d
i:
    Next
End If
End If
d:
End Sub

Private Sub lblPer_Click()
Dim x As Integer
Dim curt As Integer
Dim ext As Integer
curt = fixtime(Hour(Time)) & fixtime(Minute(Time))
For x = 0 To 7
If curt > frmTT.Dm(x) And curt < frmTT.Rg(x) Then
ext = x + 1
End If
Next
MsgBox ext & "\\" & curt & "\\" & frmTT.Dm(ext) & "\\" & frmTT.Rg(ext)
End Sub

Function fixtime(tyme As Integer)
If CStr(Len(tyme)) = 1 Then
fixtime = CInt("0" & CStr(tyme))
Else
fixtime = tyme
End If
End Function

Private Sub Timer2_Timer()
On Error Resume Next
    Dim tTime As Date
    tTime = Time
    If Second(tTime) = 0 Then
        aicHour.Rotation() = 30 * Hour(tTime) + (Minute(tTime) / 60) * 24
        aicMinute.Rotation() = 6 * Minute(tTime)
    End If
    aicSecond.Rotation() = 6 * Second(tTime)
Dim pint As POINTAPI
GetCursorPos pint
If WindowFromPoint(pint.x, pint.Y) = TimeTable1.handle Then
TimeTable1.Left = (Screen.Width / 15) - TimeTable1.Width
Else
TimeTable1.Left = (Screen.Width / 15) - TimeTable1.Width + 400
End If
Dim x As Integer
Dim curt As Integer
Dim ext As Integer
curt = fixtime(Hour(Time)) & fixtime(Minute(Time))
For x = 0 To 7
If curt > frmTT.Dm(x) And curt < frmTT.Rg(x) Then
ext = x + 1
Exit For
End If
Next
lblPer = ext
End Sub


Private Sub menu1_Click(str As String)
On Error Resume Next
MENU1.Visible = False
If str = "Refresh" Then
LoadDesktop
ElseIf str = "Create New Icon" Then
Form2.Show
Form2.titlebar.reload
ElseIf str = "Customize" Then
frmcst.Show
frmcst.titlebar.reload
ElseIf str = "Properties Page" Then
frmcln.Show
frmcln.Form_Load
frmcln.titlebar.reload
ElseIf str = "Desktop Clean Wizard" Then
Form5.Show
Form5.titlebar.reload
End If
End Sub
Public Sub load_menus()
On Error Resume Next
MENU1.Additem "Refresh", Form4.Image8
MENU1.Additem "Create New Icon", Form4.Image5
MENU1.Additem "Properties Page", Form4.Image7
MENU1.Additem "Customize", Form4.image1
MENU1.Additem "Desktop Clean Wizard", Form4.Image2
End Sub
Public Sub icon_menu()
On Error Resume Next
MENU2.Additem "Execute", Form4.Image4
MENU2.Additem "Rename", Form4.Image3
MENU2.Additem "Delete", Form4.Image9
MENU2.Additem "Set Position", Form4.Image11
MENU2.Additem "Lock", Form4.Image10
MENU2.Additem "Line-Up", Form4.Image8
End Sub

Private Sub menu2_Click(str As String)
On Error Resume Next
Dim intinput As String
Dim ipt As String
If MENU2.MItem = "Execute" Then
imgicon_DblClick imenu
ElseIf MENU2.MItem = "Rename" Then
MENU2.Visible = False
intinput = InputFrm(lblcaption(imenu).Caption, "Rename", lblcaption(imenu).Caption, "")
WriteIni "Main", "Caption", intinput, App.path & "\links\" & imgicon(imenu).ToolTipText
LoadDesktop
ElseIf MENU2.MItem = "Delete" Then
If lblcaption(imenu).Tag = "" Then
MENU2.Visible = False
Formdel.Show
Formdel.Tag = imenu
Else
MENU2.Visible = False
ipt = InputFrm("Enter Password", "Locked", "Enter Key", "")
If ipt = lblcaption(imenu).Tag Then
MENU2.Visible = False
Formdel.Show
Formdel.Tag = imenu
Else
MENU2.Visible = False
MsgBox "Invalid Key", vbCritical, "Error"
End If
End If
ElseIf MENU2.MItem = "Lock" Then
If lblcaption(imenu).Tag = "" Then
frmlk.Locked = False
Else
frmlk.Password = lblcaption(imenu).Tag
frmlk.Locked = True
End If
frmlk.Show
frmlk.Tag = imenu
ElseIf MENU2.MItem = "Line-Up" Then
LoadDesktop
ElseIf MENU2.MItem = "Set Position" Then
Call WriteIni("Main", "Marker", imgicon(imenu).Left & "," & imgicon(imenu).Top, App.path & "\links\" & imgicon(imenu).ToolTipText)
End If
MENU2.Visible = False
End Sub
Public Sub clockpos()
On Error Resume Next
aicAlphaImage1.Width = 137
aicAlphaImage1.Height = 137
aicHour.Width = 137
aicHour.Height = 137
aicMinute.Height = 137
aicMinute.Width = 137
aicSecond.Width = 137
aicSecond.Height = 137
aicAlphaImage1.Left = Me.Width / 15 - aicAlphaImage1.Width
aicAlphaImage1.Top = Me.Height / 15 - aicAlphaImage1.Height
aicHour.Left = Me.Width / 15 - aicHour.Width
aicMinute.Left = Me.Width / 15 - aicMinute.Width
aicSecond.Left = Me.Width / 15 - aicSecond.Width
aicHour.Top = Me.Height / 15 - aicHour.Height
aicMinute.Top = Me.Height / 15 - aicMinute.Height
aicSecond.Top = Me.Height / 15 - aicSecond.Height
aicAlphaImage1.Visible = True
aicHour.Visible = True
aicMinute.Visible = True
aicSecond.Visible = True
aicAlphaImage1.ZOrder 1
End Sub


Public Sub load_grid()
On Error Resume Next
s2(0).Height = Me.Height / 15.5
s1(0).Width = Me.Width / 15.5
grd = True
End Sub

Public Sub Order(Index As Integer, Button As Integer)
On Error Resume Next
Dim x As Integer
If Button = 1 Then
For x = 0 To s1.UBound
If imgicon(Index).Left + 24 > s2(x).Left Then
If imgicon(Index).Left + 24 < s2(x + 1).Left Then
imgicon(Index).Left = s2(x).Left
lblcaption(Index).Left = imgicon(Index).Left + imgicon(Index).Width + 8
End If
End If
If imgicon(Index).Top + 24 > s1(x).Top Then
If imgicon(Index).Top + 24 < s1(x + 1).Top Then
imgicon(Index).Top = s1(x).Top
lblcaption(Index).Top = imgicon(Index).Top + 20
End If
End If
Next
End If
End Sub

Private Sub Set_MP()
On Error Resume Next
MP1.Left = Me.Width / 15 - 137 - 278
MP1.Top = Me.Height / 15 - 185
TimeTable1.Left = (Screen.Width / 15) - TimeTable1.Width
TimeTable1.Top = MP1.Top - TimeTable1.Height - 5
        aicHour.Rotation() = 30 * Hour(tTime) + (Minute(tTime) / 60) * 24
        aicMinute.Rotation() = 6 * Minute(tTime)
        aicMenu.LoadImage_FromFile App.path & "\Images\Menu.png"
        aicMenu.AutoSize = True
    aicMenu.Top = TimeTable1.Top - aicMenu.Height - 2
    aicMenu.Left = (Screen.Width / 15) - aicMenu.Width - 2
    lblPer.Top = aicMenu.Top
    lblPer.Left = aicMenu.Left - lblPer.Width - 2
frmTT.load
End Sub
Private Sub MP1_DClick()
On Error Resume Next
If MP1.fullscreen = False Then
Form6.Show
End If
End Sub

Private Sub MP1_Forward()
On Error GoTo x
Form6.Album.ListIndex = Form6.Album.ListIndex + 1
Form6.Lstdir.ListIndex = Form6.Album.ListIndex
MP1.Play Form6.Lstdir.text
Exit Sub
x:
On Error Resume Next
Form6.Album.ListIndex = 0
Form6.Lstdir.ListIndex = Form6.Album.ListIndex
MP1.Play Form6.Lstdir.text
End Sub

Private Sub MP1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
MP1.Play Data.Files(1)
End Sub

Private Sub MP1_Previous()
On Error GoTo x
If Form6.Album.ListIndex = 0 Then GoTo x
Form6.Album.ListIndex = Form6.Album.ListIndex - 1
Form6.Lstdir.ListIndex = Form6.Album.ListIndex
MP1.Play Form6.Lstdir.text
Exit Sub
x:
Form6.Album.ListIndex = Form6.Album.ListCount - 1
Form6.Lstdir.ListIndex = Form6.Album.ListIndex
MP1.Play Form6.Lstdir.text
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

Public Sub mini()
On Error Resume Next
Dim shl32 As New shell
shl32.MinimizeAll
End Sub
Public Sub Set_down(x As Single, Y As Single)
On Error Resume Next
Line1.X1 = 0
Line2.X1 = 0
Line3.X1 = 0
Line4.X1 = 0
Line1.Y1 = 0
Line2.Y1 = 0
Line3.Y1 = 0
Line4.Y1 = 0
Line1.X2 = 0
Line2.X2 = 0
Line3.X2 = 0
Line4.X2 = 0
Line1.Y2 = 0
Line2.Y2 = 0
Line3.Y2 = 0
Line4.Y2 = 0
Line1.X1 = x
Line1.Y1 = Y
Line1.Y2 = Y
Line2.Y1 = Y
Line2.X1 = x
Line2.X2 = x
Line4.X1 = x
Line3.Y1 = Y
Line1.X2 = x
Line2.Y2 = Y
Line4.X2 = x
Line4.Y2 = Y
Line4.Y1 = Y
Line3.Y2 = Y
Line3.X2 = x
Line3.X1 = x
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
End Sub
Public Sub Set_move(x As Single, Y As Single)
On Error Resume Next
Line1.X2 = x
Line2.Y2 = Y
Line4.X2 = x
Line4.Y2 = Y
Line4.Y1 = Y
Line3.Y2 = Y
Line3.X2 = x
Line3.X1 = x
End Sub
Public Sub Set_up(x As Single, Y As Single)
On Error Resume Next
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
End Sub
Public Function On_Over(idx As Integer) As Boolean
On Error Resume Next
If Line4.Y1 >= imgicon(idx).Top Then
If Line3.X1 >= imgicon(idx).Left Then
If Line2.X1 <= imgicon(idx).Left + imgicon(idx).Width Then
If Line1.Y1 <= imgicon(idx).Top + imgicon(idx).Height Then
On_Over = True
GoTo P
End If
End If
End If
End If
If Line1.Y1 >= imgicon(idx).Top Then
If Line2.X1 >= imgicon(idx).Left Then
If Line3.X1 <= imgicon(idx).Left + imgicon(idx).Width Then
If Line4.Y1 <= imgicon(idx).Top + imgicon(idx).Height Then
On_Over = True
GoTo P
End If
End If
End If
End If
If Line1.Y1 <= imgicon(idx).Top + imgicon(idx).Height Then
If Line3.X1 <= imgicon(idx).Left + imgicon(idx).Width Then
If Line2.X1 >= imgicon(idx).Left Then
If Line4.Y1 >= imgicon(idx).Top Then
On_Over = True
GoTo P
End If
End If
End If
End If
If Line4.Y1 <= imgicon(idx).Top + imgicon(idx).Height Then
If Line3.X1 >= imgicon(idx).Left + imgicon(idx).Width Then
If Line1.Y1 >= imgicon(idx).Top Then
If Line2.X1 <= imgicon(idx).Left Then
On_Over = True
GoTo P
End If
End If
End If
End If
On_Over = False
P:
End Function

Sub Setcont()
On Error Resume Next
Dim str As String
Dim x
tasks(0).Left = Width / 15 - 129 - 20
tasks(0).Top = 30 + 27
For x = 1 To tasks.UBound
tasks(x).Left = Width / 15 - 129 - 20
tasks(x).Top = tasks(x - 1).Top + 17
Next
Dim Y As Integer
Cafe(0).Left = Width / 15 - 129 - 20 - 159 - 25
Cafe(0).Top = 30 + 27
For Y = 1 To Cafe.UBound
Cafe(Y).Left = Width / 15 - 129 - 20 - 159 - 25
Cafe(Y).Top = Cafe(Y - 1).Top + 17
Next
Dim z As Integer
Bag(0).Left = Width / 15 - 129 - 20 - 318 - 50 + 20
Bag(0).Top = 30 + 27
For Y = 1 To Bag.UBound
Bag(Y).Left = Width / 15 - 129 - 20 - 318 - 50 + 20
Bag(Y).Top = Bag(Y - 1).Top + 17
Next
End Sub

Private Sub tasks_Click(Index As Integer)
On Error Resume Next
Dim inti As String
If tasks(Index).Tag = "Add New" Then
inti = InputFrm("Enter Address", "New Link", "www.com", "")
Make_Target inti
Else
On Error Resume Next
exp tasks(Index).Caption
End If
End Sub

Sub PrepareTag()
On Error Resume Next
Make_Bag App.path & "\reserved\My Calculator.lnk"
Make_Bag "Functions Calculator"
Make_Bag App.path & "\reserved\My Notes.lnk"
Make_Bag App.path & "\reserved\My Browser.lnk"
Make_Bag App.path & "\reserved\My Register.lnk"
Make_Cafe App.path & "\reserved\TALK IT.lnk"
Make_Cafe App.path & "\reserved\PLAY IT.lnk"
Make_Cafe App.path & "\reserved\Windows Media Player.lnk"
Make_Target "www.wikipedia.en"
Make_Target "www.yahoo.com"
Make_Target "www.w3schools.com"
Make_Target "www.google.com"
Make_Target "www.sciencebuddies.com"
Make_Target "www.orkut.com"
Make_Target "www.mail.yahoo.com"
Make_Target "www.mail.google.com"
Make_Target "www.hotmail.com"
End Sub


Private Sub TimeTable1_PropHit(x As Integer, Y As Integer)
PopupMenu Form4.TTprop, , TimeTable1.Left + (x / 15), TimeTable1.Top + (Y / 15)
End Sub

Private Sub tmranim8_Timer()
On Error Resume Next
Setcont
tmranim8 = False
End Sub


