VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmpop 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   960
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "General Transparent Value of a window"
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu trans 
      Caption         =   "Trans"
      Begin VB.Menu switch 
         Caption         =   "Switch to "
      End
      Begin VB.Menu close 
         Caption         =   "Close"
      End
      Begin VB.Menu order 
         Caption         =   "Set Order"
         Begin VB.Menu normalizator 
            Caption         =   "Normalize"
         End
         Begin VB.Menu maxi 
            Caption         =   "Maximize"
         End
         Begin VB.Menu mini 
            Caption         =   "Minimize"
         End
         Begin VB.Menu Hider 
            Caption         =   "Hide"
         End
      End
      Begin VB.Menu level 
         Caption         =   "Set Level"
         Begin VB.Menu Shadow 
            Caption         =   "Make Always On Top"
         End
         Begin VB.Menu normal 
            Caption         =   "Make Normal"
         End
      End
      Begin VB.Menu options 
         Caption         =   "Glass Options"
         Begin VB.Menu opq 
            Caption         =   "Make Opaque"
         End
         Begin VB.Menu glass 
            Caption         =   "Custom Glass Effect"
         End
         Begin VB.Menu t200 
            Caption         =   "Make Transparent 200"
         End
         Begin VB.Menu t150 
            Caption         =   "Make Transparent 150"
         End
         Begin VB.Menu t100 
            Caption         =   "Make Transparent 100"
         End
         Begin VB.Menu t50 
            Caption         =   "Make Transparent 50"
         End
         Begin VB.Menu t25 
            Caption         =   "Make Transparent 25"
         End
      End
   End
   Begin VB.Menu hideo 
      Caption         =   "Hideo"
      Begin VB.Menu miniall 
         Caption         =   "Minimize All"
      End
      Begin VB.Menu cascade 
         Caption         =   "Cascade All"
      End
      Begin VB.Menu tver 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu tHori 
         Caption         =   "Tile Horizintally"
      End
      Begin VB.Menu mprf 
         Caption         =   "Multiple Perform"
      End
      Begin VB.Menu HP 
         Caption         =   "Hidden Programs"
         Begin VB.Menu hides 
            Caption         =   "Multiple Show"
            Index           =   0
         End
      End
   End
   Begin VB.Menu MShower 
      Caption         =   "Shower"
      Begin VB.Menu Showt 
         Caption         =   "Show with transition"
      End
      Begin VB.Menu Showwt 
         Caption         =   "Show without transition"
      End
   End
   Begin VB.Menu MHider 
      Caption         =   "Hider"
      Begin VB.Menu Hidet 
         Caption         =   "Hide with transition"
      End
      Begin VB.Menu hidewt 
         Caption         =   "Hide without transition"
      End
   End
   Begin VB.Menu Ser 
      Caption         =   "Search"
      Begin VB.Menu sr 
         Caption         =   "Search Google"
         Index           =   0
      End
      Begin VB.Menu sr 
         Caption         =   "Search Planet Source Code"
         Index           =   1
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "menu"
      Begin VB.Menu exec 
         Caption         =   "Execute"
      End
      Begin VB.Menu Cad 
         Caption         =   "Copy Address"
      End
      Begin VB.Menu Ren 
         Caption         =   "Remove Entry"
      End
      Begin VB.Menu slst 
         Caption         =   "Save list "
      End
      Begin VB.Menu llst 
         Caption         =   "Load list"
      End
   End
   Begin VB.Menu AProg 
      Caption         =   "All Programs"
      Begin VB.Menu ref 
         Caption         =   "Refresh"
      End
      Begin VB.Menu ls2 
         Caption         =   "-"
      End
      Begin VB.Menu exp 
         Caption         =   "Explore"
      End
   End
End
Attribute VB_Name = "frmpop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cad_Click()
On Error Resume Next
Clipboard.SetText frmmenu.List1.SelectedItem.text
End Sub

Private Sub cascade_Click()
On Error Resume Next
Dim shll As New Shell
shll.CascadeWindows
End Sub

Private Sub close_Click()
On Error Resume Next
SendMessage frmtaskbar.Picture4(frmtaskbar.imenu).Tag, &H10, 0, 0

End Sub

Private Sub exec_Click()
On Error Resume Next
ShellFile frmmenu.sel
End Sub

Private Sub exp_Click()
ShellFile "C:\Documents and Settings\All Users\Start Menu\Programs\"
End Sub

Private Sub glass_Click()
On Error GoTo x
Dim intinput As Integer
text1 = InputBox("Use values from 0 to 255" & vbCrLf & " for giving " & frmtaskbar.Picture4(frmtaskbar.imenu).ToolTipText & " a glass effect", "Make Glass", "Opaque")
MakeTransparent frmtaskbar.Picture4(frmtaskbar.imenu).Tag, text1
frmtaskbar.lsttrans.list(frmtaskbar.imenu) = text1
Exit Sub
x:
MakeOpaque frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub Hider_Click()
On Error Resume Next
If MsgBox("Warning!!, Hiding " & frmtaskbar.Picture4(frmtaskbar.imenu).ToolTipText & " will send it to" & vbCrLf & "Hider menu and it will run in background", vbInformation + vbOKCancel, "Hide") = vbOK Then
text1 = frmtaskbar.lsttrans.list(frmtaskbar.imenu - 1)
fade frmtaskbar.Picture4(frmtaskbar.imenu).Tag, -50
AppHide frmtaskbar.Picture4(frmtaskbar.imenu).Tag
load hides(hides.UBound + 1)
hides(hides.UBound).Tag = frmtaskbar.Picture4(frmtaskbar.imenu).Tag
hides(hides.UBound).Caption = frmtaskbar.Picture4(frmtaskbar.imenu).ToolTipText
End If
End Sub

Private Sub hides_Click(Index As Integer)
On Error Resume Next
If hides(Index).Caption = "Multiple Show" Then GoTo x
ActivateWindow hides(Index).Tag
fade hides(Index).Tag, 50
MakeOpaque hides(Index).Tag
hides(Index).Tag = ""
arrange
Frmhid.Form_Load
Exit Sub
x:
Frmhid.Show
Frmhid.Form_Load

End Sub

Private Sub llst_Click()
On Error Resume Next
cdlg.ShowOpen
Get_INIList cdlg.filename, frmmenu.List1
End Sub

Private Sub maxi_Click()
On Error Resume Next
maximize frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub mini_Click()
On Error Resume Next
minimize frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub miniall_Click()
On Error Resume Next
Dim shll As New Shell
shll.MinimizeAll
End Sub

Private Sub mprf_Click()
On Error Resume Next
Frmulti.Show
End Sub

Private Sub msho_Click()
On Error Resume Next
Frmhid.Show
End Sub

Private Sub normal_Click()
On Error Resume Next
MakeNormal frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub normalizator_Click()
On Error Resume Next
Normalize frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub opq_Click()
On Error Resume Next
MakeOpaque frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub sg_Click()

End Sub

Private Sub ref_Click()
frmmenu.StartMenu1.LoadItems
End Sub

Private Sub Ren_Click()
On Error Resume Next
frmmenu.List1.ListItems.Remove frmmenu.sel
End Sub

Private Sub sdl_Click(Index As Integer)

End Sub

Private Sub Run_Click()
ShellFile frmmenu.StartMenu1.File
End Sub

Private Sub Shadow_Click()
On Error Resume Next
FormOnTop frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub slst_Click()
On Error Resume Next
cdlg.ShowSave
Write_INIList cdlg.filename, frmmenu.List1
End Sub

Private Sub sr_Click(Index As Integer)
On Error Resume Next
Dim str As String
str = frmtaskbar.text1.text
For x = 0 To Len(frmtaskbar.text1.text)
frmtaskbar.text1.Find " "
If frmtaskbar.text1.SelText <> "" Then
frmtaskbar.text1.SelText = "+"
End If
Next
If Index = 0 Then
ShellFile "http://www.google.com/search?client=opera&rls=en&q=" & frmtaskbar.text1.text & "&sourceid=opera&ie=utf-8&oe=utf-8"
ElseIf Index = 1 Then
ShellFile "http://planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&txtCriteria=" & frmtaskbar.text1.text & "&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&lngWId=1&B1=Quick+Search"
End If
frmtaskbar.text1.text = str
End Sub

Private Sub switch_Click()
On Error Resume Next
ActivateWindow frmtaskbar.Picture4(frmtaskbar.imenu).Tag
End Sub

Private Sub t100_Click()
On Error Resume Next
MakeTransparent frmtaskbar.Picture4(frmtaskbar.imenu).Tag, 100
frmtaskbar.lsttrans.list(frmtaskbar.imenu) = "100"
End Sub

Private Sub t150_Click()
On Error Resume Next
MakeTransparent frmtaskbar.Picture4(frmtaskbar.imenu).Tag, 150
frmtaskbar.lsttrans.list(frmtaskbar.imenu) = "150"
End Sub

Private Sub t200_Click()
On Error Resume Next
MakeTransparent frmtaskbar.Picture4(frmtaskbar.imenu).Tag, 200
frmtaskbar.lsttrans.list(frmtaskbar.imenu) = "200"
End Sub

Private Sub t25_Click()
On Error Resume Next
MakeTransparent frmtaskbar.Picture4(frmtaskbar.imenu).Tag, 25
frmtaskbar.lsttrans.list(frmtaskbar.imenu) = "25"
End Sub

Private Sub t50_Click()
On Error Resume Next
MakeTransparent frmtaskbar.Picture4(frmtaskbar.imenu).Tag, 50
frmtaskbar.lsttrans.list(frmtaskbar.imenu) = "50"
End Sub

Private Sub tHori_Click()
On Error Resume Next
Dim shll As New Shell
shll.TileHorizontally
End Sub

Private Sub tver_Click()
On Error Resume Next
Dim shll As New Shell
shll.TileVertically
End Sub

Public Sub arrange()
On Error Resume Next
Dim itemX(25) As String, itemY(25) As Long, x As Integer, n As Integer
n = 0
If hides.UBound > 0 Then
For x = 1 To hides.UBound
If hides(x).Tag <> "" Then
itemX(n) = hides(x).Caption
itemY(n) = hides(x).Tag
n = n + 1
End If
Next
For x = hides.UBound To 1 Step -1
Unload hides(x)
Next
For x = 0 To n - 1
load hides(hides.UBound + 1)
hides(hides.UBound).Caption = itemX(x)
hides(hides.UBound).Tag = itemY(x)
Next
End If
End Sub
