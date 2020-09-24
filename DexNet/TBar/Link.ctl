VERSION 5.00
Begin VB.UserControl UserControl2 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Sub Show()
On Error Resume Next
frmtaskbar.Show
End Sub

Sub endit()
On Error Resume Next
UnDockForm frmtaskbar
Unload frmtaskbar
Unload Frmhid
Unload frmico
Unload frmmenu
Unload frmpop
Unload frmtip1
Unload Frmtip2
Unload Frmulti
End Sub

Sub skin(dir As String)
On Error Resume Next
On Error Resume Next
Set frmtaskbar.Image1 = LoadPicture(dir & "\bar.bmp")
frmtaskbar.PictureButton1.loadpics dir, "\Start"
Set frmtaskbar.Picture1 = LoadPicture(dir & "\tbarBlu.bmp")
Set frmtaskbar.imgico = LoadPicture(dir & "\tbariconHL.bmp")
frmtaskbar.Clear
Set frmtaskbar.Picture4(0) = LoadPicture(dir & "\tbariconBK.bmp")
frmtaskbar.InitButtons
Set frmtip1.img = LoadPicture(dir & "\str.bmp")
End Sub

Sub TNrml()
On Error Resume Next
frmtaskbar.Timer3 = False
End Sub

Sub Ttop()
On Error Resume Next
frmtaskbar.Timer3 = True
End Sub

Sub Undock()
UnDockForm frmtaskbar
End Sub

