VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image11 
      Height          =   360
      Left            =   600
      Picture         =   "Form4.frx":0000
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image Image10 
      Height          =   360
      Left            =   600
      Picture         =   "Form4.frx":0704
      Top             =   1560
      Width           =   360
   End
   Begin VB.Image Image9 
      Height          =   360
      Left            =   600
      Picture         =   "Form4.frx":0E08
      Top             =   1080
      Width           =   360
   End
   Begin VB.Image Image8 
      Height          =   360
      Left            =   600
      Picture         =   "Form4.frx":150C
      Top             =   600
      Width           =   360
   End
   Begin VB.Image Image7 
      Height          =   360
      Left            =   600
      Picture         =   "Form4.frx":1C10
      Top             =   120
      Width           =   360
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   120
      Picture         =   "Form4.frx":2314
      Top             =   2520
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   360
      Left            =   120
      Picture         =   "Form4.frx":2A18
      Top             =   2040
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   120
      Picture         =   "Form4.frx":311C
      Top             =   1560
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   120
      Picture         =   "Form4.frx":3820
      Top             =   1080
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   120
      Picture         =   "Form4.frx":3F24
      Top             =   600
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "Form4.frx":4628
      Top             =   120
      Width           =   360
   End
   Begin VB.Menu RT 
      Caption         =   "Remove Tool"
      Begin VB.Menu remc 
         Caption         =   "Remove Shortcut"
      End
      Begin VB.Menu rlo 
         Caption         =   "Remove Linked Object"
      End
   End
   Begin VB.Menu Lview 
      Caption         =   "List View"
      Begin VB.Menu iLarge 
         Caption         =   "Large Icons"
      End
      Begin VB.Menu ismall 
         Caption         =   "Small Icons"
      End
      Begin VB.Menu iReport 
         Caption         =   "Report View"
      End
      Begin VB.Menu iList 
         Caption         =   "List View"
      End
   End
   Begin VB.Menu TTprop 
      Caption         =   "TTable"
      Begin VB.Menu RTT 
         Caption         =   "Reset Time Table"
      End
      Begin VB.Menu STT 
         Caption         =   "Save Time Table"
      End
      Begin VB.Menu etmng 
         Caption         =   "Edit Timings"
      End
   End
   Begin VB.Menu TechMenu 
      Caption         =   "techMenu"
      Begin VB.Menu Att 
         Caption         =   "Attendence"
      End
      Begin VB.Menu Res 
         Caption         =   "Resource"
      End
      Begin VB.Menu etmng1 
         Caption         =   "Edit Timings"
      End
      Begin VB.Menu Snp 
         Caption         =   "Snap Shot"
      End
      Begin VB.Menu Diary 
         Caption         =   "Diary"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Att_Click()
frmSAS.Show
frmSAS.titleBar.blink
End Sub

Private Sub Diary_Click()
ShellFile "Notepad.exe"
End Sub

Private Sub ETmng_Click()
frmTT.Show
frmTT.titleBar.blink
End Sub

Private Sub etmng1_Click()
frmTT.Show
frmTT.titleBar.blink
End Sub

Private Sub iLarge_Click()
On Error Resume Next
Form6.file1.List_Size True
End Sub

Private Sub iList_Click()
On Error Resume Next
Form6.file1.View False
End Sub

Private Sub iReport_Click()
On Error Resume Next
Form6.file1.View True
End Sub

Private Sub ismall_Click()
On Error Resume Next
Form6.file1.List_Size False

End Sub

Private Sub remc_Click()
On Error Resume Next
Formdel.Delete_Click
End Sub

Private Sub Res_Click()
ShellFile App.path & "\Resource"
End Sub

Private Sub rlo_Click()
On Error Resume Next
Formdel.Delete2_Click
Formdel.Delete_Click
End Sub

Private Sub RTT_Click()
Form1.TimeTable1.LoadTable App.path & "\TimeTable.table"
End Sub

Private Sub Snp_Click()
Randomize
snapshot App.path & "\Resource\SnpSht " & CStr(Rnd) & ".jpg"
End Sub

Private Sub STT_Click()
Form1.TimeTable1.SaveTable App.path & "\TimeTable.table"
End Sub
