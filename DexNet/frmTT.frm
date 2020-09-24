VERSION 5.00
Begin VB.Form frmTT 
   BackColor       =   &H00D6AEA7&
   BorderStyle     =   0  'None
   Caption         =   "Edit Timings"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form7"
   ScaleHeight     =   4935
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0059341C&
      Caption         =   "Periods Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      Begin Project1.MacButton MacButton1 
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   4080
         Width           =   1215
         _extentx        =   2143
         _extenty        =   450
         btype           =   4
         tx              =   "Save"
         enab            =   -1
         font            =   "frmTT.frx":0000
         coltype         =   2
         focusr          =   -1
         bcol            =   15128530
         fcol            =   0
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   7
         Left            =   1920
         TabIndex        =   26
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   7
         Left            =   1200
         TabIndex        =   25
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   23
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   6
         Left            =   1200
         TabIndex        =   22
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   20
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   19
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   4
         Left            =   1920
         TabIndex        =   17
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   4
         Left            =   1200
         TabIndex        =   16
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   3
         Left            =   1200
         TabIndex        =   13
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   11
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   1200
         TabIndex        =   10
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Rg 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Dm 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 8"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   24
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 7"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   21
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 6"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   18
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 5"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 4"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Period 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time of Periods are given below, You can also change these,"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin Project1.title titlebar 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _extentx        =   5318
      _extenty        =   450
   End
End
Attribute VB_Name = "frmTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub load()
Dim i As Integer
Dim pth As String
pth = App.path & "\TimeTable.table"
For i = 0 To 7
Dm(i) = GetFromIni("Time", "Domain " & CStr(i), pth)
Rg(i) = GetFromIni("Time", "Range " & CStr(i), pth)
DoEvents
Next
End Sub

Sub Save()
Dim i As Integer
Dim pth As String
pth = App.path & "\TimeTable.table"
For i = 0 To 7
WriteIni "Time", "Domain " & CStr(i), Dm(i), pth
WriteIni "Time", "Range " & CStr(i), Rg(i), pth
Next
End Sub

Private Sub Form_Load()
titlebar.sett Me
load
End Sub

Private Sub MacButton1_Click()
Save
End Sub
