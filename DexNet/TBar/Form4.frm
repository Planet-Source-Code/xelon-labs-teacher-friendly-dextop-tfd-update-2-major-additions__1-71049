VERSION 5.00
Begin VB.Form Frmulti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple Tasks"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Multiple Normalize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Multiple Minimize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Multiple Maximize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Make Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Multiple On top"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Multiple Glass"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Multiple Hide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   810
         Left            =   720
         MultiSelect     =   2  'Extended
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Frmulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
fade List2.list(x), -50
AppHide List2.list(x)
load frmpop.hides(frmpop.hides.UBound + 1)
frmpop.hides(frmpop.hides.UBound).Tag = List2.list(x)
frmpop.hides(frmpop.hides.UBound).Caption = List1.list(x)
End If
Next
Frmhid.Form_Load
Form_Load
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim x As Integer
On Error Resume Next
text1 = InputBox("Use values from 0 to 255" & vbCrLf & " for giving these windows a glass effect", "Make Glass", "Opaque")
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
If text1 = "Opaque" Then
MakeOpaque List2.list(x)
MakeTransparent List2.list(x), text1
End If
frmtaskbar.lsttrans.list(x) = text1
End If
Next
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
FormOnTop List2.list(x)
End If
Next
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
MakeNormal List2.list(x)
End If
Next
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
maximize List2.list(x)
End If
Next
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
minimize List2.list(x)
End If
Next
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
ActivateWindow List2.list(x)
End If
Next
End Sub

Public Sub Form_Load()
On Error Resume Next
List1.Clear
List2.Clear
Dim x As Integer
For x = 1 To frmtaskbar.Picture4.UBound
List1.AddItem frmtaskbar.Picture4(x).ToolTipText
List2.AddItem frmtaskbar.Picture4(x).Tag
Next
End Sub
