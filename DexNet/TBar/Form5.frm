VERSION 5.00
Begin VB.Form Frmhid 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple Shower"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Hidden Windows"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "Multiple Show"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Frmhid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim x As Integer
For x = 0 To List1.ListCount - 1
If List1.Selected(x) = True Then
ActivateWindow List2.list(x)
fade List2.list(x), 50
MakeOpaque List2.list(x)
frmpop.hides(List3.list(x)).Tag = ""
End If
Next
frmpop.arrange
Form_Load
End Sub

Public Sub Form_Load()
On Error Resume Next
List1.Clear
List2.Clear
List3.Clear
For x = 1 To frmpop.hides.UBound
List1.AddItem frmpop.hides(x).Caption
List2.AddItem frmpop.hides(x).Tag
List3.AddItem x
Next
End Sub

Private Sub List1_Click()

End Sub
