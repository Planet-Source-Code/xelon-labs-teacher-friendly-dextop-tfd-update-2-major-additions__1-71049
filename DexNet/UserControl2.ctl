VERSION 5.00
Begin VB.UserControl UserControl2 
   ClientHeight    =   1770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ScaleHeight     =   1770
   ScaleWidth      =   1455
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   2520
      Width           =   2415
   End
End
Attribute VB_Name = "UserControl2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Extra As Integer
Public name As String
Event Click()

Private Sub List1_Click()
On Error Resume Next
List1.ListIndex = List2.ListIndex
name = List1.text
Extra = List2.ListIndex
RaiseEvent Click
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
List1.ListIndex = List2.ListIndex
name = List1.text
Extra = List2.ListIndex
RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
List1.Top = 0
List1.Left = 0
List1.Width = UserControl.Width
List1.Height = UserControl.Height
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
List1.Width = UserControl.Width
List1.Height = UserControl.Height
End Sub
Public Sub Add(name As String, Tag As Integer)
On Error Resume Next
List1.Additem name
List2.Additem Tag
End Sub
