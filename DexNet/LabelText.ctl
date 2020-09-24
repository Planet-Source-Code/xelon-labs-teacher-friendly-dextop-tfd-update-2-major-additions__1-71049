VERSION 5.00
Begin VB.UserControl LabelText 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox Label1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Text            =   "Label"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Browse 
      Caption         =   "..."
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "LabelText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Browsed()
Event Change()
Event KeyPress(KeyAscii As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Dim bro As Integer


Private Sub Browse_Click()
On Error Resume Next
RaiseEvent Browsed
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub Text1_Change()
On Error Resume Next
RaiseEvent Change

End Sub

Private Sub Text1_GotFocus()
On Error Resume Next
Label1.BackColor = 8421504
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
On Error Resume Next
Label1.BackColor = &HFFFFFF
End Sub

Private Sub UserControl_GotFocus()
On Error Resume Next
Apply
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
On Error Resume Next
Label1.Height = UserControl.Height
Text1.Height = UserControl.Height
Text1.Width = UserControl.Width - Label1.Width
Browse.Left = UserControl.Width - Browse.Width
Browse.Height = UserControl.Height
Apply
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next
Apply
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    Label1.text = PropBag.ReadProperty("Caption", Label1)
    Text1.text = PropBag.ReadProperty("Text", "")
    bro = PropBag.ReadProperty("Button", "0")
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Label1.Height = UserControl.Height
Text1.Height = UserControl.Height
Text1.Width = UserControl.Width - Label1.Width
Browse.Left = UserControl.Width - Browse.Width
Browse.Height = UserControl.Height
Apply
End Sub

Public Sub Apply()
On Error Resume Next
Text1.text = text
Label1.text = Caption
End Sub
Public Sub Set_Browse()
On Error Resume Next
Browse.Visible = True
End Sub
Public Sub Hide_Browse()
On Error Resume Next
Browse.Visible = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("Caption", Label1.text, "Label")
    Call PropBag.WriteProperty("Text", Text1.text, "")
    Call PropBag.WriteProperty("Button", bro, bro)

End Sub
Public Property Get Caption() As String
On Error Resume Next
    Caption = Label1
End Property

Public Property Let Caption(ByVal New_Caption As String)
On Error Resume Next
    Label1 = New_Caption
    PropertyChanged "Caption"
End Property
Public Property Get text() As String
On Error Resume Next
    text = Text1
End Property

Public Property Let text(ByVal New_Caption As String)
On Error Resume Next
    Text1 = New_Caption
    PropertyChanged "Text"
End Property

Public Property Let Button(ByVal New_Button As String)
On Error Resume Next
    If New_Button = 1 Then
    Set_Browse
    bro = 1
    Else
    Hide_Browse
    bro = 0
    End If
    PropertyChanged "Button"
End Property
Sub pword(char As String)
On Error Resume Next
Text1.PasswordChar = char
End Sub

Sub Sfocus()
On Error Resume Next
Text1.SetFocus
End Sub
