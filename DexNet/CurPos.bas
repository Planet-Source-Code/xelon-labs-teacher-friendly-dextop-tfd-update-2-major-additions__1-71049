Attribute VB_Name = "CurPos"
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_MOVE = &H1

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Public Sub MouseClick(Clickr As String)
On Error Resume Next
Select Case Clickr
    Case "Left"
        mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, cButt, dwEI
        mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, cButt, dwEI
    Case "Right"
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0&, 0&, cButt, dwEI
        mouse_event MOUSEEVENTF_RIGHTUP, 0&, 0&, cButt, dwEI
    Case "Middle"
        mouse_event MOUSEEVENTF_MIDDLEDOWN, 0&, 0&, cButt, dwEI
        mouse_event MOUSEEVENTF_MIDDLEUP, 0&, 0&, cButt, dwEI
    Case "Double Click"
        mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, cButt, dwEI
        mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, cButt, dwEI
        mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, cButt, dwEI
        mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, cButt, dwEI
    Case "Middle Up"
    Case "Middle Down"
End Select
End Sub
