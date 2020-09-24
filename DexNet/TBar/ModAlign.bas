Attribute VB_Name = "ModAlign"
Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long
Enum DockTypes
    DockLeft = 1
        DockTop = 2
            DockRight = 3
                DockBottom = 4
                End Enum

Private Type vbRECT
    vbLeft As Long
    vbTop As Long
    vbWidth As Long
    vbHeight As Long
    End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
    Private Const SPIF_SENDWININICHANGE = &H2
    Private Const SPI_GETWORKAREA = 48
    Private Const SPI_SETWORKAREA = 47
    Private vbFormOldRect As vbRECT
    Private LastDock As DockTypes
    Private DockAmount As Integer

Sub UnDockForm(vbForm As Form)
On Error Resume Next
    Dim Desktop As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0&, Desktop, 0&
    With Desktop
        Select Case LastDock
            Case DockBottom
            .Bottom = .Bottom + DockAmount
            Case DockLeft
            .Left = .Left - DockAmount
            Case DockTop
            .Top = .Top - DockAmount
            Case DockRight
            .Right = .Right + DockAmount
            Case Else
            Exit Sub
        End Select
End With

With vbFormOldRect
    vbForm.Move .vbLeft, .vbTop, .vbWidth, .vbHeight
End With
SystemParametersInfo SPI_SETWORKAREA, 0&, Desktop, SPIF_SENDWININICHANGE
LastDock = 0

End Sub

Sub DockForm(vbForm As Form, DockPos As DockTypes)
On Error Resume Next
    If LastDock <> 0 Then

        MsgBox "Please don't re-dock without un-docking.", vbOKOnly, "Docking aborted"
        Exit Sub
    End If


    With vbFormOldRect
        .vbHeight = vbForm.Height
        .vbLeft = vbForm.Left
        .vbTop = vbForm.Top
        .vbWidth = vbForm.Width
    End With
    Dim Desktop As RECT

    SystemParametersInfo SPI_GETWORKAREA, 0&, Desktop, 0&
    Dim V As vbRECT
    V = vbFormOldRect

    With V
        Select Case DockPos
            Case DockLeft
            .vbTop = (Desktop.Top * 15)
            .vbLeft = (Desktop.Left * 15)
            .vbHeight = (Desktop.Bottom * 15) - .vbTop
            Case DockRight
            .vbTop = (Desktop.Top * 15)
            .vbLeft = (Desktop.Right * 15) - .vbWidth
            .vbHeight = (Desktop.Bottom * 15) - .vbTop
            Case DockBottom
            .vbTop = (Desktop.Bottom * 15) - .vbHeight
            .vbLeft = (Desktop.Left * 15)
            .vbWidth = (Desktop.Right * 15) - .vbLeft
            Case DockTop
            .vbTop = (Desktop.Top * 15)
            .vbLeft = (Desktop.Left * 15)
            .vbWidth = (Desktop.Right * 15) - .vbLeft
            Case Else
            Exit Sub
        End Select
End With

With Desktop

    Select Case DockPos
        Case DockBottom
        DockAmount = (vbForm.Height / 15)
            .Bottom = .Bottom - DockAmount
            
            Case DockRight
            DockAmount = (vbForm.Width / 15)
                .Right = .Right - DockAmount
                
                Case DockTop
                DockAmount = (vbForm.Height / 15)
                    .Top = .Top + DockAmount
                    
                    Case DockLeft
                    DockAmount = (vbForm.Width / 15)
                        .Left = .Left + DockAmount
                    End Select
            End With
            SystemParametersInfo SPI_SETWORKAREA, 0&, Desktop, SPIF_SENDWININICHANGE
            With V
                vbForm.Move .vbLeft, .vbTop, .vbWidth, .vbHeight
            End With
            LastDock = DockPos
        End Sub
