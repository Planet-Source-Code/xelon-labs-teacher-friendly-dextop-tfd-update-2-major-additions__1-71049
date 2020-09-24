Attribute VB_Name = "Module2"
Global Const GFM_STANDARD = 0
Global Const GFM_RAISED = 1
Global Const GFM_SUNKEN = 2
' Control Shadow Styles
Global Const GFM_BACKSHADOW = 1
Global Const GFM_DROPSHADOW = 2
' Color constants
Global Const BOX_WHITE& = &HFFFFFF
Global Const BOX_LIGHTGRAY& = &HC0C0C0
Global Const BOX_DARKGRAY& = &H808080
Global Const BOX_BLACK& = &H0&
'Here is shadow routine:


Sub FormControlShadow(f As Form, C As Control, shadow_effect As Integer, shadow_width As Integer, shadow_color As Long)
    'This routine is used to create a Back o
    '     r Drop shadow
    'effect on any controls which are placed
    '     on a form.
    'Simply place the control as normal and
    '     invoke the
    'shadow with the code below.
    '
    ' Parameters TypeComment
    ' fFormthe form containing the control
    ' CControl the control to shadow
    ' shadow_effect integer GFM_DROPSHADOW o
    '     r GFM_BACKSHADOW
    ' shadow_width integer width of the shad
    '     ow in pixels
    ' shadow_color longcolor of the shadow
    Dim shColor As Long
    Dim shWidth As Integer
    Dim oldWidth As Integer
    Dim oldScale As Integer
    
    shWidth = shadow_width
    shColor = shadow_color
    oldWidth = f.DrawWidth
    oldScale = f.ScaleMode
    
    f.ScaleMode = 3 'Pixels
    f.DrawWidth = 1


    Select Case shadow_effect
        Case GFM_DROPSHADOW
        f.Line (C.Left + shWidth, C.Top + shWidth)-Step(C.Width - 1, C.Height - 1), shColor, BF
        Case GFM_BACKSHADOW
        f.Line (C.Left - shWidth, C.Top - shWidth)-Step(C.Width - 1, C.Height - 1), shColor, BF
    End Select

f.DrawWidth = oldWidth
f.ScaleMode = oldScale
End Sub





 
 

