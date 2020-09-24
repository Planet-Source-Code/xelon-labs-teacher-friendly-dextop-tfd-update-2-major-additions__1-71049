VERSION 5.00
Begin VB.UserControl aicAlphaImage 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   MaskColor       =   &H80000014&
   PropertyPages   =   "aicAlphaImage.ctx":0000
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   Windowless      =   -1  'True
End
Attribute VB_Name = "aicAlphaImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Credits/Acknowledgements
'   Relies almost totally on my c32bppDIB Suite project. Credits included in that project
'       http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1
'   Paul Caton for his thunking routines.
'       Timer callbacks created using his code
' For most current updates/enhancements visit the following:
'   Visit http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67466&lngWId=1

' See the Usage.RTF file provided for more information

' Common Public Events
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object"
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object"
Public Event MouseExit()
Attribute MouseExit.VB_Description = "Occurs when the user first moves the mouse cursor out of the control"
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user first moves the mouse cursor into the control"
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus"
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus"
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual"
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual"

Public Event FadeTerminated(ByVal CurrentOpacity As Long)
Private z_CbMem   As Long    'Callback allocated memory address
Private z_Cb()    As Long    'Callback thunk array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------

' Timer and HitTest related APIs
Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long

' Drawing related APIs
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetClipBox Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long

' Window properties related APIs
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32.dll" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER As Long = 4

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MASKUSAGE
    color As Long           ' current mask color
    Applied As Boolean      ' mask has been applied
    AppliedColor As Long    ' color used to create mask; may not be same as Color
    Source As aiMaskSource  ' mask option: see aiMaskSource enum
End Type
Private Type FADERCONTROL
    tmrAddr As Long         ' AddressOf timer call back procedure
    fStep As Long           ' percent of opacity to change between steps
    fDelay As Long          ' length to delay before next step occurs
    fOpacity As Long        ' final opacity that also terminates the fader
End Type
Private Type SCALEDCOORD
    Left As Long            ' position of image within usercontrol
    Top As Long
    Width As Long           ' size of image within usercontrol
    Height As Long
    RotatedSize As Long     ' when rotated, the size needed to completely rotate image 360 degrees
    OneToOne As Boolean     ' flag used for painting; when image is actual size, faster renders
End Type

Public Enum aiMaskSource
    aiNoMask = 0
    aiUseMaskColor = 1
    aiUseTopLeft = 2
    aiUseTopRight = 3
    aiUseBottomLeft = 4
    aiUseBottomRight = 5
End Enum
Public Enum aiMirrorEnum
    aiMirrorNone = 0
    aiMirrorHorizontal = 1
    aiMirrorVertical = 2
    aiMirrorAll = 3
End Enum
Public Enum aiScaleMethod
    aiScaled = 0
    aiStretch = 1
    aiScaleDownOnly = 2
    aiActualSize = 3
    aiLockScale = 4
End Enum
Public Enum aiGrayScales
    aiNTSCPAL = 1     ' R=R*.299, G=G*.587, B=B*.114 - Default
    aiCCIR709 = 2     ' R=R*.213, G=G*.715, B=B*.072
    aiSimpleAvg = 3   ' R,G, and B = (R+G+B)/3
    aiRedMask = 4     ' uses only the Red sample value: RGB = Red / 3
    aiGreenMask = 5   ' uses only the Green sample value: RGB = Green / 3
    aiBlueMask = 6    ' uses only the Blue sample value: RGB = Blue / 3
    aiRedGreenMask = 7 ' uses Red & Green sample value: RGB = (Red+Green) / 2
    aiBlueGreenMask = 8 ' uses Blue & Green sample value: RGB = (Blue+Green) / 2
    aiNoGrayScale = 0
End Enum
Public Enum aiHitTestStyle  ' see HitTest property
    aiBoundingRgn = 1
    aiEnclosedRgn = 2
    aiShapedRgn = 3
    aiEntireControl = 0
End Enum
Public Enum aiOLEDropMode
    aiDropNone = vbOLEDropNone
    aiDropManual = vbOLEDropManual
End Enum

Private cKeyProps As Long
'1=HighQuality,2=Stretch,4=AutoSize,8=AutoRedraw;16=KeepBytes;32=OffscreenActive,64=MaskUsed

'//Rotation related variables
Private cRotated As Boolean
Private cAngle As Long

Private cHitTest As aiHitTestStyle
Private cRegion As Long     ' used when cHitTest is aiShapedRgn, aiEnclosedRgn
Private cRgnBox As RECT     ' used when cHitTest is aiEntireControl, aiBoundingRgn

Private cGrayScale As aiGrayScales
Private cScaleMode As ScaleModeConstants    ' parent container's scalemode; used for public events
Private cScaleMethod As aiScaleMethod
Private cMirror As aiMirrorEnum
Private cOpacity As Long
Private cMask As MASKUSAGE
Private cScaledCoords As SCALEDCOORD
Private cKeepOrigFormat As Boolean          ' testing. Still playing with ideas for this

Private cImage As c32bppDIB
Private cOffscreen As c32bppDIB             ' used when AutoRedraw=True

'//Timer & mouse enter/exit related variables
Private cProjOwner As Long
Private cPropKey As String
Private cTmrAddrOf As Long
Private cTmrHwnd As Long
Private cTopLeftPos As POINTAPI
Private cFader As FADERCONTROL
Public Sub ClearImage()
Attribute ClearImage.VB_Description = "Removes image from control"
End Sub
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object"
End Sub
Public Function LoadImage_FromArray(inStream() As Byte, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromArray.VB_Description = "Option to load an image from a stream of data"
End Function
Public Function LoadImage_FromFile(filename As String, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromFile.VB_Description = "Option to load an image from a file"
End Function
Public Function LoadImage_FromStdPicture(stdPic As StdPicture) As Boolean
Attribute LoadImage_FromStdPicture.VB_Description = "Option to load an image from a standard picture object"
End Function
Public Function LoadImage_FromClipboard() As Boolean
Attribute LoadImage_FromClipboard.VB_Description = "Option to load an image from the clipboard"
End Function
Public Function LoadImage_FromHandle(Handle As Long) As Boolean
Attribute LoadImage_FromHandle.VB_Description = "Option to load an image from an existing memory handle"
End Function
Public Function LoadImage_FromResource(VBglobal As IUnknown, ResSection As Variant, ResID As Variant, Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
Attribute LoadImage_FromResource.VB_Description = "Option to load an image from a resource file"
End Function
Public Function LoadImage_FromOrignalBytes(Optional desiredIconWidth As Long, Optional desiredIconHeight As Long, Optional desiredIconBitDepth As Long) As Boolean
End Function

Public Function GetImageBytes(imgBytes() As Byte, ByRef scanWidth As Long, _
                                Optional ByVal asArray2D As Boolean = False, _
                                Optional ByVal asBGRformat As Boolean = True, _
                                Optional ByVal asBottomUp As Boolean = False, _
                                Optional ByVal asPremultiplied As Boolean = False) As Boolean
    End Function
Public Function SetImageBytes(imgBytes() As Byte, Optional ByVal isArray2D As Boolean = False, _
                                Optional ByVal isBGRformat As Boolean = True, _
                                Optional ByVal isBottomUp As Boolean = False) As Boolean
End Function

Public Sub GetImageScales(ByRef Width As Long, ByRef Height As Long, _
            Optional ByVal ScaleMethod As aiScaleMethod = -1, _
            Optional ByVal desiredWidth As Long, Optional ByVal desiredHeight As Long, _
            Optional ByVal asRotatedImage As Boolean = False)
End Sub
            

Public Sub FadeInOut(ByVal FinalOpacity As Long, Optional ByVal FadeStepPercent As Long = 10, Optional ByVal FadeDelayInterval As Long = 30)
End Sub
Public Property Let AutoRedraw(Enable As Boolean)
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap"
End Property
Public Property Get AutoRedraw() As Boolean
End Property
Public Property Let HitTest(Style As aiHitTestStyle)
Attribute HitTest.VB_Description = "Returns/Sets method used to determine whether control responds to mouse events"
End Property
Public Property Get HitTest() As aiHitTestStyle
End Property
Public Property Let MaskColor(color As OLE_COLOR)
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the image"
End Property
Public Property Get MaskColor() As OLE_COLOR
End Property
Public Property Let MaskUsed(Style As aiMaskSource)
Attribute MaskUsed.VB_Description = "Returns/Sets whether the mask is to be applied to the image"
End Property
Public Property Get MaskUsed() As aiMaskSource
End Property
Public Property Let InversedImage(Inverse As Boolean)
Attribute InversedImage.VB_Description = "Returns/Sets whether the image colors are inverted"
End Property
Public Property Get InversedImage() As Boolean
End Property
Public Property Let AutoSize(Value As Boolean)
End Property
Public Property Get AutoSize() As Boolean
End Property
Public Property Let ScaleMethod(Method As aiScaleMethod)
End Property
Public Property Get ScaleMethod() As aiScaleMethod
End Property
Public Property Let StretchQuality(highQuality As Boolean)
Attribute StretchQuality.VB_Description = "Returns/sets whether a graphic will be resized using the best sizing algorithms"
End Property
Public Property Get StretchQuality() As Boolean
End Property
Public Property Let Opacity(ByVal Opaqueness As Long)
Attribute Opacity.VB_Description = "Returns/Sets the level of translucency for the control. 100 is fully opaque and 0 is transparent"
End Property
Public Property Get Opacity() As Long
End Property
Public Property Let Mirror(MirrorType As aiMirrorEnum)
Attribute Mirror.VB_Description = "Returns/Sets the current mirroring effect for the image"
End Property
Public Property Get Mirror() As aiMirrorEnum
End Property
Public Property Let KeepOriginalFormat(bValue As Boolean)
Attribute KeepOriginalFormat.VB_Description = "Returns/Sets whether control will maintain original image data"
End Property
Public Property Get KeepOriginalFormat() As Boolean
End Property
Public Property Let grayScale(Style As aiGrayScales)
Attribute grayScale.VB_Description = "Returns/Sets gray scale formula used when rendering image"
End Property
Public Property Get grayScale() As aiGrayScales
End Property
Public Property Let Rotation(ByVal newAngle As Long)
End Property
Public Property Get Rotation() As Long
End Property
Public Property Let Enabled(Enable As Boolean)
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events"
Attribute Enabled.VB_UserMemId = -514
End Property
Public Property Let MousePointer(Pointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object"
End Property
Public Property Get MousePointer() As MousePointerConstants
End Property
Public Property Let MouseIcon(icon As StdPicture)
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon"
End Property
Public Property Set MouseIcon(icon As StdPicture)
End Property
Public Property Get MouseIcon() As StdPicture
End Property
Public Property Let OLEDropMode(Value As aiOLEDropMode)
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target"
End Property
Public Property Get OLEDropMode() As aiOLEDropMode
End Property
Friend Function ppgGetStream(outStream() As Byte) As Boolean
End Function
Friend Sub ppgSetStream(inStream() As Byte, cX As Long, cY As Long, bitDepth As Long)
End Sub
Friend Property Get ppgDIBclass() As c32bppDIB
End Property
Friend Sub iccRemoteMouseExit()
End Sub
Private Sub sptReplaceImage()
End Sub
Private Sub sptRefreshRegion()
End Sub
Private Function sptConvertSysColor(color As Long) As Long
End Function
Private Function sptUpdateOffscreen(bResize As Boolean, bUpdateRegion As Boolean) As Boolean
End Function
Private Sub sptResize()
End Sub
Private Sub sptUndoMask()
End Sub
Private Sub sptMirrorImage(newMirrorValue As aiMirrorEnum)
End Sub
Private Sub sptValidateSession()
End Sub
Private Sub sptInvalidateSession()
End Sub
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long
End Function
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
End Function
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
End Function
Private Sub zTerminate()
End Sub
Private Function Timer_Fader(ByVal hwnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
End Function

Private Function Timer_MouseExit(ByVal hwnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
End Function
