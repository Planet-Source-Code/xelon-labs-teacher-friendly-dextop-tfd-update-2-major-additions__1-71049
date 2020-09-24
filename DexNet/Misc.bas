Attribute VB_Name = "Module2"
Public Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type
Public frm As Integer
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const CS_DROPSHADOW = &H20000
Public Const GCL_STYLE = (-26)

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Public IsReady As Boolean
Public NotifyTop As Long

Sub DropShadow(hwnd As Long, Optional Silent As Boolean = True)
    On Error Resume Next
        SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Public Function MyManifestFile() As String
    On Error Resume Next
    MyManifestFile = FindPath(App.path, App.EXEName & ".exe.manifest")
End Function

Public Function XPVB(Optional ForceWriteManifest As Boolean = False) As Boolean
    On Error Resume Next
    If Dir(MyManifestFile) <> "" And ForceWriteManifest = False Then GoTo Written
    Dim XPStr As String
    Dim FF As Integer
    XPStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
            "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
            "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""Microsoft.VB6.VBnetStyles"" type=""win32""/>" & vbCrLf & _
            "<description>Windows XP manifest file</description>" & vbCrLf & "<dependency>" & vbCrLf & _
            "<dependentAssembly>" & vbCrLf & "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*""/>" & vbCrLf & _
            "</dependentAssembly>" & vbCrLf & "</dependency>" & vbCrLf & "</assembly>"
    FF = FreeFile
    Open MyManifestFile For Output As #FF
        Print #FF, XPStr
    Close #FF
Written:
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    XPVB = (err.Number = 0)
    On Error GoTo 0
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function
