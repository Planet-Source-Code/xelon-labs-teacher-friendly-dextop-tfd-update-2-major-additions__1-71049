Attribute VB_Name = "modWorm"
Public Type Dot
 Left As Double
 Top As Double
 Visible As Boolean
End Type

Public Type typeApple
    Left As Double
    Top As Double
    Width As Long
    Height As Long
    pic As Integer
    Visible As Boolean
End Type

Public bLoaded As Boolean 'Whether or not the game has been loaded yet

'Bitblt for the game's graphics
Declare Function BitBlt Lib "GDI32" ( _
        ByVal hDestDC As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal XSrc As Long, _
        ByVal YSrc As Long, _
        ByVal dwRop As Long) As Long

Public GameType As Long
Public Difficulty As Long
Public Speed As Long
Public Control As Long
Public Multiplier As Long
Public HighScore(0 To 5) As Long
Public Size As Long

'Used for collision detection
Public Declare Function GetPixel Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Wave Functions
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Midi functions
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

'Ini Functions
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Function GetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
On Error Resume Next
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
On Error Resume Next
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
Sub PlayMidi(strMidi As String)
On Error Resume Next
strMidi = GetShortPath(App.path) & "\" & strMidi & ".mid"
If strMidi = "" Then Exit Sub
Call mciSendString("play " & strMidi$, 0&, 0, 0)
End Sub
Sub StopMidi(strMidi As String)
On Error Resume Next
strMidi = GetShortPath(App.path) & "\" & strMidi & ".mid"
If strMidi = "" Then Exit Sub
Call mciSendString("stop " & strMidi$, 0&, 0, 0)
End Sub
Public Function GetShortPath(strFileName As String) As String
On Error Resume Next
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function
