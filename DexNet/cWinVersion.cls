VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWinVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EnumWindowsVersion
  
  wvWindows32 = 0
  wvWindows9x = 1
  wvWindowsNT = 2
  
End Enum

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private m_OSInfo As OSVERSIONINFO

Public Property Get Version2() As String
On Error Resume Next
  
  Version2 = m_OSInfo.szCSDVersion

End Property

Public Property Get Build() As Long
On Error Resume Next
  
  Build = m_OSInfo.dwBuildNumber

End Property

Public Property Get Minor() As Long
On Error Resume Next
  
  Minor = m_OSInfo.dwMinorVersion

End Property

Public Property Get Major() As Long
On Error Resume Next
  
  Major = m_OSInfo.dwMajorVersion

End Property

Public Property Get Version() As EnumWindowsVersion
On Error Resume Next
  
  'Print the information to the form

  Select Case m_OSInfo.dwPlatformId

    Case 0
      Version = wvWindows32

    Case 1
      Version = wvWindows9x

    Case 2
      Version = wvWindowsNT

  End Select
  
End Property

Private Function m_Load()
On Error Resume Next
    
  Dim Ret&
    
  'Set the structure size
  m_OSInfo.dwOSVersionInfoSize = Len(m_OSInfo)
    
  'Get the Windows version
  Ret& = GetVersionEx(m_OSInfo)
    
  'Check for errors

  If Ret& = 0 Then

    err.Raise vbObjectError + 100, "GetVersionEx API", "Error Getting Version Information"
    Exit Function

  End If
            
End Function

Private Sub Class_Initialize()
On Error Resume Next
  
  m_Load
  
End Sub

