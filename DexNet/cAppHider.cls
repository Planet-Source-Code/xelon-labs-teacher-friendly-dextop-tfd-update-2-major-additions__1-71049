VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAppHider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const RSP_SIMPLE_SERVICE = 1
Const RSP_UNREGISTER_SERVICE = 0

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Public Sub HideApplication()
On Error Resume Next
  
  Dim ws As cWinVersion
  
  Set ws = New cWinVersion
  
  'Hide application depending on the Windows Version
  '-----------------------------------------
  Select Case ws.Version
  
    Case EnumWindowsVersion.wvWindows9x
      MakeMeService
      
    Case EnumWindowsVersion.wvWindowsNT
      App.Title = ""
  
  End Select
  '-----------------------------------------
  
End Sub

Public Sub ShowApplication(ByVal AppTitle As String)
On Error Resume Next
    
  Dim ws As cWinVersion
  
  Set ws = New cWinVersion
  
  'Show App
  '-----------------------------------------
  Select Case ws.Version
  
    Case EnumWindowsVersion.wvWindows9x
      UnMakeMeService
      
    Case EnumWindowsVersion.wvWindowsNT
      App.Title = AppTitle
  
  End Select
  '-----------------------------------------
  
End Sub

Private Sub MakeMeService()
On Error Resume Next

  Dim PId As Long, regserv As Long
  
  'Get the current process ID
  PId = GetCurrentProcessId()
  
  'Register as service
  regserv = RegisterServiceProcess(PId, RSP_SIMPLE_SERVICE)

End Sub

Private Sub UnMakeMeService()
On Error Resume Next

  Dim PId As Long, regserv As Long
  
  'Get the current process ID
  PId = GetCurrentProcessId()
  
  'Unregister as service
  regserv = RegisterServiceProcess(PId, RSP_UNREGISTER_SERVICE)

End Sub

