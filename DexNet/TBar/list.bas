Attribute VB_Name = "list"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Sub Write_INIList(ini As String, lst As ListBox)
On Error Resume Next
WriteIni "Main", lst.Name & " Count", lst.ListCount - 1, ini
Dim x As Integer
For x = 0 To lst.ListCount - 1
WriteIni "Main", lst.Name & x, lst.list(x), ini
Next
End Sub

Sub Get_INIList(ini As String, lst As ListBox)
On Error Resume Next
Dim cnt As String
cnt = GetFromIni("Main", lst.Name & " Count", ini)
Dim x As Integer
For x = 0 To Val(cnt) - 1
lst.AddItem GetFromIni("Main", lst.Name & x, ini)
Next
End Sub

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
