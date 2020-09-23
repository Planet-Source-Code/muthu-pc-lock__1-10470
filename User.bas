Attribute VB_Name = "User"
Option Explicit

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long 'username only

Public Function GetCurrentUser()
Dim lpBuff As String * 25
Dim ret As Long ', username As String
ret = GetUserName(lpBuff, 25)
GetCurrentUser = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
frmPassword.U.Text = GetCurrentUser

End Function
