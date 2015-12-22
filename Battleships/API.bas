Attribute VB_Name = "API"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Code written by David C. Thompson                       *
'*                                                         *
'* Copyright Restrictions:                                 *
'*     This code is here for reference, and may NOT be     *
'*     sold or leased under any circumstances. Changes     *
'*     will be allowed to be made so long as the original  *
'*     author's name is retained in the credits.           *
'*                                                         *
'* February 2003                                           *
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Sub ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLenght As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1
Public strwindir$
Const NoError = 0


Public Function GetTheComputerName()
    Const lpnLength As Integer = 255     ' Buffer size for the return string. Get return buffer space.
    Dim status As Integer                ' For getting user information.
    Dim lpUserName As String     ' Assign the buffer size constant to lpUserName.
    
    lpUserName = Space$(lpnLength + 1)   ' Get the log-on name of the person using product.
    status = GetComputerName(lpUserName, lpnLength) ' See whether error occurred.
    lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    ' This line removes the null character. Strings in C are null-
    ' terminated. Strings in Visual Basic are not null-terminated.
    ' The null character must be removed from the C strings to be used
    ' cleanly in Visual Basic.
    GetTheComputerName = lpUserName
End Function


Public Function Find_WinDir()
  Dim lpBuffer As String * 144

  Dim Length%
  Length% = GetWindowsDirectory(lpBuffer, Len(lpBuffer))
  strwindir = Left(lpBuffer, Length%) & "\"
End Function

Sub Log_Out_PC()
  Const EWX_LOGOFF = 0
  Const EWX_SHUTDOWN = 1
  Const EWX_REBOOT = 2
  Const EWX_FORCE = 4
  Call ExitWindowsEx(EWX_LOGOFF, 0)
End Sub

Function GetUserName() As String
  Const lpnLength As Integer = 255
  Dim status As Integer
  Dim lpName, lpUserName As String
      
      lpUserName = Space$(lpnLength + 1)
      status = WNetGetUser(lpName, lpUserName, lpnLength)
      If status = NoError Then
         lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
        Else
         lpUserName = "error"
      End If
      GetUserName = LCase(lpUserName)
End Function


Function GoToWeb(URL$)
  Dim Success As Long

  Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)

End Function




