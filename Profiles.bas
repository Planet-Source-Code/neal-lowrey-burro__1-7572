Attribute VB_Name = "Profiles"
Option Explicit

Global Const MAX_N_USERS = 25        'maximum number of contemporary users
Global Const N_RECOGNIZED_USERS = 3 'number of recognized users
Global Const DEFAULT_DRIVE = "D:"   'default drive

Global Privtyp As Privtyp

'Type UserInfo
'  Name As String 'list of the users which can access to server file-system
'  Pass As String 'list of passwords of each user which can access to server file-system
'  Pcnt As Integer
'  Priv(20) As Privtyp
'  Home As String 'default directory of each user
'End Type

Type User_IDs
  Count As Integer
  No(0 To MAX_N_USERS) As UserInfo
End Type

Global UserIDs As User_IDs
'the list of the access rights of each user,
'every element is a string formed by 2 characters:
'the 2nd char. is relative to write & delete right
'(Y=Yes, N=No).

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName _
    As String) As Integer

Declare Function WritePrivateProfileString% Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName$, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal _
    lpFileName$)
Global Version As Integer
Global CurrentProfile As String
'
'   Loads program settings from disk.
'
Public Function LoadProfile(ByVal Filename As String) As Boolean
  Dim tStr As String
  Dim Ctr As Integer, x As Integer, Pcnt As Integer
  Dim i As Integer, Number As Integer
  '
  '   Check for existence of INI file
  '
  On Error Resume Next
  Ctr = FileLen(Filename)
  If Err.Number > 0 Then
    Err.Clear
    LoadProfile = False
    Exit Function
  End If
  On Error Resume Next
  LoadProfile = True
  If Ctr < 1 Then      ' ini file empty
    Exit Function
  End If
  '
  '   Load saved settings
  '
  Version = Val(GetFromIni("Settings", "Version", Filename))
  If Len(Version) < 1 Then
    LoadProfile = False
    Exit Function
  End If
  '   Load Users
  Number = Val(GetFromIni("Users", "Users", Filename))
  UserIDs.Count = Number
  If Number > 0 Then
    For Ctr = 1 To Number
      UserIDs.No(Ctr).Name = GetFromIni("Users", "Name" & Ctr, Filename)
      UserIDs.No(Ctr).Pass = GetFromIni("Users", "Pass" & Ctr, Filename)
      Pcnt = Val(GetFromIni("Users", "DirCnt" & Ctr, Filename))
      UserIDs.No(Ctr).Pcnt = Pcnt
      Debug.Print "User:" & Ctr & ", DirCnt=" & Pcnt
      For x = 1 To Pcnt
        tStr = GetFromIni("Users", "Access" & Ctr & "_" & x, Filename)
        i = InStr(tStr, ",")
        UserIDs.No(Ctr).Priv(x).Path = Left(tStr, i - 1)
        UserIDs.No(Ctr).Priv(x).Accs = Right(tStr, (Len(tStr) - i))
      Next
      UserIDs.No(Ctr).Home = GetFromIni("Users", "Home" & Ctr, Filename)
    Next
  End If
  CurrentProfile = Filename
End Function
'
'   Saves program settings to disk.
'
Public Function SaveProfile(ByVal Filename As String, SaveSettings As Boolean) As Boolean
  Dim Terminal As String, Alias As String
  Dim Ctr As Integer, x As Integer
  SaveProfile = False
  If SaveSettings Then
   ' SettingsChanged = False
    If WritePrivateProfileString("Settings", "Version", _
        App.Major & "." & App.Minor & "." & App.Revision, Filename) = 0 Then
      SaveProfile = False
      Exit Function
    End If

    WritePrivateProfileString "Users", "Users", CStr(UserIDs.Count), Filename
    For Ctr = 1 To UserIDs.Count
      WritePrivateProfileString "Users", "Name" & Ctr, CStr(UserIDs.No(Ctr).Name), Filename
      WritePrivateProfileString "Users", "Pass" & Ctr, UserIDs.No(Ctr).Pass, Filename
      WritePrivateProfileString "Users", "DirCnt" & Ctr, CStr(UserIDs.No(Ctr).Pcnt), Filename
      For x = 1 To UserIDs.No(Ctr).Pcnt
        WritePrivateProfileString "Users", "Access" & Ctr & "_" & x, _
          UserIDs.No(Ctr).Priv(x).Path & "," & UserIDs.No(Ctr).Priv(x).Accs, Filename
        WritePrivateProfileString "Users", "Home" & Ctr, CStr(UserIDs.No(Ctr).Home), Filename
      Next
    Next

    CurrentProfile = Filename
    SaveProfile = True
  End If
End Function
'
'   Gets a string from an INI file.
'
Public Function GetFromIni(strSectionHeader As String, strVariableName As _
    String, strFileName As String) As String
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    GetFromIni = Left$(strReturn, _
      GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", _
      strReturn, Len(strReturn), strFileName))
End Function


