Attribute VB_Name = "FndFile"
Option Explicit

Public Const MAX_PATH As Long = 260
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Type FileTime
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FileTime
  ftLastAccessTime As FileTime
  ftLastWriteTime As FileTime
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
    
Public Function FindFile(ByVal Filename As String, ByVal Path As String) As String
Dim hFile As Long, result As Long
Dim ts As String, szPath As String
Dim WFD As WIN32_FIND_DATA
Dim szPath2 As String, szFilename As String
Dim dwBufferLen As Long, szBuffer As String, lpFilePart As String
  szPath = GetRDP(Path) & "*.*" & Chr$(0)
  szPath2 = Path & Chr$(0)
  szFilename = Filename & Chr$(0)
  szBuffer = String$(MAX_PATH, 0)
  dwBufferLen = Len(szBuffer)
  result = SearchPath(szPath2, szFilename, vbNullString, dwBufferLen, szBuffer, lpFilePart)
  If result Then
    FindFile = StripNull(szBuffer)
    Exit Function
  End If
  hFile = FindFirstFile(szPath, WFD)  'Start asking windows for files.
  Do
    If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
      ts = StripNull(WFD.cFileName)
      If ts <> "." Then
        FindFolder.FolderList.AddItem (ts)
      End If
    End If
    WFD.cFileName = ""
    result = FindNextFile(hFile, WFD)
  Loop Until result = 0
  FindClose hFile
End Function

Public Function StripNull(ByVal WhatStr As String) As String
  If InStr(WhatStr, Chr$(0)) > 0 Then
    StripNull = Left$(WhatStr, InStr(WhatStr, Chr$(0)) - 1)
  Else
    StripNull = WhatStr
  End If
End Function

Public Function GetRDP(ByVal sPath As String) As String
'Adds a backslash on the end of a path, if required.
  If sPath = "" Then Exit Function
  If Right$(sPath, 1) = "\" Then GetRDP = sPath: Exit Function
    GetRDP = sPath & "\"
End Function
