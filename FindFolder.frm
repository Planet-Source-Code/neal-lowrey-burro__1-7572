VERSION 5.00
Begin VB.Form FindFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FindFolder"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "FindFile"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton FldrDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4600
      Width           =   1575
   End
   Begin VB.TextBox DirPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.ListBox FolderList 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      ItemData        =   "FindFolder.frx":0000
      Left            =   120
      List            =   "FindFolder.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "FindFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DrvS(32) As String
Dim LastStr As String
Dim DrvC As Integer

Private Sub FldrDone_Click()
  Form_Terminate
End Sub

Private Sub FolderList_Click()
Dim s As String, t As String, s2 As String
Dim i As Integer
  i = FolderList.ListIndex + 1
  s2 = FolderList.Text
  If Mid(s2, 1, 1) = "[" Then
    s2 = Mid(s2, 2, 2) & "\"
    DirPath = s2
  Else
    If FolderList.Text = ".." Then
      s = Left(LastStr, Len(LastStr) - 1)
      Do Until Right(s, 1) = "\"
        s = Left(s, Len(s) - 1)
      Loop
      s2 = s
      DirPath = s2
    Else
      s2 = DirPath & FolderList.Text & "\"
      DirPath = s2
    End If
  End If
  LastStr = s2
  FolderList.Clear
  'Debug.Print i; s2
  s = FindFile("*.*", s2)
  Add_Drives
End Sub

Private Sub Form_Load()
Dim s As String
  GetSystemDrives 'load the system drives
  If AddEditDir.Tag <> "" Then
    LastStr = AddEditDir.Tag
    DirPath = LastStr
    s = FindFile("*.*", AddEditDir.Tag)
  End If
  Add_Drives
End Sub

Private Sub Add_Drives()
Dim x As Integer
  For x = 1 To DrvC
    FolderList.AddItem "[" & DrvS(x) & "]"
  Next
End Sub
Private Sub Form_Terminate()
  AddEditDir.Tag = DirPath.Text
  Unload Me
End Sub

Private Sub GetSystemDrives()
Dim rtn As Long
Dim d As Integer
Dim AllDrives As String
Dim CurrDrive As String
Dim tmp As String
  tmp = Space(64)
  rtn = GetLogicalDriveStrings(64, tmp)
  AllDrives = Trim(tmp)               'get the list of all available drives
  d = 0
  Do Until AllDrives = Chr$(0)
    d = d + 1
    CurrDrive = StripNulls(AllDrives) 'strip off one drive item from the allDrives
    CurrDrive = Left(CurrDrive, 2)    'we can't have the trailing slash, so ..
    DrvS(d) = CurrDrive
    DrvC = d
  Loop
End Sub

Private Function StripNulls(startstr) As String
Dim pos As Integer
  pos = InStr(startstr, Chr$(0))
  If pos Then
    StripNulls = Mid(startstr, 1, pos - 1)
    startstr = Mid(startstr, pos + 1, Len(startstr))
    Exit Function
  End If
End Function


