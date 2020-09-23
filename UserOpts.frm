VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form UserOpts 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Options"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton UsrDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   2640
      TabIndex        =   25
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setup"
      Height          =   5175
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   4575
      Begin VB.TextBox UsrName 
         Height          =   285
         Left            =   1080
         TabIndex        =   27
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox HomeDir 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Pword 
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.Frame frm1 
         Caption         =   "File/Dir Access Rules"
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4335
         Begin VB.CommandButton FDUpdate 
            Caption         =   "Update"
            Height          =   375
            Left            =   1920
            TabIndex        =   26
            Top             =   3000
            Width           =   735
         End
         Begin VB.CheckBox FRead 
            Caption         =   "Read"
            Height          =   255
            Left            =   3000
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox FWrite 
            Caption         =   "Write"
            Height          =   255
            Left            =   3000
            TabIndex        =   16
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox FDelete 
            Caption         =   "Delete"
            Height          =   255
            Left            =   3000
            TabIndex        =   15
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox FEx 
            Caption         =   "Execute"
            Height          =   255
            Left            =   3000
            TabIndex        =   14
            Top             =   1200
            Width           =   975
         End
         Begin VB.CheckBox DList 
            Caption         =   "List"
            Height          =   255
            Left            =   3000
            TabIndex        =   13
            Top             =   1800
            Width           =   615
         End
         Begin VB.CheckBox DMake 
            Caption         =   "Make"
            Height          =   255
            Left            =   3000
            TabIndex        =   12
            Top             =   2040
            Width           =   735
         End
         Begin VB.CheckBox DRemove 
            Caption         =   "Remove"
            Height          =   255
            Left            =   3000
            TabIndex        =   11
            Top             =   2280
            Width           =   975
         End
         Begin VB.CheckBox DSub 
            Caption         =   "Inherit Subs"
            Height          =   255
            Left            =   3000
            TabIndex        =   10
            Top             =   2520
            Width           =   1215
         End
         Begin VB.ListBox AccsList 
            Height          =   2595
            ItemData        =   "UserOpts.frx":0000
            Left            =   120
            List            =   "UserOpts.frx":0002
            TabIndex        =   9
            Top             =   240
            Width           =   2655
         End
         Begin VB.CommandButton FDAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   240
            TabIndex        =   8
            Top             =   3000
            Width           =   615
         End
         Begin VB.CommandButton FDEdit 
            Caption         =   "Edit"
            Height          =   375
            Left            =   1080
            TabIndex        =   7
            Top             =   3000
            Width           =   615
         End
         Begin VB.CommandButton FDRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   2880
            TabIndex        =   6
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Files"
            Height          =   255
            Left            =   2880
            TabIndex        =   19
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Directories"
            Height          =   255
            Left            =   2880
            TabIndex        =   18
            Top             =   1560
            Width           =   975
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Home Dir:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Users"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton UsrRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton UsrAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   855
      End
      Begin VB.ListBox UserList 
         Height          =   4155
         ItemData        =   "UserOpts.frx":0004
         Left            =   120
         List            =   "UserOpts.frx":0006
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "UserOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim uItem As Integer
Dim aItem As Integer
Dim tStrng As String
Dim uUser As Integer
Dim Pcnt As Integer

Private Type Priv
  Path As String
  Accs As String '[R]ead,[W]rite,[D]elete,e[X]ecute > Files
                 '[L]ist,[M]ake,[K]ill,[S]ubs       > Dirs
End Type
Private Privs(20) As Priv

Private Sub FDAdd_Click()
  tStrng = Get_Path("")
  If tStrng <> "" Then
    AccsList.AddItem (tStrng)
    Pcnt = Pcnt + 1
    UserIDs.No(uUser).Priv(Pcnt).Path = tStrng
    FDUpdate.Enabled = True
    FDRemove.Enabled = True
  End If
  AccsList_False
End Sub

Private Sub FDEdit_Click()
  tStrng = Get_Path(AccsList.Text)
  If tStrng <> "" Then
    AccsList.List(aItem) = tStrng
    UserIDs.No(uUser).Priv(aItem + 1).Path = tStrng
  End If
  AccsList_False
End Sub

Private Sub FDRemove_Click()
Dim z As Integer
  For z = (aItem + 1) To UserIDs.No(uUser).Pcnt
    UserIDs.No(uUser).Priv(z).Path = UserIDs.No(uUser).Priv(z + 1).Path
    UserIDs.No(uUser).Priv(z).Accs = UserIDs.No(uUser).Priv(z + 1).Accs
  Next
  UserIDs.No(uUser).Pcnt = UserIDs.No(uUser).Pcnt - 1
  AccsList.RemoveItem (aItem)
  AccsList_False
End Sub

Private Sub FDUpdate_Click()
Dim z As Integer, s As String
  UserIDs.No(uUser).Name = UsrName
  UserIDs.No(uUser).Pass = Pword
  UserIDs.No(uUser).Home = HomeDir
  UserIDs.No(uUser).Pcnt = Pcnt
  s = ""
  z = aItem + 1
  If FRead.Value = 1 Then s = s & "R"
  If FWrite.Value = 1 Then s = s & "W"
  If FDelete.Value = 1 Then s = s & "D"
  If FEx.Value = 1 Then s = s & "X"
  If DList.Value = 1 Then s = s & "L"
  If DMake.Value = 1 Then s = s & "M"
  If DRemove.Value = 1 Then s = s & "K"
  If DSub.Value = 1 Then s = s & "S"
  Privs(z).Accs = s
  UserIDs.No(uUser).Priv(z).Accs = s
  AccsList_False
End Sub

Private Sub Form_Load()
Dim x As Integer, y As Integer
  y = UserIDs.Count
  If (y > 0) Then
    For x = 1 To UserIDs.Count
      UserList.AddItem UserIDs.No(x).Name
    Next
  End If
  aItem = -1
  uItem = -1
  AccsList_False
  UserList_False
  FDAdd.Enabled = False
End Sub

Private Sub Form_Terminate()
  Unload Me
End Sub

Private Sub UserList_LostFocus()
  ' If uItem >= 0 Then UserList_False
End Sub

Private Sub UsrDone_Click()
Dim z As Integer
  Form_Terminate
End Sub

Private Sub UsrRemove_Click()
Dim z As Integer, i As Integer
  z = UserIDs.Count
  For i = uUser To z
    UserIDs.No(i) = UserIDs.No(i + 1)
  Next
  UserList.RemoveItem (uItem)
  UserIDs.Count = z - 1
  AccsList.Clear
  ClearAccs
  UsrName = ""
  Pword = ""
  HomeDir = ""
  aItem = -1
  UserList_False
End Sub

Private Sub UsrAdd_Click()
Dim i As Integer, S1 As String
  S1 = "New User"
  UsrName = S1
  UserList.AddItem S1
  i = UserIDs.Count + 1
  UserIDs.No(i).Name = S1
  UserIDs.Count = i
  UserList_False
End Sub

Private Sub UserList_Click()
Dim x As Integer, z As Integer
  uItem = UserList.ListIndex
  Debug.Print "User List Item = " & uItem
  '[R]ead,[W]rite,[D]elete,e[X]ecute > Files
  '[L]ist,[M]ake,[K]ill,[S]ubs       > Dirs
  uUser = uItem + 1
  AccsList.Clear
  ClearAccs
  Pword = ""
  HomeDir = ""
  aItem = -1
  UserList_True
  AccsList_False
  FDAdd.Enabled = True
  UsrName = UserIDs.No(uUser).Name
  Pword = UserIDs.No(uUser).Pass
  HomeDir = UserIDs.No(uUser).Home
  Pcnt = UserIDs.No(uUser).Pcnt
  For z = 1 To Pcnt
    Privs(z).Path = UserIDs.No(uUser).Priv(z).Path
    Privs(z).Accs = UserIDs.No(uUser).Priv(z).Accs
    AccsList.AddItem Privs(z).Path
  Next
End Sub

Private Sub AccsList_Click()
Dim x As Integer, z As Integer
  aItem = AccsList.ListIndex
  Debug.Print "Access List Item = " & aItem
  ClearAccs
  AccsList_True
  z = aItem + 1
  Debug.Print UserIDs.No(uUser).Priv(z).Accs
  If InStr(Privs(z).Accs, "R") Then
    FRead.Value = 1
  End If
  If InStr(Privs(z).Accs, "W") Then
    FWrite.Value = 1
  End If
  If InStr(Privs(z).Accs, "D") Then
    FDelete.Value = 1
  End If
  If InStr(Privs(z).Accs, "X") Then
    FEx.Value = 1
  End If
  If InStr(Privs(z).Accs, "L") Then
    DList.Value = 1
  End If
  If InStr(Privs(z).Accs, "M") Then
    DMake.Value = 1
  End If
  If InStr(Privs(z).Accs, "K") Then
    DRemove.Value = 1
  End If
  If InStr(Privs(z).Accs, "S") Then
    DSub.Value = 1
  End If
End Sub

Private Sub AccsList_DblClick()
  aItem = AccsList.ListIndex
  tStrng = Get_Path(AccsList.Text)
  If tStrng <> "" Then
    AccsList.List(aItem) = tStrng
    UserIDs.No(uUser).Priv(aItem + 1).Path = tStrng
  End If
  AccsList.Selected(aItem) = False
End Sub

Private Sub UserList_True()
  UsrRemove.Enabled = True
End Sub

Private Sub UserList_False()
  Debug.Print "uItem=" & uItem
  UsrRemove.Enabled = False
  If uItem >= 0 Then
    UserList.Selected(uItem) = False
    uItem = -1
  End If
End Sub

Private Sub AccsList_True()
  FDEdit.Enabled = True
  FDRemove.Enabled = True
  FDUpdate.Enabled = True
End Sub

Private Sub AccsList_False()
  Debug.Print "aItem=" & aItem
  FDEdit.Enabled = False
  FDRemove.Enabled = False
  FDUpdate.Enabled = False
  If aItem >= 0 Then
    AccsList.Selected(aItem) = False
    aItem = -1
  End If
End Sub

Private Sub ClearAccs()
  FRead.Value = 0
  FWrite.Value = 0
  FDelete.Value = 0
  FEx.Value = 0
  DList.Value = 0
  DMake.Value = 0
  DRemove.Value = 0
  DSub.Value = 0
End Sub

Function Get_Path(olds As String) As String
  AddEditDir.DirPath = olds
  AddEditDir.Show 1
  If Tag <> "" Then
    Get_Path = Tag
    Tag = ""
  End If
End Function
