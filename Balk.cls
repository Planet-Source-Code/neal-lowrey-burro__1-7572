VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Balk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
Dim Messages As MainApp
Event Finished(Code As Integer)

Public Sub SetUserData(UserData As User)
  Dim iL As Integer
  FTP_Index = UserData.list_index
  iL = UserData.list_index
  users(iL).IP_Address = UserData.IP_Address
  users(iL).list_index = UserData.list_index
  users(iL).control_slot = UserData.control_slot
  users(iL).data_slot = UserData.data_slot
  users(iL).IP_Address = UserData.IP_Address
  users(iL).Port = UserData.Port
  users(iL).data_representation = users(iL).data_representation
  users(iL).data_format_ctrls = UserData.data_format_ctrls
  users(iL).data_structure = UserData.data_structure
  users(iL).data_tx_mode = UserData.data_tx_mode
  'current working directory
  users(iL).cur_dir = UserData.cur_dir
  users(iL).State = UserData.State
  users(iL).full = UserData.full
  users(iL).temp_data = UserData.temp_data
End Sub

Public Sub SetUserPermissions(Perms As UserInfo, Users_Index As Integer)
  UserIDs.No(Users_Index) = Perms
  'Doing this to keep old security code
  'somewhat compatable

End Sub

Public Function GetUserData() As User
 ' GetUserData = users(FTP_Index)
End Function

Public Sub New_Cmd(Command As String, Arguments() As String)
  Dim iL As Integer
  Dim UB As Integer
  UB = UBound(Arguments())
  FTP_Command = Command
  ReDim FTP_Args(UB)
  For iL = 0 To UB
    FTP_Args(iL) = Arguments(iL)
  Next
  StartTimer
End Sub


Public Sub SetCallBack(MainApp As MainApp)
  Set Messages = MainApp
  SetThread Me
End Sub

Public Sub SendMessage(str As String)
  Messages.SvrLogToScreen str
End Sub

Public Sub Terminate()
  KillThread
  Set Messages = Nothing
End Sub

