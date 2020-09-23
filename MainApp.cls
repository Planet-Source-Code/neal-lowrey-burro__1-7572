VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eState
  Log_In_Out = 0
  Transfer_Parameters = 1
  Service_Commands = 2
  Busy = 3
End Enum

Public Type Privtyp
  Path As String
  Accs As String '[R]ead,[W]rite,[D]elete,e[X]ecute > Files
                 '[L]ist,[M]ake,[K]ill,[S]ubs       > Dirs
End Type

Public Type UserInfo
  Name As String 'list of the users which can access to server file-system
  Pass As String 'list of passwords of each user which can access to server file-system
  Pcnt As Integer
  Priv(20) As Privtyp
  Home As String 'default directory of each user
End Type
 
Public Type User
  'indexes user name and password inside
  '<usernames_list> and <passwords_list> arrays
  list_index As Integer
  control_slot As Long 'slot number used by client to send commands to server. On this slot also travel the replies of server.
  data_slot As Long    'slot number used by server to send data to client
  IP_Address As String 'client IP address
  Port As Integer      'number of client data port
  'representation type to use for data
  'encoding (ex. ASCII o EBCDIC), default type is A (= ASCII)
  data_representation As String
  'type of vertical format control
  '(ie. line-feed, form-feed), default
  'value is N (No print = no command)
  data_format_ctrls As String
  'file structure (ex. File-structure, Record-structure),
  'default value is F (File-structure = no internal
  'structure)
  data_structure As String
  'indicates if the data are processing before
  'transmission (ex. Stream-mode, Block-mode),
  'default value is S (Stream-mode = no
  'processing)
  data_tx_mode As String
  'current working directory
  cur_dir As String
  'user state:
  'in state 0 the user sends access control commands;
  'in state 1 the user sends transfer parameter commands;
  'in state 2 the user sends FTP service commands.
  State As eState
  full As Integer  'if true the record is already in use
  temp_data As String  'contains temporary data
  Jenny As Balk
End Type

Private frmFTP As New frmFTP


Private Sub Class_Initialize()
   Set frmFTP.MainApp = Me
   frmFTP.Caption = frmFTP.Caption
   frmFTP.Show
   FTP_Init Me
End Sub

Friend Sub Closing()
   Unload frmFTP
   Set frmFTP = Nothing
   Call FreeProcessWindow
End Sub

Public Sub SvrLogToScreen(ByVal str As String)
  frmFTP.LogWnd.AddItem str
  frmFTP.LogWnd.Selected(frmFTP.LogWnd.ListCount - 1) = True
End Sub

Public Sub UsrCnt(ByVal users As Integer)
  frmFTP.UsrCnt.Text = users
End Sub
