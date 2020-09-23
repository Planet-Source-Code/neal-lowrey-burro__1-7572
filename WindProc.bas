Attribute VB_Name = "WindProc"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
  (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_WNDPROC = (-4)

Public intSocket As Integer
Public OldWndProc As Long
Public IPDot As String

' Root value for hidden window caption
Public Const PROC_CAPTION = "ApartmentDemoProcessWindow"

Public Const ERR_InternalStartup = &H600
Public Const ERR_NoAutomation = &H601

Public Const ENUM_STOP = 0
Public Const ENUM_CONTINUE = 1

Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
   (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function GetWindowThreadProcessId Lib "user32" _
   (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Declare Function EnumThreadWindows Lib "user32" _
   (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) _
   As Long

Private mhwndVB As Long
' Window handle retrieved by EnumThreadWindows.
Private mfrmProcess As New frmProcess
' Hidden form used to id main thread.
Private mlngProcessID As Long
' Process ID.

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private MainApp As MainApp
Private Thread As Balk

Private mlngTimerID As Long

Sub Main()
  Dim ma As MainApp

  ' Borrow a window handle to use to obtain the process
  '   ID (see EnumThreadWndMain call-back, below).
  Call EnumThreadWindows(App.ThreadID, AddressOf EnumThreadWndMain, 0&)
  If mhwndVB = 0 Then
    Err.Raise ERR_InternalStartup + vbObjectError, , _
             "Internal error starting thread"
  Else
    GetWindowThreadProcessId mhwndVB, mlngProcessID
    ' The process ID makes the hidden window caption unique.
    If 0 = FindWindow(vbNullString, PROC_CAPTION & CStr(mlngProcessID)) Then
      ' The window wasn't found, so this is the first thread.
      If App.StartMode = vbSModeStandalone Then
        ' Create hidden form with unique caption.
        mfrmProcess.Caption = PROC_CAPTION & CStr(mlngProcessID)
        ' The Initialize event of MainApp (Instancing =
        '   PublicNotCreatable) shows the main user interface.
        Set ma = New MainApp
        ' (Application shutdown is simpler if there is no
        '   global reference to MainApp; instead, MainApp
        '   should pass Me to the main user form, so that
        '   the form keeps MainApp from terminating.)
      Else
        Err.Raise ERR_NoAutomation + vbObjectError, , _
             "Application can't be started with Automation"
      End If
    End If
  End If
End Sub

Public Sub SetThread(lThread As Balk)
  Set Thread = lThread
End Sub

' Call-back function used by EnumThreadWindows.
Public Function EnumThreadWndMain(ByVal hWnd As Long, ByVal _
                                  lParam As Long) As Long
  ' Save the window handle.
  mhwndVB = hWnd
  ' The first window is the only one required.
  ' Stop the iteration as soon as a window has been found.
  EnumThreadWndMain = ENUM_STOP
End Function

' MainApp calls this Sub in its Terminate event;
'   otherwise the hidden form will keep the
'   application from closing.
Public Sub FreeProcessWindow()
  SetWindowLong mhwndVB, GWL_WNDPROC, OldWndProc
  vbWSACleanup
  Unload mfrmProcess
  Set mfrmProcess = Nothing
End Sub

Public Sub FTP_Init(lMainApp As MainApp)
  Dim i As Integer
  Dim hdr As String, item As String
  
  '--- Initialization
  'an FTP command is terminated by Carriage_Return & Line_Feed
  'possible sintax errors in FTP commands
  sintax_error_list(0) = "200 Command Ok."
  sintax_error_list(1) = "202 Command not implemented, superfluous at this site."
  sintax_error_list(2) = "500 Sintax error, command unrecognized."
  sintax_error_list(3) = "501 Sintax error in parameters or arguments."
  sintax_error_list(4) = "502 Command not implemented."
  sintax_error_list(6) = "504 Command not implemented for that parameter."
  'initializes the list which contains the names,
  'passwords, access rights and default directory
  'recognized by the server
  If LoadProfile(App.Path & "\Burro.ini") Then
    '
  Else
    'frmFTP.StatusBar.Panels(1) = "Error Loading Ini File!"
  End If
  
  'initializes the records which contain the
  'informations on the connected users
  For i = 1 To MAX_N_USERS
    users(i).list_index = 0
 '   users(i).control_slot = INVALID_SLOT
 '   users(i).data_slot = INVALID_SLOT
    users(i).IP_Address = ""
    users(i).Port = 0
    users(i).data_representation = "A"
    users(i).data_format_ctrls = "N"
    users(i).data_structure = "F"
    users(i).data_tx_mode = "S"
    users(i).cur_dir = ""
    users(i).State = Log_In_Out '0
    users(i).full = False
  Next
 
  OldWndProc = SetWindowLong(mhwndVB, GWL_WNDPROC, AddressOf WindowProc)
  
  Set MainApp = lMainApp
 
  vbWSAStartup
  
  'begins SERVER mode on port 21
  ServerSlot = ListenForConnect(21, mhwndVB)
  
  If ServerSlot > 0 Then
   ' frmFTP.StatusBar.Panels(1) = Description
  Else
  '  frmFTP.StatusBar.Panels(1) = "Error Creating Listening Socket"
  End If
End Sub

Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                            ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim retf As Long
  Dim SendBuffer As String, msg$
  Dim lenBuffer As Integer 'send-buffer lenght
  Dim RecvBuffer As String
  Dim BytesRead As Integer 'receive-buffer lenght
  Dim i As Integer, GoAhead As Boolean
  Dim fixstr As String * 1024
  Dim lct As String
  Dim lcv As Integer
  Dim WSAEvent As Long
  Dim WSAError As Long
  Dim Valid_Slot As Boolean
  
  Valid_Slot = False
  GoAhead = True
  
  Select Case uMsg
  Case 5150
    
    'ServerLog "NOTIFICATION - " & wParam & " - " & lParam & "> " & Format$(Date$, "dd/mm/yy ") & Format$(Time$, "hh:mm - ")
    MainApp.SvrLogToScreen "NOTIFICATION - " & wParam & " - " & lParam & "> " & Format$(Date$, "dd/mm/yy ") & Format$(Time$, "hh:mm - ")
    For i = 1 To MAX_N_USERS       'registers the slot number in the first free user record
      If wParam = users(i).control_slot And users(i).full Then
        Valid_Slot = True
        Exit For
      End If
    Next
    If (wParam = ServerSlot) Or (wParam = NewSlot) Or Valid_Slot Then 'event on server slot
   '   frmFTP.StatusBar.Panels(1) = CStr(wParam)
      WSAEvent = WSAGetSelectEvent(lParam)
      WSAError = WSAGetAsyncError(lParam)
      'Debug.Print "Retf = "; WSAEvent; WSAError
      Select Case WSAEvent
        'FD_READ    = &H1    = 1
        'FD_WRITE   = &H2    = 2
        'FD_OOB     = &H4    = 4
        'FD_ACCEPT  = &H8    = 8
        'FD_CONNECT = &H10   = 16
        'FD_CLOSE   = &H20   = 32
      Case FD_CONNECT
        Debug.Print "FD_Connect " & wParam; lParam
   '     retf = getpeername(NewSlot, SockAddr, SockAddr_Size)
   '     Debug.Print "Peername = " & retf
   '     Debug.Print "IPAddr1 =" & SockAddr.sin_addr
   '     Debug.Print "IPPort1 =" & SockAddr.sin_port
      Case FD_ACCEPT
        Debug.Print "Doing FD_Accept"

        SockAddr.sin_family = AF_INET
        SockAddr.sin_port = 0
        'SockAddr.sin_addr = 0
        NewSlot = accept(ServerSlot, SockAddr, SockAddr_Size) 'try to accept new TCP connection
        If NewSlot = INVALID_SOCKET Then
          msg$ = "Can't accept new socket."
      '    frmFTP.StatusBar.Panels(1) = msg$ & CStr(NewSlot)
 
        Else
          Debug.Print "NewSlot OK "; NewSlot; num_users; MAX_N_USERS
   '       retf = getpeername(NewSlot, SockAddr, SockAddr_Size)
          IPDot = GetAscIP(SockAddr.sin_addr)
'Had to comment out the GetHostByAddress thing cause we don't do dns
      '    frmFTP.StatusBar.Panels(1) = IPDot & "<>" '& vbGetHostByAddress(IPDot)
          'Debug.Print "Peername = " & retf
          'Debug.Print "IPAddr2 =" & SockAddr.sin_addr & " IPdot=" & IPDot
          'Debug.Print "IPPort2 =" & SockAddr.sin_port & " Port:" & ntohs(SockAddr.sin_port)
          If num_users >= MAX_N_USERS Then        'new service request
            'the number of users exceeds the maximum allowed
            SendBuffer = "421 Service not available at this time, closing control connection." & vbCrLf
            lenBuffer = Len(SendBuffer)
            retf = send(NewSlot, SendBuffer, lenBuffer, 0)
            retf = closesocket(NewSlot)           'close connection
          Else
            SendBuffer = "220-Welcome to my demo Server v0.0.1!" & vbCrLf _
                       & "220 This program is written in VB 5.0" & vbCrLf
            lenBuffer = Len(SendBuffer)
            retf = send(NewSlot, SendBuffer, lenBuffer, 0)          'send welcome message
            Debug.Print "Send = " & retf
            num_users = num_users + 1      'increases the number of connected users
            For i = 1 To MAX_N_USERS       'registers the slot number in the first free user record
              If Not users(i).full Then
                users(i).control_slot = NewSlot
                users(i).full = True
                Exit For
              End If
            Next
          End If  'If num_users
        End If  'If NewSlot
      Case FD_READ
        Debug.Print "Doing FD_Read"
        BytesRead = recv(wParam, fixstr, 1024, 0) 'store read bytes in RecvBuffer
        RecvBuffer = Left$(fixstr, BytesRead)

        If InStr(RecvBuffer, vbCrLf) > 0 Then     'if received string is a command then executes it
          For i = 1 To MAX_N_USERS                'event on control slots
            If (wParam = users(i).control_slot) Then
              retf = FTP_Cmd(i, RecvBuffer)          'tr
              Exit For
            End If
          Next
        End If
      Case FD_CLOSE
        Debug.Print "Doing FD_Close"
        For i = 1 To MAX_N_USERS  'event on control slots
          If (wParam = users(i).control_slot) Then
            retf = closesocket(wParam)        'connection closed by client
            users(i).control_slot = INVALID_SOCKET        'frees the user record
            
            Set users(i).Jenny = Nothing
            users(i).full = False
            'ServerLog "<" & Format$(i, "000") & "> " & Format$(Date$, "dd/mm/yy ") & Format$(Time$, "hh:mm") & " - Logged Off"
            MainApp.SvrLogToScreen "<" & Format$(i, "000") & "> " & Format$(Date$, "dd/mm/yy ") & Format$(Time$, "hh:mm") & " - Logged Off"
            num_users = num_users - 1
            Exit For
          ElseIf (wParam = users(i).data_slot) Then
            retf = closesocket(wParam)        'connection closed by client
            users(i).data_slot = INVALID_SOCKET   'reinitilizes data slot
            users(i).State = Service_Commands '  2
            Exit For
          End If
       Next
      Case FD_WRITE
        Debug.Print "Doing FD_Write"
        'enables sending
      End Select
    End If
    'Debug.Print GetWSAErrorString(WSAGetLastError)
    MainApp.UsrCnt num_users
  End Select
  retf = CallWindowProc(OldWndProc, hWnd, uMsg, wParam, ByVal lParam)
  WindowProc = retf
End Function

Public Function FTP_Cmd(ID_User As Integer, cmd As String) As Integer
  
  Dim Kwrd As String 'keyword
  Dim argument(5) As String 'arguments
  Dim ArgN As Long
  Dim FTP_Err As Integer 'error
  Dim PathName As String, Drv As String
  
  Dim Full_Name As String 'pathname & file name
  Dim File_Len As Long 'file lenght in bytes
  Dim i As Long
  
  Dim Ok As Integer
  Dim Buffer As String
  Dim DummyS As String
  
  'variables used during the data exchange
  Dim ExecSlot As Integer
  Dim NewSockAddr As SockAddr
  
  On Error Resume Next 'routine for error interception
  
  FTP_Err = sintax_ctrl(cmd, Kwrd, argument())
  'log commands
  'ServerLog "<" & Format$(ID_User, "000") & "> " & Format$(Date$, "dd/mm/yy ") & Format$(Time$, "hh:mm - ") & cmd
  MainApp.SvrLogToScreen "<" & Format$(ID_User, "000") & "> " & Format$(Date$, "dd/mm/yy ") & Format$(Time$, "hh:mm - ") & cmd
  If FTP_Err <> 0 Then
    retf = send_reply(sintax_error_list(FTP_Err), ID_User)
    Exit Function
  End If
  
  Select Case UCase$(Kwrd)
  Case "USER"  'USER <username>
    Ok = False
    Debug.Print N_RECOGNIZED_USERS;
    For i = 1 To N_RECOGNIZED_USERS
      'Debug.Print UserIDs.No(i).Name
      'controls if the user is in the list of known users
      If argument(0) = UserIDs.No(i).Name Then
        'the user must enter a password but anonymous users can be accepted
        If UserIDs.No(i).Name = "anonymous" Then
          retf = send_reply("331 User anonymous accepted, please type your e-mail address as password.", ID_User)
        Else
          retf = send_reply("331 User name Ok, type in your password.", ID_User)
        End If
        users(ID_User).list_index = i
        users(ID_User).cur_dir = UserIDs.No(i).Home
        users(ID_User).State = Transfer_Parameters ' 1
        Ok = True
        Exit For
      End If
    Next
    If Not Ok Then  'unknown user
      retf = send_reply("530 Not logged in, user " & argument(0) & " is unknown.", ID_User)
      retf = logoff(ID_User)
    End If
  
  Case "PASS" 'PASS <password>
    If users(ID_User).State = Transfer_Parameters Then '1
      If LCase(UserIDs.No(users(ID_User).list_index).Name) = "anonymous" Then
        'anonymous user
        retf = send_reply("230 User anonymous logged in, proceed.", ID_User)
        users(ID_User).State = Service_Commands ' 2
        Set users(ID_User).Jenny = CreateObject("Burro.Balk")
        users(ID_User).Jenny.SetUserData users(ID_User)
        users(ID_User).Jenny.SetUserPermissions UserIDs.No(users(ID_User).list_index), users(ID_User).list_index
        users(ID_User).Jenny.SetCallBack MainApp
      Else
        If argument(0) = UserIDs.No(users(ID_User).list_index).Pass Then
          'correct password, the user can proceed
          retf = send_reply("230 User logged in, proceed.", ID_User)
          users(ID_User).State = Service_Commands ' 2
          Set users(ID_User).Jenny = CreateObject("Burro.Balk")
          users(ID_User).Jenny.SetUserData users(ID_User)
          users(ID_User).Jenny.SetUserPermissions UserIDs.No(users(ID_User).list_index), users(ID_User).list_index
          users(ID_User).Jenny.SetCallBack MainApp
        Else
          'wrong password, the user is disconnected
          retf = send_reply("530 Not logged in, wrong password.", ID_User)
          retf = logoff(ID_User)
        End If
      End If
    Else
      'the user must enter his name
      retf = send_reply("503 I need your username.", ID_User)
    End If
  Case "QUIT": 'QUIT
    retf = logoff(ID_User)
  Case Else
'MainApp.SvrLogToScreen "Ftp Command Fired"
    users(ID_User).Jenny.New_Cmd Kwrd, argument()
  End Select

End Function

Public Function FTP_Cmd2() As Integer
 
  Dim ArgN As Long
  Dim PathName As String, Drv As String
  
  Dim i As Long
  
  Dim Ok As Integer
  Dim DummyS As String
  
  'variables used during the data exchange
  Dim ExecSlot As Integer
  Dim NewSockAddr As SockAddr
  
  Dim Full_Name As String
  Dim data_representation As String * 1
  Dim open_file As Integer
  Dim retr_stor As Integer  '0=RETR; 1=STOR
  Dim Buffer As String  'contains data to send
  Dim File_Len As Long  '--- Binary mode only
  Dim blocks As Long  'number of 1024 bytes blocks in file
  Dim spare_bytes As Long
  Dim next_block As Long  'next block to send
  Dim next_byte As Long  'points to position in file of the next block to send
  Dim try_again As Integer  'if try_again=true the old line is sent =Ascii mode only
  Dim Dummy As String
  
  Dim DirFnd As Boolean
  Dim error_on_data_cnt As Boolean
  Dim close_data_cnt As Boolean
  
  On Error Resume Next 'routine for error interception
  
  Select Case UCase$(FTP_Command)
  Case "CWD", "XCWD" 'CWD <pathname>
    If users(FTP_Index).State = 2 Then
      
      PathName = ChkPath(FTP_Index, FTP_Args(0))
      Drv = Left(PathName, 2)
      
      '#######################################tr####################
      'controls access rights
      DirFnd = False
      For i = 1 To UserIDs.No(users(FTP_Index).list_index).Pcnt
        If UserIDs.No(users(FTP_Index).list_index).Priv(i).Path = PathName Then
        'To do drive letter permissions use this line
        'If Left(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path, 2) = Drv Then
          DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(i).Accs
          DirFnd = True
          Exit For
        End If
      Next

      If InStr(DummyS, "L") And DirFnd Then
      
      '######################################end tr#####################
         ChDrive Drv
         ChDir PathName
         If Err.Number = 0 Then
           users(FTP_Index).cur_dir = CurDir
           'existing directory
           retf = send_reply("250 CWD command executed.", FTP_Index)
         ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
           'no existing directory
           retf = send_reply("550 CWD command not executed: " & Error$, FTP_Index)
         Else
      '     frmFTP.StatusBar.Panels(1) = "Error " & CStr(Err) & " occurred."
           retf = logoff(FTP_Index)
           'End
         End If
      '#######################################tr####################
      Else
        retf = send_reply("550 CWD command not executed: User does not have permissions", FTP_Index)
      End If
      '#######################################end tr####################
    Else
      'user not logged in
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  
  Case "CDUP", "XCUP": 'CDUP
    If users(FTP_Index).State = 2 Then
      ChDir users(FTP_Index).cur_dir
      ChDir ".."
      users(FTP_Index).cur_dir = CurDir
      retf = send_reply("200 CDUP command executed.", FTP_Index)
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  Case "PORT" 'PORT <host-port>
    If users(FTP_Index).State = Service_Commands Then    ' 2
      'opens a data connection
      ExecSlot = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
      If ExecSlot < 0 Then
        'error
        retf = send_reply("425 Can't build data connection.", FTP_Index)
      Else
        NewSockAddr.sin_family = PF_INET
        'remote IP address
        IPLong.Byte4 = Val(FTP_Args(0))
        IPLong.Byte3 = Val(FTP_Args(1))
        IPLong.Byte2 = Val(FTP_Args(2))
        IPLong.Byte1 = Val(FTP_Args(3))
        CopyMemory i, IPLong, 4
        NewSockAddr.sin_addr = i

        'remote port
        ArgN = Val(FTP_Args(4))
        NewSockAddr.sin_port = htons(ArgN)
        retf = connect(ExecSlot, NewSockAddr, 16)
        If retf < 0 Then
          retf = send_reply("425 Can't build data connection.", FTP_Index)
        Else
          retf = send_reply("200 PORT command executed.", FTP_Index)
          'stores the IP-address and port number in user record
          users(FTP_Index).data_slot = ExecSlot
          users(FTP_Index).IP_Address = FTP_Args(0) & "." & FTP_Args(1) & "." & _
                                        FTP_Args(2) & "." & FTP_Args(3)
          users(FTP_Index).Port = Val(FTP_Args(4))
          'ServerLog ("IP=" & users(FTP_Index).IP_Address & ":" & FTP_Args(4))
          Thread.SendMessage "IP=" & users(FTP_Index).IP_Address & ":" & FTP_Args(4)
'          '<state> field establishes that now is
'          'possible to exec commands requiring a data connection
          users(FTP_Index).State = 3
          Debug.Print "data "; ExecSlot
          Debug.Print "ctrl "; users(FTP_Index).control_slot
        End If
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
'
  
  Case "TYPE" 'TYPE <type-code>
    If users(FTP_Index).State = 2 Then
      'stores the access parameters in user record
      retf = send_reply("200 TYPE command executed.", FTP_Index)
      users(FTP_Index).data_representation = FTP_Args(0)
      users(FTP_Index).data_format_ctrls = FTP_Args(1)
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  
  Case "STRU" 'STRU <structure-code>
    If users(FTP_Index).State = 2 Then
      'stores access parameters in the user record
      retf = send_reply("200 STRU command executed.", FTP_Index)
      users(FTP_Index).data_structure = FTP_Args(0)
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
    
  Case "MODE" 'MODE <mode-code>
    If users(FTP_Index).State = 2 Then
      'stores access parameters in the user record
      retf = send_reply("200 MODE command executed.", FTP_Index)
      users(FTP_Index).data_tx_mode = FTP_Args(0)
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  
  Case "RETR" 'RETR <pathname>
    On Error GoTo FileError
    If users(FTP_Index).State = 3 Then
      Dim Counter As Integer
      Full_Name = ChkPath(FTP_Index, FTP_Args(0))
        'file exist?
      i = FileLen(Full_Name)
      If Err.Number = 0 Then 'Yes
          'controls access rights
        'DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(1).Accs
        'If InStr(DummyS, "R") Then
        DirFnd = False
        PathName = LCase$(Left(Full_Name, InStrRev(Full_Name, "\")))
        For i = 1 To UserIDs.No(users(FTP_Index).list_index).Pcnt
          If LCase$(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path) = PathName Then
          'To do drive letter permissions use this line
          'If Left(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path, 2) = Drv Then
            DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(i).Accs
            DirFnd = True
            Exit For
          End If
        Next
  
        If InStr(DummyS, "R") And DirFnd Then
          retf = open_data_connect(FTP_Index)
          
          If Not open_file Then
            Open Full_Name For Binary Access Read Lock Write As #FTP_Index
            open_file = True
          End If
          Do
            If users(FTP_Index).data_representation = "A" Then
              If try_again Then
              Else      're-send old line
                Line Input #FTP_Index, Buffer
              End If
              retf = send_data(Buffer & vbCrLf, FTP_Index)
              If retf < 0 Then 'SOCKET_ERROR
                retf = WSAGetLastError()
                If retf = WSAEWOULDBLOCK Then
                  try_again = True
                Else        'error on sending
                  error_on_data_cnt = True
                  close_data_cnt = True
                End If
              Else
                try_again = False
              End If
              If EOF(FTP_Index) Then close_data_cnt = True
            Else  'binary transfer
              'sends file on data connection; data are sent in blocks of 1024 bytes
              If next_block = 0 Then
                File_Len = LOF(FTP_Index)
                blocks = Int(File_Len / 1024)    '# of blocks
                spare_bytes = File_Len Mod 1024  '# of remaining bytes
                Buffer = String$(1024, " ")
              End If
              If next_block < blocks Then 'sends blocks
                Get #FTP_Index, next_byte + 1, Buffer
                retf = send_data(Buffer, FTP_Index)
                If retf < 0 Then
                  retf = WSAGetLastError()
                  If retf = WSAEWOULDBLOCK Then  'try again
                  Else
                    error_on_data_cnt = True
                    close_data_cnt = True
                  End If
                Else   'next block
                  next_block = next_block + 1
                  next_byte = next_byte + 1024
                End If
              Else    'sends remaining bytes
                Buffer = String$(spare_bytes, " ")
                Get #FTP_Index, , Buffer
                retf = send_data(Buffer, FTP_Index)
                close_data_cnt = True
              End If
            End If
          Loop Until close_data_cnt
          If close_data_cnt Then  're-initialize files_info record
          '  files_info(index).open_file = False
          '  files_info(index).next_block = 0  'blocks count
          '  files_info(index).next_byte = 0   'pointer to next block
          '  files_info(index).try_again = False
            
            Close #FTP_Index    'close file
            If error_on_data_cnt Then    'replies to user
              retf = send_reply("550 RETR command not executed.", FTP_Index)
            Else
              retf = send_reply("226 RETR command completed.", FTP_Index)
            End If
            retf = close_data_connect(FTP_Index)    'close data connection
          End If
        Else
            'the user can't retrieves files
          retf = send_reply("550 You can't take this file action.", FTP_Index)
          retf = close_data_connect(FTP_Index)
        End If
      ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
        'no existing file
        retf = send_reply("550 RETR command not executed: " & Error$, FTP_Index)
        retf = close_data_connect(FTP_Index)
      Else
        frmFTP.StatusBar.Panels(1) = "Error " & Err.Number & " occurred."
        retf = close_data_connect(FTP_Index)
        retf = logoff(FTP_Index)
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
'MsgBox App.ThreadID & " done his retr duty as " & users(FTP_Index).data_representation
  Case "STOR" 'STOR <pathname>
    If users(FTP_Index).State = 3 Then
      Full_Name = ChkPath(FTP_Index, FTP_Args(0))
      'controls access rights
'      DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(1).Accs
      
      DirFnd = False
      PathName = LCase$(Left(Full_Name, InStrRev(Full_Name, "\")))
      For i = 1 To UserIDs.No(users(FTP_Index).list_index).Pcnt
        If LCase$(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path) = PathName Then
        'To do drive letter permissions use this line
        'If Left(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path, 2) = Drv Then
          DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(i).Accs
          DirFnd = True
          Exit For
        End If
      Next
  
      If InStr(DummyS, "W") And DirFnd Then
        If Not open_file Then
          Open Full_Name For Binary Access Write Lock Read Write As #FTP_Index
          open_file = True
        End If
        retf = open_data_connect(FTP_Index)
        Do
          If users(FTP_Index).data_representation = "A" Then
            retf = receive_data(Buffer, FTP_Index)
            If retf < 0 Then   'SOCKET_ERROR
              retf = WSAGetLastError()
              If retf = WSAEWOULDBLOCK Then   'try_again
              Else       'error on receiving
                error_on_data_cnt = True
                close_data_cnt = True
              End If
            ElseIf retf = 0 Then  'connection closed by peer
              close_data_cnt = True
            Else 'retf > 0  write on file
              Dummy$ = Left$(Buffer, retf)
              Print #FTP_Index, Dummy$
            End If
          Else  'binary transfer
            retf = receive_data(Buffer, FTP_Index)
            If retf < 0 Then
              retf = WSAGetLastError()
              If retf = WSAEWOULDBLOCK Then  'try again
              Else
                error_on_data_cnt = True
                close_data_cnt = True
              End If
            ElseIf retf = 0 Then     'connection closed by peer
              close_data_cnt = True
            Else
              Dummy$ = Left$(Buffer, retf)
              Put #FTP_Index, , Dummy$
            End If
          End If
        Loop Until close_data_cnt
        If close_data_cnt Then   're-initialize files_info record
          'files_info(Index).open_file = False
          'files_info(Index).next_block = 0 'blocks count
          'files_info(Index).next_byte = 0  'pointer to next block
          'files_info(Index).try_again = False
          Close #FTP_Index    'close file
          If error_on_data_cnt Then    'replies to user
            retf = send_reply("550 STOR command not executed.", FTP_Index)
          Else
            retf = send_reply("226 STOR command completed.", FTP_Index)
          End If
          retf = close_data_connect(FTP_Index)     'closes data connection
          
        End If
      Else
        'the user can't stores files
        retf = send_reply("550 You can't take this file action.", FTP_Index)
        retf = close_data_connect(FTP_Index)
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
MsgBox App.ThreadID & " done his stor duty as " & users(FTP_Index).data_representation
  Case "RNFR"  'RNFR <pathname>
    If users(FTP_Index).State = 2 Then
      Full_Name = ChkPath(FTP_Index, FTP_Args(0))
      'file exists?
      i = FileLen(Full_Name)
      If Err.Number = 0 Then 'Yes
        'controls access rights
        DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(1).Accs
        If InStr(DummyS, "M") Then
          'The user can updates files.
          'The name of file to rename is temporarily stored in the user record.
          users(FTP_Index).temp_data = Full_Name
          'next command must be a RNTO
          users(FTP_Index).State = 6
          retf = send_reply("350 ReName command expect further information.", FTP_Index)
        Else
          'the user can't writes on files
          retf = send_reply("550 You can't take this file action.", FTP_Index)
        End If
      ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
        'no existing file
        retf = send_reply("550 RNFR command not executed: " & Error$, FTP_Index)
      Else
   '     frmFTP.StatusBar.Panels(1) = "Error " & Err.Number & " occurred."
        retf = logoff(FTP_Index)
        'End
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  
  Case "RNTO"  'RNTO <pathname>
    If users(FTP_Index).State = 6 Then
      Full_Name = ChkPath(FTP_Index, FTP_Args(0))
      Name users(FTP_Index).temp_data As Full_Name
      If Err.Number = 0 Then
        users(FTP_Index).State = 2
        'file exists
        retf = send_reply("350 ReName command executed.", FTP_Index)
      ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
        'no existing file
        retf = send_reply("550 RNTO command not executed: " & Error$, FTP_Index)
      Else
  '      frmFTP.StatusBar.Panels(1) = "Error " & Err.Number & " occurred."
        retf = logoff(FTP_Index)
        'End
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
    
  Case "DELE"  'DELE <pathname>
    If users(FTP_Index).State = 2 Then
      Full_Name = ChkPath(FTP_Index, FTP_Args(0))
      'controls access rights
      'DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(1).Accs
      'If InStr(DummyS, "K") Then
      DirFnd = False
      PathName = Left(Full_Name, InStrRev(Full_Name, "\"))
      For i = 1 To UserIDs.No(users(FTP_Index).list_index).Pcnt
        If UserIDs.No(users(FTP_Index).list_index).Priv(i).Path = PathName Then
        'To do drive letter permissions use this line
        'If Left(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path, 2) = Drv Then
          DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(i).Accs
          DirFnd = True
          Exit For
        End If
      Next
  
      If InStr(DummyS, "K") And DirFnd Then
        'the user can updates files
        Kill Full_Name
        If Err.Number = 0 Then
          'file exists
          retf = send_reply("250 DELE command executed.", FTP_Index)
        ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
          'file no exists
          retf = send_reply("550 DELE command not executed: " & Error$, FTP_Index)
        Else
    '      frmFTP.StatusBar.Panels(1) = "Error " & Err.Number & " occurred."
          retf = logoff(FTP_Index)
          'End
        End If
      Else
        'the user can't delete files
        retf = send_reply("550 You can't take this file action.", FTP_Index)
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
    
  Case "RMD", "XRMD" 'RMD <pathname>
    If users(FTP_Index).State = 2 Then
      PathName = ChkPath(FTP_Index, FTP_Args(0))
      'controls access rights
      'DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(1).Accs
      'If InStr(DummyS, "D") Then
      DirFnd = False
      For i = 1 To UserIDs.No(users(FTP_Index).list_index).Pcnt
        If UserIDs.No(users(FTP_Index).list_index).Priv(i).Path = PathName Then
        'To do drive letter permissions use this line
        'If Left(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path, 2) = Drv Then
          DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(i).Accs
          DirFnd = True
          Exit For
        End If
      Next
  
      If InStr(DummyS, "K") And DirFnd Then
        'the user can updates files
        Kill PathName & "\*.*"
        If Err.Number = 53 Or Err.Number = 708 Then Err.Number = 0 'empty directory
        RmDir PathName
        If Err.Number = 0 Then
          'directory exists
          retf = send_reply("250 RMD command executed.", FTP_Index)
        ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
          'directory no exists
          retf = send_reply("550 RMD command not executed: " & Error$, FTP_Index)
        Else
   '       frmFTP.StatusBar.Panels(1) = "Error " & Err.Number & " occurred."
          retf = logoff(FTP_Index)
          'End
        End If
      Else
        'the user can't delete files
        retf = send_reply("550 You can't take this file action.", FTP_Index)
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  
  Case "MKD", "XMKD" 'MKD <pathname>
    If users(FTP_Index).State = 2 Then
      PathName = ChkPath(FTP_Index, FTP_Args(0))
      'controls access rights
      'DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(1).Accs
      'If InStr(DummyS, "M") Then
      DirFnd = False
      For i = 1 To UserIDs.No(users(FTP_Index).list_index).Pcnt
        If UserIDs.No(users(FTP_Index).list_index).Priv(i).Path = PathName Then
        'To do drive letter permissions use this line
        'If Left(UserIDs.No(users(FTP_Index).list_index).Priv(i).Path, 2) = Drv Then
          DummyS = UserIDs.No(users(FTP_Index).list_index).Priv(i).Accs
          DirFnd = True
          Exit For
        End If
      Next
  
      If InStr(DummyS, "M") And DirFnd Then
        'the user can updates files
        MkDir PathName
        If Err.Number = 0 Then
          'the directory is been created
          retf = send_reply("257 " & FTP_Args(0) & " created.", FTP_Index)
        ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
          'the directory isn't been created
          retf = send_reply("550 MKD command not executed: " & Error$, FTP_Index)
        Else
     '     frmFTP.StatusBar.Panels(1) = "Error " & Err.Number & " occurred."
          retf = logoff(FTP_Index)
          'End
        End If
      Else
        'the user can't write on files
        retf = send_reply("550 You can't take this file action.", FTP_Index)
      End If
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  
  Case "PWD", "XPWD" 'PWD
    If users(FTP_Index).State = 2 Then
      PathName = users(FTP_Index).cur_dir
      'Who doesn't want to know the the drive they are on?
      'PathName = Right$(PathName, Len(PathName) - 2)
      retf = send_reply("257 """ & PathName & """ is the current directory.", FTP_Index)
    Else
      retf = send_reply("530 User not logged in.", FTP_Index)
    End If
  
  Case "LIST", "NLST"   'LIST <pathname>Or InStr(FTP_Args(0), "-L")
      LIST_NLST FTP_Index, FTP_Command, FTP_Args(0)
    
  Case "STAT"  'STAT <pathname>
      retf = send_reply("200 Not Implemented..", FTP_Index)
  Case "HELP"  'HELP <string>
    DummyS = "214-This is the list of recognized FTP commands:"
    retf = send_reply(DummyS, FTP_Index)
      DummyS = "214-   USER  PASS  CWD   XCWD  CDUP  XCUP  QUIT  PORT" & vbCrLf _
             & "214-   PASV  TYPE  STRU  MODE  RETR  STOR  RNFR  RNTO" & vbCrLf _
             & "214-   DELE  RMD   XRMD  MKD   XMKD  PWD   XPWD" & vbCrLf _
             & "214    LIST  NLST  SYST  STAT  HELP  NOOP"
    retf = send_reply(DummyS, FTP_Index)
  
  Case "NOOP" 'NOOP
    retf = send_reply("200 NOOP command executed.", FTP_Index)
  Case ""
    Thread.SendMessage "error with ftpCommand"
  Case Else
    retf = send_reply("200 Not Implemented.." & FTP_Command, FTP_Index)
  End Select
Exit Function
FileError:
  Close #FTP_Index    'close file
  retf = send_reply("550 RETR command not executed. File Error", FTP_Index)
  retf = close_data_connect(FTP_Index)    'close data connection
End Function

Public Sub StartTimer()
  mlngTimerID = SetTimer(0, 0, 100, AddressOf TimerProc)
End Sub

Private Sub TimerProc(ByVal hWnd As Long, ByVal msg As Long, _
                      ByVal idEvent As Long, ByVal curTime As Long)
'Thread.SendMessage "Timer Fired"
  StopTimer
  FTP_Cmd2
End Sub

Public Sub StopTimer()
  If mlngTimerID > 0 Then
    KillTimer 0, mlngTimerID
    mlngTimerID = 0
  End If
End Sub

Public Sub KillThread()
  Set Thread = Nothing
End Sub
