Attribute VB_Name = "VBSOCK"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Public Declare Function lstrlen Lib "kernel32" (ByVal lpString As Any) As Integer

Public DnsHost As String
Public MaxSockets As Integer
Public MaxUDP As Long
Public Description As String
Public Status As String
Public sintax_error_list(10) As String 'the list of the messages which signal a sintax error in a FTP command

Public users(MAX_N_USERS) As User

Public Type file_info
  Full_Name As String
  data_representation As String * 1
  open_file As Integer
  retr_stor As Integer  '0=RETR; 1=STOR
  Buffer As String  'contains data to send
  File_Len As Long  '--- Binary mode only
  blocks As Long  'number of 1024 bytes blocks in file
  spare_bytes As Long
  next_block As Long  'next block to send
  next_byte As Long  'points to position in file of the next block to send
  try_again As Integer  'if try_again=true the old line is sent =Ascii mode only
End Type

Public files_info(5) As file_info

'contains error during function call
Public retf As Integer

'*** Variables used during TCP/IP exchange
'slot number assigned to Server
Public ServerSlot As Long
'number of clients connected to server
Public num_users As Integer
Public ListenSock As Long
Public NewSlot As Long

'-------------------------------------
'used by jenny
 Public FTP_Index As Integer
 Public FTP_Command As String
 Public FTP_Args() As String

Function ConnectSocket(ByVal hWndtyp As Long, ByVal Host As String, ByVal Port As Integer) As Long
  Dim SockreadBuffer As String, RetIpPort As String
  Dim s As Long, Dummy As Long
  'Dim NewSock As SockAddr
  Dim SelectOps As Integer
    
    SockreadBuffer = ""
    SockAddr.sin_family = AF_INET
    SockAddr.sin_port = htons(Port)
    If Val(SockAddr.sin_zero) = INVALID_SOCKET Then
        ConnectSocket = INVALID_SOCKET
        Exit Function
    End If
    SockAddr.sin_addr = GetHostByNameAlias(Host)
    If SockAddr.sin_addr = INADDR_NONE Then
        ConnectSocket = INVALID_SOCKET
        Exit Function
    End If
    RetIpPort = GetAscIP(SockAddr.sin_addr) & ":" & ntohs(SockAddr.sin_port)
    Debug.Print RetIpPort
    s = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ConnectSocket = INVALID_SOCKET
        Exit Function
    End If
    'If SetSockLinger(S, 1, 0) = SOCKET_ERROR Then
    '    If S > 0 Then
    '       Dummy = closesocket(S)
    '    End If
    '    ConnectSocket = INVALID_SOCKET
    '    Exit Function
    'End If
    SelectOps = FD_CONNECT Or FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, hWndtyp, ByVal 5152, ByVal SelectOps) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ConnectSocket = INVALID_SOCKET
        Exit Function
    End If
    If connect(s, SockAddr, SockAddr_Size) <> -1 Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ConnectSocket = INVALID_SOCKET
        Exit Function
    End If
    ConnectSocket = s
End Function

Function WSAGetSelectEvent(ByVal lParam As Long) As Long
    WSAGetSelectEvent = Int(lParam Mod 65536)
End Function

'Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
'    If (lParam And &HFFFF&) > &H7FFF Then
'        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
'    Else
'        WSAGetSelectEvent = lParam And &HFFFF&
'    End If
'End Function

Public Function WSAGetAsyncError(ByVal lParam As Long) As Long
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

Function DNS_Lookup(ByVal dnsip As String) As String
  DnsHost = ""
  vbWSAStartup
  DoEvents
  DNS_Lookup = vbGetHostByAddress(dnsip)
  DoEvents
  vbWSACleanup
End Function

Function vbGetHostByAddress(ByVal sAddress As String) As String
  Dim lAddress As Long
  Dim PointerToMemoryLocation As Long
  Dim HostName As String
  Dim hostent As hostent
  lAddress = inet_addr(sAddress)
  PointerToMemoryLocation = gethostbyaddr(lAddress, 4, PF_INET)
  If PointerToMemoryLocation <> 0 Then
    CopyMemory hostent, ByVal PointerToMemoryLocation, Len(hostent)
    HostName = String(256, 0)
    CopyMemory ByVal HostName, ByVal hostent.h_name, 256
    If HostName = "" Then
      vbGetHostByAddress = "Unable to Resolve Address"
    Else
      vbGetHostByAddress = Left(HostName, InStr(HostName, Chr(0)) - 1)
    End If
  Else
    vbGetHostByAddress = "No DNS Entry"
  End If
End Function

Function LoByte(ByVal wParam As Integer)
  LoByte = wParam And &HFF&
End Function

Function HiByte(ByVal wParam As Integer)
  HiByte = wParam / &H100 And &HFF&
End Function

Sub vbWSAStartup()
  Dim iReturn As Integer
  Dim sHighByte As String
  Dim sLowByte As String
  Dim sMsg As String
  Dim i As Integer
  iReturn = WSAStartup(&H101, WSAdata)
   If LoByte(WSAdata.wVersion) < WS_VERSION_MAJOR Or _
    (LoByte(WSAdata.wVersion) = WS_VERSION_MAJOR _
    And HiByte(WSAdata.wVersion) < WS_VERSION_MINOR) Then
        sHighByte = Trim(str(HiByte(WSAdata.wVersion)))
        sLowByte = Trim(str(LoByte(WSAdata.wVersion)))
    End If
    If WSAdata.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(str$(MIN_SOCKETS_REQD)) & " supported sockets."
    End If
    MaxSockets = WSAdata.iMaxSockets
    If MaxSockets < 0 Then
        MaxSockets = 65536 + MaxSockets
    End If
    MaxUDP = WSAdata.iMaxUdpDg
    If MaxUDP < 0 Then
        MaxUDP = 65536 + MaxUDP
    End If
    Description = WSAdata.szDescription
    Status = ""
    Status = WSAdata.szSystemStatus
End Sub

Sub vbWSACleanup()
  Dim iReturn As Long
  Dim sMsg As String
  iReturn = WSACleanup()
  If iReturn <> 0 Then
    sMsg = "WSock32 Error - " & Trim$(str$(iReturn)) & " occurred in Cleanup"
  End If
End Sub


'returns IP as long, in network byte order
Public Function GetHostByNameAlias(ByVal HostName$) As Long
    'Return IP address as a long, in network byte order
    Dim phe&
    Dim heDestHost As hostent
    Dim addrList&
    Dim retIP&
    retIP = inet_addr(HostName$)
    If retIP = INADDR_NONE Then
        phe = gethostbyname(HostName$)
        If phe <> 0 Then
            CopyMemory heDestHost, ByVal phe, Len(heDestHost)
            CopyMemory addrList, ByVal heDestHost.h_addr_list, 4
            CopyMemory retIP, ByVal addrList, heDestHost.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
End Function

Public Function GetAscIP(ByVal inn As Long) As String
  Dim nStr&
  Dim lpStr&
  Dim retString$
  retString = String(32, 0)
  lpStr = inet_ntoa(inn)
  If lpStr Then
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    CopyMemory ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    GetAscIP = retString
  Else
    GetAscIP = "255.255.255.255"
  End If
End Function

Public Function SetSockLinger(ByVal SockNum As Long, ByVal OnOff As Integer, ByVal LingerTime As Integer) As Long
    Dim Linger As LingerType
    Linger.l_onoff = OnOff
    Linger.l_linger = LingerTime
    If setsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
        Debug.Print "Error setting linger info: " & WSAGetLastError()
        SetSockLinger = SOCKET_ERROR
    Else
        If getsockopt(SockNum, SOL_SOCKET, SO_LINGER, Linger, 4) Then
            Debug.Print "Error getting linger info: " & WSAGetLastError()
            SetSockLinger = SOCKET_ERROR
        Else
            Debug.Print "Linger is on if nonzero: "; Linger.l_onoff
            Debug.Print "Linger time if linger is on: "; Linger.l_linger
        End If
    End If
End Function

Public Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&) As Long
Dim s As Long, Dummy As Long
Dim SelectOps As Integer

    s = Socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    
    SockAddr.sin_family = AF_INET
    SockAddr.sin_port = htons(Port)
    If SockAddr.sin_port = INVALID_SOCKET Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    SockAddr.sin_addr = htonl(INADDR_ANY)
    If SockAddr.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If

    If bind(s, SockAddr, SockAddr_Size) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    
    If listen(s, 1) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    
    SelectOps = FD_CONNECT Or FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal 5150, ByVal SelectOps) Then
        If s > 0 Then
            Dummy = closesocket(s)
        End If
        ListenForConnect = SOCKET_ERROR
        Exit Function
    End If
    ListenForConnect = s
End Function

Function GetWSAErrorString(ByVal errnum As Long) As String
  On Error Resume Next
  Select Case errnum
  Case 10004: GetWSAErrorString = "Interrupted system call."
  Case 10009: GetWSAErrorString = "Bad file number."
  Case 10013: GetWSAErrorString = "Permission Denied."
  Case 10014: GetWSAErrorString = "Bad Address."
  Case 10022: GetWSAErrorString = "Invalid Argument."
  Case 10024: GetWSAErrorString = "Too many open files."
  Case 10035: GetWSAErrorString = "Operation would block."
  Case 10036: GetWSAErrorString = "Operation now in progress."
  Case 10037: GetWSAErrorString = "Operation already in progress."
  Case 10038: GetWSAErrorString = "Socket operation on nonsocket."
  Case 10039: GetWSAErrorString = "Destination address required."
  Case 10040: GetWSAErrorString = "Message too long."
  Case 10041: GetWSAErrorString = "Protocol wrong type for socket."
  Case 10042: GetWSAErrorString = "Protocol not available."
  Case 10043: GetWSAErrorString = "Protocol not supported."
  Case 10044: GetWSAErrorString = "Socket type not supported."
  Case 10045: GetWSAErrorString = "Operation not supported on socket."
  Case 10046: GetWSAErrorString = "Protocol family not supported."
  Case 10047: GetWSAErrorString = "Address family not supported by protocol family."
  Case 10048: GetWSAErrorString = "Address already in use."
  Case 10049: GetWSAErrorString = "Can't assign requested address."
  Case 10050: GetWSAErrorString = "Network is down."
  Case 10051: GetWSAErrorString = "Network is unreachable."
  Case 10052: GetWSAErrorString = "Network dropped connection."
  Case 10053: GetWSAErrorString = "Software caused connection abort."
  Case 10054: GetWSAErrorString = "Connection reset by peer."
  Case 10055: GetWSAErrorString = "No buffer space available."
  Case 10056: GetWSAErrorString = "Socket is already connected."
  Case 10057: GetWSAErrorString = "Socket is not connected."
  Case 10058: GetWSAErrorString = "Can't send after socket shutdown."
  Case 10059: GetWSAErrorString = "Too many references: can't splice."
  Case 10060: GetWSAErrorString = "Connection timed out."
  Case 10061: GetWSAErrorString = "Connection refused."
  Case 10062: GetWSAErrorString = "Too many levels of symbolic links."
  Case 10063: GetWSAErrorString = "File name too long."
  Case 10064: GetWSAErrorString = "Host is down."
  Case 10065: GetWSAErrorString = "No route to host."
  Case 10066: GetWSAErrorString = "Directory not empty."
  Case 10067: GetWSAErrorString = "Too many processes."
  Case 10068: GetWSAErrorString = "Too many users."
  Case 10069: GetWSAErrorString = "Disk quota exceeded."
  Case 10070: GetWSAErrorString = "Stale NFS file handle."
  Case 10071: GetWSAErrorString = "Too many levels of remote in path."
  Case 10091: GetWSAErrorString = "Network subsystem is unusable."
  Case 10092: GetWSAErrorString = "Winsock DLL cannot support this application."
  Case 10093: GetWSAErrorString = "Winsock not initialized."
  Case 10101: GetWSAErrorString = "Disconnect."
  Case 11001: GetWSAErrorString = "Host not found."
  Case 11002: GetWSAErrorString = "Nonauthoritative host not found."
  Case 11003: GetWSAErrorString = "Nonrecoverable error."
  Case 11004: GetWSAErrorString = "Valid name, no data record of requested type."
  Case Else:  GetWSAErrorString = "Unknown Error..."
  End Select
End Function

Public Function args_ctrl(ArgS As String, Type_Args As String, ByRef argument() As String) As Integer
  Dim Dummy As String
  Dim len_args As Integer, i As Integer, ascii As Integer
  Dim s As Integer, e As Integer
  Dim S1 As String
  
  ReDim h(6) As Long
  
  'the arguments of type <username>, <password> and
  '<pathname> are strings
  If Type_Args = "username" Or Type_Args = "password" _
  Or Type_Args = "pathname" Then
    Type_Args = "string"
  End If
  
  'command Ok
  args_ctrl = 0
  
  len_args = Len(ArgS)
  
  Select Case Type_Args
  
  Case "string" '<string>  <string:= char | char&string>
    For i = 1 To len_args
      ascii = Asc(Mid$(ArgS, i, 1))
      If ascii < 32 Or ascii > 126 Then      'only printable characters
        args_ctrl = 3           'sintax error in parameters or arguments
        Exit For
      End If
    Next
    argument(0) = ArgS

  Case "host-port" '<h1,h2,h3,h4,p1,p2>  <h?:=1..255>  <p?:=1..255>
    '<Host> is formed by 4 elements, divided by comma, which representing IP address;
    '<port> is formed by 2 elements, divided by comma, which representing the MSB and LSB of the port.
    'add a separator for simplifing the procedure
    Dummy = ArgS & ","
    Debug.Print "Port String = " & Dummy
    e = 1    'point to next element
    For i = 1 To 6
      s = InStr(e, Dummy, ",") 's point to next separator (ie. comma)
      If s = 0 Then
        args_ctrl = 3          'sintax error in parameters or arguments
        Exit For
      Else
        'every element of the argument must be an integer,
        'represented as string, in the range 1 to 255
        h(i) = Val(Mid$(Dummy, e, s - e))
        Debug.Print "h(" & CStr(i) & ") = " & h(i)
        If h(i) < 0 Or h(i) > 255 Then
          args_ctrl = 3       'sintax error in parameters or arguments
          Exit For
        End If
      End If
      e = s + 1       'point to next element
    Next
    argument(0) = Format$(h(1))              'IP address
    argument(1) = Format$(h(2))
    argument(2) = Format$(h(3))
    argument(3) = Format$(h(4))
    argument(4) = Format$(h(5) * 256 + h(6)) 'port
  
  Case "type-code"  '<A [A N] | I>
    S1 = InStr(ArgS, " ")
    If S1 = 0 Then
      If ArgS = "A" Or ArgS = "" Then
        'arguments assume default values
        argument(0) = "A"  'Ascii
        argument(1) = "N"  'No print
      ElseIf ArgS = "E" Then
        'command not implemented for that parameter
        args_ctrl = 6
        argument(0) = ArgS
      ElseIf ArgS = "I" Then
        argument(0) = "I"
      Else
        'sintax error in parameters or arguments
        args_ctrl = 3
        argument(0) = ArgS
      End If
    Else
      If Left$(ArgS, S1 - 1) = "A" Then
        argument(0) = "A"
        While Mid$(ArgS, S1, 1) = " "
          S1 = S1 + 1
        Wend
        If Mid$(ArgS, S1) = "" Or Mid$(ArgS, S1) = "N" Then
          argument(1) = "N"
        ElseIf Mid$(ArgS, S1) = "T" Then
          'command not implemented for that parameter
          args_ctrl = 6
          argument(1) = Mid$(ArgS, S1)
        ElseIf Mid$(ArgS, S1) = "C" Then
          'command not implemented for that parameter
          args_ctrl = 6
          argument(1) = Mid$(ArgS, S1)
        Else
          'sintax error in parameters or arguments
          args_ctrl = 3
          argument(1) = Mid$(ArgS, S1)
        End If
      ElseIf Left$(ArgS, S1 - 1) = "L" Then
        'command not implemented for that parameter
        args_ctrl = 6
        argument(1) = Mid$(ArgS, S1)
      ElseIf Left$(ArgS, S1 - 1) = "I" Then
        argument(0) = "I"
      Else
        'sintax error in parameters or arguments
        args_ctrl = 3
        argument(0) = Left$(ArgS, S1 - 1)
      End If
    End If
  
  Case "mode-code"  '<S>
    If ArgS = "" Or ArgS = "S" Then
      'argument assumes default value
      argument(0) = "S"  'Stream
    ElseIf ArgS = "B" Then
      'command not implemented for that parameter
      args_ctrl = 6
      argument(0) = ArgS
    ElseIf ArgS = "C" Then
      'command not implemented for that parameter
      args_ctrl = 6
      argument(0) = ArgS
    Else
      'sintax error in parameters or arguments
      args_ctrl = 3
      argument(0) = Left$(ArgS, S1 - 1)
    End If

  Case "structure-code"  '<F | R>
    If ArgS = "" Or ArgS = "F" Then
      'argument assumes default value
      argument(0) = "F" 'File
    ElseIf ArgS = "R" Then
      'command not implemented for that parameter
      args_ctrl = 6
      argument(0) = ArgS
    ElseIf ArgS = "P" Then
      'command not implemented for that parameter
      args_ctrl = 6
      argument(0) = ArgS
    Else
      'sintax error in parameters or arguments
      args_ctrl = 3
      argument(0) = ArgS
    End If
  
  End Select

End Function

Public Function close_data_connect(ID_User As Integer) As Integer
  
  retf = closesocket(users(ID_User).data_slot)
  If retf = 0 Then
    'updates user record
    users(ID_User).data_slot = INVALID_SOCKET
    users(ID_User).IP_Address = ""
    users(ID_User).Port = 0
    users(ID_User).State = Service_Commands ' 2
  End If
  close_data_connect = retf

End Function

Public Function logoff(ID_User As Integer) As Integer

  retf = send_reply("221 Closing control connection, GoodBye!", ID_User)
  retf = closesocket(users(ID_User).control_slot)
  If retf = 0 Then
    're-initialize the record containing user informations
    users(ID_User).list_index = 0
    users(ID_User).control_slot = INVALID_SOCKET
    users(ID_User).data_slot = INVALID_SOCKET
    users(ID_User).IP_Address = ""
    users(ID_User).Port = 0
    users(ID_User).data_representation = "A"
    users(ID_User).data_format_ctrls = "N"
    users(ID_User).data_structure = "F"
    users(ID_User).data_tx_mode = "S"
    users(ID_User).cur_dir = ""
    users(ID_User).State = Log_In_Out ' 0
    users(ID_User).full = False
    users(ID_User).Jenny.Terminate
    Set users(ID_User).Jenny = Nothing
  Else
 '   frmFTP.StatusBar.Panels(1) = "Error: Couldn't Close Connection!"
  End If
  num_users = num_users - 1
 ' frmFTP.UsrCnt = CStr(num_users)
  logoff = retf

End Function

Public Function open_data_connect(ID_User As Integer) As Integer
  
  'open data connection
  retf = send_reply("150 Open data connection.", ID_User)
  open_data_connect = retf

End Function

Public Function receive_data(RecvBuffer As String, ID_User As Integer) As Integer
  Dim fixstr As String * 1024

  'receives data on connection
  retf = recv(users(ID_User).data_slot, fixstr, 1024, 0)
  If retf > 0 Then
    RecvBuffer = Left$(fixstr, retf)
  End If
  receive_data = retf

End Function

Public Function send_data(data_ As String, ID_User As Integer) As Integer
  Dim WriteBuffer As String
  Dim lenBuffer As Integer

  'sends data on connection
  WriteBuffer = data_
  lenBuffer = Len(WriteBuffer)
  retf = send(users(ID_User).data_slot, WriteBuffer, lenBuffer, 0)
  send_data = retf

End Function

Public Function send_reply(reply As String, ID_User As Integer) As Integer
  Dim WriteBuffer As String
  Dim lenBuffer As Integer

  WriteBuffer = reply & vbCrLf
  lenBuffer = Len(WriteBuffer)
  retf = send(users(ID_User).control_slot, WriteBuffer, lenBuffer, 0)
  If retf = SOCKET_ERROR Then
'    ServerLog "Error sending reply:" & CStr(retf)
  Else
    'log replies
'    ServerLog "<" & Format$(ID_User, "000") & "> " & Format$(Date$, "dd/mm/yy ") & Format$(Time$, "hh:mm - ") & reply
  End If
  send_reply = retf

End Function

Public Function sintax_ctrl(cmd As String, ByRef Kwrd As String, ByRef argument() As String) As Integer
  Dim ArgS As String
  Dim k As Integer
  Dim len_cmd As Integer
  
  'the command must be terminated by CR&LF characters
  len_cmd = InStr(cmd, vbCrLf) - 1
  If len_cmd = 0 Then
    sintax_ctrl = 2 'sintax error, command unrecognized
    Exit Function
  Else
    'suppresses CR&LF characters
    cmd = Left$(cmd, len_cmd)
  End If
  
  'extract keyword
  k = InStr(cmd, " ")
  If k <> 0 Then
    'command with arguments
    Kwrd = Left$(cmd, k - 1)  'keyword
    While Mid$(cmd, k, 1) = " "
     k = k + 1
    Wend
    ArgS = Mid$(cmd, k)       'arguments
  Else
    'command without arguments
    Kwrd = cmd
    ArgS = ""
  End If
  
  'command Ok
  sintax_ctrl = 0
  
  Select Case UCase$(Kwrd)
    
  Case "USER"  'USER <username>
    sintax_ctrl = args_ctrl(ArgS, "username", argument())
    
  Case "PASS" 'PASS <password>
    sintax_ctrl = args_ctrl(ArgS, "password", argument())
  
  Case "ACCT"
    sintax_ctrl = 4 'command not implemented
    
  Case "CWD", "XCWD" 'CWD <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
  Case "CDUP", "XCUP"  'CDUP
    '------------------
  
  Case "SMNT"
    sintax_ctrl = 4 'command not implemented
  
  Case "QUIT" 'QUIT
    '-----------------
  
  Case "PORT" 'PORT <host-port>
    sintax_ctrl = args_ctrl(ArgS, "host-port", argument())
  
  Case "PASV"
    sintax_ctrl = 4 'command not implemented
  
  Case "TYPE" 'TYPE <type-code>
    sintax_ctrl = args_ctrl(ArgS, "type-code", argument())
  
  Case "STRU" 'STRU <structure-code>
    sintax_ctrl = args_ctrl(ArgS, "structure-code", argument())
    
  Case "MODE" 'MODE <mode-code>
    sintax_ctrl = args_ctrl(ArgS, "mode-code", argument())
    
  Case "RETR" 'RETR <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
  Case "STOR" 'STOR <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
  Case "RNFR"  'RNFR <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
  Case "RNTO"  'RNTO <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
  Case "ABOR"
    sintax_ctrl = 4 'command not implemented
    
  Case "DELE"  'DELE <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
  Case "RMD", "XRMD" 'RMD <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
  
  Case "MKD", "XMKD" 'MKD <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
  
  Case "PWD", "XPWD" 'PWD
    '----------------
  
  Case "LIST" 'LIST <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
  Case "NLST" 'NLST <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
    
    
  Case "SYST"  'SYST
    '------------------
  
  Case "STAT"  'STAT <pathname>
    sintax_ctrl = args_ctrl(ArgS, "pathname", argument())
  
  Case "HELP"  'HELP <string>
    sintax_ctrl = args_ctrl(ArgS, "string", argument())
    
  Case "NOOP": 'NOOP
    '-----------------
  
  Case "REIN" 'REIN
    sintax_ctrl = 4 'command not implemented
  Case "STOU"
    sintax_ctrl = 4 'command not implemented
  
  Case "APPE"
    sintax_ctrl = 4 'command not implemented
  
  Case "ALLO"
    sintax_ctrl = 1 'command not implemented, superfluous at this side
  
  Case "REST"
    sintax_ctrl = 4 'command not implemented
  
  Case "SITE"
    sintax_ctrl = 4 'command not implemented
  
  Case Else
    sintax_ctrl = 2 'sintax error, command unrecognized
  End Select
  
End Function

Public Sub ServerLog(ByVal str As String)
    
  frmFTP.LogWnd.AddItem str
  frmFTP.LogWnd.Selected(frmFTP.LogWnd.ListCount - 1) = True
End Sub

'EXEC A FTP COMMAND:
'<id_user> is a number in the range 1 to MAX_N_USERS
'identifing the user who sends the command;
'<cmd> is the command.

Public Function ChkPath(ByVal ID_User As Integer, ByVal Arg As String) As String
    If Left$(Arg, 1) = "\" Then
      ChkPath = Left$(users(ID_User).cur_dir, 2) & Arg                  'absolute path
      'ChkPath = DEFAULT_DRIVE & Arg                   'absolute path
    Else
      If Right$(Arg, 1) = ":" And Len(Arg) = 2 Then 'Change Drive letter
        ChkPath = Arg
      ElseIf Right$(users(ID_User).cur_dir, 1) = "\" Then 'relative path
        ChkPath = users(ID_User).cur_dir & Arg        'radix
      Else
        ChkPath = users(ID_User).cur_dir & "\" & Arg
      End If
    End If
End Function

Public Sub SendBuffer(ID_User As Integer, ByRef Buffer As String)
Dim ii As Long
  Debug.Print Buffer
  'sends data in buffer on data connection;
  'data are sending in blocks of 1024 bytes
  ii = 1
  Do While Mid$(Buffer, ii, 1024) <> ""
    retf = send_data(Mid$(Buffer, ii, 1024), ID_User)
    If retf < 0 Then
      retf = WSAGetLastError()
      If retf = WSAEWOULDBLOCK Then
        'try again
      Else
        'error on send
        Exit Do
      End If
    Else
      ii = ii + 1024
    End If
    DoEvents
  Loop
  Buffer = ""
End Sub

Public Sub LIST_NLST(ByVal ID_User As Integer, ByVal Typ As String, ByVal Arg As String)
  Dim File_Name As String, name_ As String, exte_ As String
  Dim DummyS As String
  Dim SepN As Integer
  Dim Full_Name As String 'pathname & file name
  Dim PathName As String, Buffer As String

  If users(ID_User).State = Busy Then  '3
    If InStr(Arg, "-a -L") Then Arg = Left(Arg, (InStr(Arg, "-a -L") - 1))
    If Arg = "" Then
      'if LIST/NLST command has no argument the working directory is the current directory
      PathName = users(ID_User).cur_dir
    Else
      PathName = ChkPath(ID_User, Arg)
    End If
    If InStr(PathName, "*") Or InStr(PathName, "?") Then
      'the GettAttr command blows up with a * or ?
      'possibly because file doesn't exist?
      
      'the pathname indicates a file
      Full_Name = PathName
      File_Name = Dir$(Full_Name)
    ElseIf (GetAttr(PathName) And 16) <> 0 Then
      '--- the pathname indicates a directory
      'if radix then elides final backslash
      If Right$(PathName, 1) = "\" Then
        PathName = Left$(PathName, Len(PathName) - 1)
      End If
      File_Name = Dir$(PathName & "\*.*", 16)
      'rebuilds the full file name
      '(pathname & file name)
      Full_Name = PathName & "\" & File_Name
    Else
      'the pathname indicates a file
      Full_Name = PathName
      File_Name = Dir$(Full_Name)
    End If
    If Err.Number = 0 Then
      'opens data connection
      retf = open_data_connect(ID_User)
      Do
        If Not File_Name = "pagefile.sys" Then
        
        If File_Name = "." Or File_Name = ".." Then
          'parent directories
          DummyS = Format$(File_Name, "@@@@@@@@@@@@!") & " <DIR>"
        ElseIf InStr(Full_Name, "*") Or InStr(Full_Name, "?") Then
          'file
          SepN = InStr(File_Name, ".")
          If SepN <> 0 Then
            'name
            name_ = Left$(File_Name, SepN - 1)
            'extension
            exte_ = Mid$(File_Name, SepN + 1)
          Else
            name_ = File_Name
            exte_ = "   "
          End If
          DummyS = "-rwxr--r--   1 user    group  "
          If Typ = "LIST" Then
            DummyS = DummyS & Format$(FileLen(Full_Name), " @@@@@@@@@") _
             & " " & Format$(FileDateTime(Full_Name), " mmm dd hh:nn ") & File_Name
          ElseIf Typ = "NLST" Then
            'DummyS = Format$(FileLen(Full_Name), " @@@@@@@@@") & " " & File_Name
            DummyS = " " & File_Name & " "
          End If
        ElseIf GetAttr(Full_Name) = 16 Then
          'subdirectory
          SepN = InStr(File_Name, ".")
          If SepN <> 0 Then
            'name
            name_ = Left$(File_Name, SepN - 1)
            'extension
            exte_ = Mid$(File_Name, SepN + 1)
          Else
            name_ = File_Name
            exte_ = "   "
          End If
          DummyS = "drwxr-xr-x   1 user    group  "
          If Typ = "LIST" Then
            DummyS = DummyS & Format$(FileLen(Full_Name), " @@@@@@@@@") _
             & " " & Format$(FileDateTime(Full_Name), " mmm dd hh:nn ") & File_Name
          ElseIf Typ = "NLST" Then
            DummyS = Format$(FileLen(Full_Name), " @@@@@@@@@") & " "
          End If
        Else
          'file
          SepN = InStr(File_Name, ".")
          If SepN <> 0 Then
            'name
            name_ = Left$(File_Name, SepN - 1)
            'extension
            exte_ = Mid$(File_Name, SepN + 1)
          Else
            name_ = File_Name
            exte_ = "   "
          End If
          DummyS = "-rwxr--r--   1 user    group  "
          If Typ = "LIST" Then
            DummyS = DummyS & Format$(FileLen(Full_Name), " @@@@@@@@@") _
             & " " & Format$(FileDateTime(Full_Name), " mmm dd hh:nn ") & File_Name
          ElseIf Typ = "NLST" Then
            DummyS = File_Name
            'DummyS = Format$(FileLen(Full_Name), " @@@@@@@@@") & " " & File_Name
          End If
        End If
        Buffer = Buffer & DummyS & vbCrLf
        File_Name = Dir$
        If Left(File_Name, 1) = "p" Then
          File_Name = Dir$
        End If
      Debug.Print "File Name = " & File_Name
              If File_Name = "" Then Exit Do
              Full_Name = PathName & "\" & File_Name
      Else
        File_Name = Dir$
      End If
      Loop
      SendBuffer ID_User, Buffer
      'close data connection
      retf = send_reply("226 " & Typ & " command completed.", ID_User)
      retf = close_data_connect(ID_User)
    ElseIf (Err.Number > 51 And Err.Number < 77) Or (Err.Number > 707 And Err.Number < 732) Then
      retf = send_reply("450 " & Typ & " command not executed: " & Error$, ID_User)
      retf = close_data_connect(ID_User)
    Else
   '   frmFTP.StatusBar.Panels(1) = "Error " & Err.Number & " occurred."
      retf = close_data_connect(ID_User)
      retf = logoff(ID_User)
      'End
    End If
  ElseIf users(ID_User).State = Service_Commands Then '2
    retf = send_reply("425 Can't open data connection.", ID_User)
  Else
    retf = send_reply("530 User not logged in.", ID_User)
  End If
End Sub

