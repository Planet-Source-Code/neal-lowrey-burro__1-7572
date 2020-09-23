Attribute VB_Name = "Winsock"
Option Explicit

Public Const AF_UNSPEC = 0             '  /* unspecified */
Public Const AF_UNIX = 1               '  /* local to host (pipes, portals) */
Public Const AF_INET = 2               '  /* internetwork: UDP, TCP, etc. */
Public Const AF_IMPLINK = 3            '  /* arpanet imp addresses */
Public Const AF_PUP = 4                '  /* pup protocols: e.g. BSP */
Public Const AF_CHAOS = 5              '  /* mit CHAOS protocols */
Public Const AF_IPX = 6                '  /* IPX and SPX */
Public Const AF_NS = 6                 '  /* XEROX NS protocols */
Public Const AF_ISO = 7                '  /* ISO protocols */
Public Const AF_OSI = AF_ISO           '  /* OSI is ISO */
Public Const AF_ECMA = 8               '  /* european computer manufacturers */
Public Const AF_DATAKIT = 9            '  /* datakit protocols */
Public Const AF_CCITT = 10             '  /* CCITT protocols, X.25 etc */
Public Const AF_SNA = 11               '  /* IBM SNA */
Public Const AF_DECnet = 12            '  /* DECnet */
Public Const AF_DLI = 13               '  /* Direct data link interface */
Public Const AF_LAT = 14               '  /* LAT */
Public Const AF_HYLINK = 15            '  /* NSC Hyperchannel */
Public Const AF_APPLETALK = 16         '  /* AppleTalk */
Public Const AF_NETBIOS = 17           '  /* NetBios-style addresses */

Public Const FD_READ = &H1
Public Const FD_WRITE = &H2
Public Const FD_OOB = &H4
Public Const FD_ACCEPT = &H8
Public Const FD_CONNECT = &H10
Public Const FD_CLOSE = &H20
Public Const FD_SETSIZE% = 64

Public Const SOL_SOCKET = &HFFFF
Public Const SO_LINGER = &H80

Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1

Public Const BAD_SOCKET = -1
Public Const UNRESOLVED_HOST = -2
Public Const UNABLE_TO_BIND = -3
Public Const UNABLE_TO_CONNECT = -4

 
Public Const WIN_SOCKET_MSG = 2000
Public Const MAX_WSADescription = 257
Public Const MAX_WSASYSStatus = 129

Public Const WS_VERSION_REQD As Integer = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD / &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const IP_OPTIONS = 1
Public Const MIN_SOCKETS_REQD = 0

'--- additional declarations
'Types
Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

'Protocol families, same as address families for now
Public Const PF_UNSPEC = 0
Public Const PF_UNIX = 1
Public Const PF_INET = 2
Public Const PF_IMPLINK = 3
Public Const PF_PUP = 4
Public Const PF_CHAOS = 5
Public Const PF_IPX = 6
Public Const PF_NS = 6
Public Const PF_ISO = 7
Public Const PF_OSI = AF_ISO
Public Const PF_ECMA = 8
Public Const PF_DATAKIT = 9
Public Const PF_CCITT = 10
Public Const PF_SNA = 11
Public Const PF_DECnet = 12
Public Const PF_DLI = 13
Public Const PF_LAT = 14
Public Const PF_HYLINK = 15
Public Const PF_APPLETALK = 16
Public Const PF_NETBIOS = 17

Public Const MAXGETHOSTSTRUCT = 1024

Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17

Public Const INADDR_NONE = &HFFFF
Public Const INADDR_ANY = &H0

' Windows Sockets definitions of regular Microsoft C error constants
Public Const WSAEINTR = 10004
Public Const WSAEBADF = 10009
Public Const WSAEACCES = 10013
Public Const WSAEFAULT = 10014
Public Const WSAEINVAL = 10022
Public Const WSAEMFILE = 10024
' Windows Sockets definitions of regular Berkeley error constants
Public Const WSAEWOULDBLOCK = 10035
Public Const WSAEINPROGRESS = 10036
Public Const WSAEALREADY = 10037
Public Const WSAENOTSOCK = 10038
Public Const WSAEDESTADDRREQ = 10039
Public Const WSAEMSGSIZE = 10040
Public Const WSAEPROTOTYPE = 10041
Public Const WSAENOPROTOOPT = 10042
Public Const WSAEPROTONOSUPPORT = 10043
Public Const WSAESOCKTNOSUPPORT = 10044
Public Const WSAEOPNOTSUPP = 10045
Public Const WSAEPFNOSUPPORT = 10046
Public Const WSAEAFNOSUPPORT = 10047
Public Const WSAEADDRINUSE = 10048
Public Const WSAEADDRNOTAVAIL = 10049
Public Const WSAENETDOWN = 10050
Public Const WSAENETUNREACH = 10051
Public Const WSAENETRESET = 10052
Public Const WSAECONNABORTED = 10053
Public Const WSAECONNRESET = 10054
Public Const WSAENOBUFS = 10055
Public Const WSAEISCONN = 10056
Public Const WSAENOTCONN = 10057
Public Const WSAESHUTDOWN = 10058
Public Const WSAETOOMANYREFS = 10059
Public Const WSAETIMEDOUT = 10060
Public Const WSAECONNREFUSED = 10061
Public Const WSAELOOP = 10062
Public Const WSAENAMETOOLONG = 10063
Public Const WSAEHOSTDOWN = 10064
Public Const WSAEHOSTUNREACH = 10065
Public Const WSAENOTEMPTY = 10066
Public Const WSAEPROCLIM = 10067
Public Const WSAEUSERS = 10068
Public Const WSAEDQUOT = 10069
Public Const WSAESTALE = 10070
Public Const WSAEREMOTE = 10071
' Extended Windows Sockets error constant definitions
Public Const WSASYSNOTREADY = 10091
Public Const WSAVERNOTSUPPORTED = 10092
Public Const WSANOTINITIALISED = 10093
Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004
Public Const WSANO_ADDRESS = 11004

Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public hostent As hostent

Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * MAX_WSADescription '(0 To 255) As Byte
    szSystemStatus As String * MAX_WSASYSStatus  '(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public WSAdata As WSAdata

Type Inet_Address     ' IP Address in Network Order
    Byte4 As Byte     '
    Byte3 As Byte     '
    Byte2 As Byte     '
    Byte1 As Byte     '
End Type

Public IPLong As Inet_Address


'socket address
Type SockAddr
    sin_family As Integer   ' Address family
    sin_port As Integer     ' Port Number in Network Order
    sin_addr As Long        ' IP Address as Long
    sin_zero As String * 8  '(8) As Byte             ' Padding
End Type

Public SockAddr As SockAddr

Public Const SockAddr_Size = 16

Type hostent_async
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
    h_asyncbuffer(MAXGETHOSTSTRUCT) As Byte
End Type

Public hostent_async As hostent_async

Type fd_set
  fd_count As Integer          '' how many are in the set
  fd_array(FD_SETSIZE) As Long '' array of SOCKET handles (64)
End Type

Public fd_set As fd_set

Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Public timeval As timeval

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As SockAddr, addrlen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As SockAddr, ByVal namelen As Long) As Long
    Public Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
    Public Declare Function connect Lib "wsock32.dll" (ByVal s As Long, addr As SockAddr, ByVal namelen As Long) As Long
    Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long
    Public Declare Function getpeername Lib "wsock32.dll" (ByVal s As Long, sName As SockAddr, namelen As Long) As Long
    Public Declare Function getsockname Lib "wsock32.dll" (ByVal s As Long, sName As SockAddr, namelen As Long) As Long
    Public Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal CP As String) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Public Declare Function recv Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function recvfrom Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, from As SockAddr, fromlen As Long) As Long
    Public Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Public Declare Function send Lib "wsock32.dll" (ByVal s As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function sendto Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, to_addr As SockAddr, ByVal tolen As Long) As Long
    Public Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Public Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
    Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Public Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Public Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSAdata) As Long
    Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Public Declare Function WSASetLastError Lib "wsock32.dll" (ByVal iError As Long) As Long
    Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Number As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long

