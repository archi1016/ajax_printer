Attribute VB_Name = "api_socket"
Option Explicit

Public Const WSADESCRIPTION_LEN = 256
Public Const WSASYS_STATUS_LEN = 128

Public Const SOCKET_ERROR = -1
Public Const INVALID_SOCKET = -1

Public Const WSAETIMEDOUT = 10060

Public Type WSAData
   wVersion       As Integer
   wHighVersion   As Integer
   szDescription  As String * WSADESCRIPTION_LEN
   szSystemStatus As String * WSASYS_STATUS_LEN
   iMaxSockets    As Integer
   iMaxUdpDg      As Integer
   lpVendorInfo   As Long
End Type

Public Declare Function WSAStartup Lib "ws2_32" _
    (ByVal wVersionRequired As Long, _
     lpWSAData As WSAData) As Long

Public Declare Function WSACleanup Lib "ws2_32" () As Long

Public Declare Function WSAGetLastError Lib "ws2_32" () As Long

Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2
Public Const AF_IMPLINK = 3
Public Const AF_PUP = 4
Public Const AF_CHAOS = 5
Public Const AF_IPX = 6
Public Const AF_NS = 6
Public Const AF_ISO = 7
Public Const AF_OSI = AF_ISO
Public Const AF_ECMA = 8
Public Const AF_DATAKIT = 9
Public Const AF_CCITT = 10
Public Const AF_SNA = 11
Public Const AF_DECnet = 12
Public Const AF_DLI = 13
Public Const AF_LAT = 14
Public Const AF_HYLINK = 15
Public Const AF_APPLETALK = 16
Public Const AF_NETBIOS = 17
Public Const AF_VOICEVIEW = 18
Public Const AF_FIREFOX = 19
Public Const AF_UNKNOWN1 = 20
Public Const AF_BAN = 21
Public Const AF_MAX = 22

Public Const SOCK_STREAM = 1 'Stream socket
Public Const SOCK_DGRAM = 2 'Datagram socket
Public Const SOCK_RAW = 3 'Raw data socket
Public Const SOCK_RDM = 4 'Reliable Delivery socket
Public Const SOCK_SEQPACKET = 5 'Sequenced Packet socket

Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_IGMP = 2
Public Const IPPROTO_GGP = 3
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77

Public Declare Function Socket Lib "ws2_32" _
    Alias "socket" (ByVal af As _
     Long, ByVal s_type As _
     Long, ByVal Protocol As Long) As Long

Public Declare Function closesocket Lib "ws2_32" _
    (ByVal S As Long) As Long

Public Type IN_ADDR
    S_addr As Long
End Type

Public Const SOCKADDR_LENGTH = 16

Public Type SOCKADDR
    sin_family As Integer
    sin_port As Integer
    sin_addr As IN_ADDR
    sin_zero(7) As Byte
End Type

Public Const INADDR_ANY = 0
Public Const INADDR_BROADCAST = -1
Public Const INADDR_LOOPBACK = &H100007F
Public Const INADDR_NONE = -1

Public Const SD_RECEIVE = 0
Public Const SD_SEND = 1
Public Const SD_BOTH = 2

Public Declare Function inet_addr Lib "ws2_32" _
    (ByVal cp As String) As Long
    
Public Declare Function inet_addr_by_addr Lib "ws2_32" Alias "inet_addr" _
    (cp As Byte) As Long
    
Public Declare Function bind Lib "ws2_32" _
    (ByVal S As Long, _
     name As SOCKADDR, _
     ByVal namelen As Long) As Long
     
Public Const SOL_SOCKET = 65535

Public Const SO_DEBUG = &H1&         ' Turn on debugging info recording
Public Const SO_ACCEPTCONN = &H2&    ' Socket has had listen= - READ-ONLY.
Public Const SO_REUSEADDR = &H4&     ' Allow local address reuse.
Public Const SO_KEEPALIVE = &H8&     ' Keep connections alive.
Public Const SO_DONTROUTE = &H10&    ' Just use interface addresses.
Public Const SO_BROADCAST = &H20&    ' Permit sending of broadcast msgs.
Public Const SO_USELOOPBACK = &H40&  ' Bypass hardware when possible.
Public Const SO_LINGER = &H80&       ' Linger on close if data present.
Public Const SO_OOBINLINE = &H100&   ' Leave received OOB data in line.

Public Const SO_DONTLINGER = Not SO_LINGER
Public Const SO_EXCLUSIVEADDRUSE = Not SO_REUSEADDR ' Disallow local address reuse.

Public Const SO_SNDBUF = &H1001&     ' Send buffer size.
Public Const SO_RCVBUF = &H1002&     ' Receive buffer size.
Public Const SO_SNDLOWAT = &H1003
Public Const SO_RCVLOWAT = &H1004
Public Const SO_SNDTIMEO = &H1005
Public Const SO_RCVTIMEO = &H1006
Public Const SO_ERROR = &H1007&      ' Get error status and clear.
Public Const SO_TYPE = &H1008&       ' Get socket type - READ-ONLY.

Public Const TCP_NODELAY = &H1

Public Const SOMAXCONN = &H7FFFFFFF

Public Declare Function setsockopt Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal level As Long, _
     ByVal optname As Long, _
     optval As Any, _
     ByVal optlen As Long) As Long

Public Const FD_READ = &H1
Public Const FD_WRITE = &H2
Public Const FD_OOB = &H4
Public Const FD_ACCEPT = &H8
Public Const FD_CONNECT = &H10
Public Const FD_CLOSE = &H20

Public Declare Function WSAAsyncSelect Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal hWnd As Long, _
     ByVal wMsg As Integer, _
     ByVal lEvent As Long) As Integer

Public Declare Function sendto Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal buf As Long, _
     ByVal lngLen As Long, _
     ByVal flags As Long, _
     udtTo As SOCKADDR, _
     ByVal tolen As Long) As Long
     
Public Declare Function recvfrom Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal buf As Long, _
     ByVal lngLen As Long, _
     ByVal flags As Long, _
     from As SOCKADDR, _
     fromlen As Long) As Long

Public Const MSG_DONTROUTE = &H4

Public Declare Function htons Lib "ws2_32" _
    (ByVal hostshort As Integer) As Integer

Public Declare Function htonl Lib "ws2_32" _
    (ByVal hostlong As Long) As Long

Public Declare Function gethostbyname Lib "ws2_32" _
    (ByVal host_name As String) As Long
    
Public Type HOSTENT
    hName     As Long
    hAliases  As Long
    hAddrType As Integer
    hLength   As Integer
    hAddrList As Long
End Type


Public Declare Function shutdown Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal how As Long) As Long
     
Public Declare Function send Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal buf As Long, _
     ByVal buflen As Long, _
     ByVal flags As Long) As Long
     
Public Declare Function recv Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal buf As Long, _
     ByVal buflen As Long, _
     ByVal flags As Long) As Long

Public Declare Function connect Lib "ws2_32" _
    (ByVal S As Long, _
     Addr As SOCKADDR, _
     ByVal namelen As Long) As Long
     
Public Declare Function listen Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal backlog As Long) As Long

Public Declare Function accept Lib "ws2_32" _
    (ByVal S As Long, _
     Addr As SOCKADDR, _
     addrlen As Long) As Long

Public Declare Function getsockname Lib "ws2_32" _
    (ByVal S As Long, _
     sname As SOCKADDR, _
     namelen As Long) As Long

Public Declare Function getpeername Lib "ws2_32" _
    (ByVal S As Long, _
     sname As SOCKADDR, _
     namelen As Long) As Long
     
Public Declare Function ioctlsocket Lib "ws2_32" _
    (ByVal S As Long, _
     ByVal V As Long, _
     ut As Long) As Long

Public Const FIONREAD = &H8004667F
Public Const FIONBIO = &H8004667E
Public Const FIOASYNC = &H8004667D


Public Declare Function inet_ntoa Lib "ws2_32" _
    (ByVal InAddr As Long) As Long

