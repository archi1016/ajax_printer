Attribute VB_Name = "api_iphelp"
Option Explicit
    
Public Declare Function inet_addr_by_addr Lib "ws2_32" Alias "inet_addr" _
    (cp As Byte) As Long
    
Public Declare Function GetAdaptersInfo Lib "iphlpapi" _
    (ByVal IpAdapterInfo As Long, _
     pOutBufLen As Long) As Long
     
Public Const ERROR_BUFFER_OVERFLOW = 111

Public Type IP_ADDR_STRING
    Next As Long
    IpAddress(15) As Byte
    IpMask(15) As Byte
    Context As Long
End Type

Public Const MAX_ADAPTER_NAME_LENGTH = 256
Public Const MAX_ADAPTER_ADDRESS_LENGTH = 8
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH = 128

Public Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName(MAX_ADAPTER_NAME_LENGTH + 4 - 1) As Byte
    Description(MAX_ADAPTER_DESCRIPTION_LENGTH + 4 - 1) As Byte
    AddressLength As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DhcpServer As IP_ADDR_STRING
    HaveWins As Long
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Public Const MIB_IF_TYPE_OTHER = 1
Public Const MIB_IF_TYPE_ETHERNET = 6
Public Const IF_TYPE_ISO88025_TOKENRING = 9
Public Const MIB_IF_TYPE_PPP = 23
Public Const MIB_IF_TYPE_LOOPBACK = 24
Public Const MIB_IF_TYPE_SLIP = 28
Public Const IF_TYPE_IEEE80211 = 71     'Vista and later, An IEEE 802.11 wireless network interface


Public Declare Function GetInterfaceInfo Lib "iphlpapi" _
    (ByVal pIfTable As Long, _
     dwOutBufLen As Long) As Long

Public Const NO_ERROR = 0
Public Const ERROR_INSUFFICIENT_BUFFER = 122

Public Const MAX_ADAPTER_NAME = 128

Public Type IP_ADAPTER_INDEX_MAP
    ifIndex As Long
    ifName(MAX_ADAPTER_NAME * 2 - 1) As Byte
End Type


Public Declare Function FlushIpNetTable Lib "iphlpapi" _
    (ByVal dwIfIndex As Long) As Long
   
Public Declare Function CreateIpNetEntry Lib "iphlpapi" _
    (pArpEntry As MIB_IPNETROW) As Long
    
Public Declare Function GetIpNetTable Lib "iphlpapi" ( _
    ByVal pIpNetTable As Long, _
    pdwSize As Long, _
    ByVal bOrder As Long) As Long

Public Const INR_TYPE_Static = 4
Public Const INR_TYPE_Dynamic = 3
Public Const INR_TYPE_Invalid = 2
Public Const INR_TYPE_Other = 1

Public Type MIB_IPNETROW
    dwIndex As Long
    dwPhysAddrLen As Long
    bPhysAddrL As Long
    bPhysAddrH As Long
    dwAddr As Long
    dwType As Long
End Type

Public Declare Function IcmpCreateFile Lib "iphlpapi" () As Long

Public Declare Function IcmpCloseHandle Lib "iphlpapi" _
    (ByVal IcmpHandle As Long) As Long

Public Declare Function IcmpSendEcho Lib "iphlpapi" _
    (ByVal IcmpHandle As Long, _
     ByVal DestinationAddress As Long, _
     ByVal RequestData As Long, _
     ByVal RequestSize As Long, _
     ByVal RequestOptions As Long, _
     ByVal ReplyBuffer As Long, _
     ByVal ReplySize As Long, _
     ByVal Timeout As Long) As Long

Public Declare Function IcmpSendEcho2 Lib "iphlpapi" _
    (ByVal IcmpHandle As Long, _
     ByVal hEvent As Long, _
     ByVal ApcRoutine As Long, _
     ByVal ApcContext As Long, _
     ByVal DestinationAddress As Long, _
     ByVal RequestData As Long, _
     ByVal RequestSize As Long, _
     ByVal RequestOptions As Long, _
     ByVal ReplyBuffer As Long, _
     ByVal ReplySize As Long, _
     ByVal Timeout As Long) As Long
     
Public Type IP_OPTION_INFORMATION
    Ttl As Byte
    Tos As Byte
    flags As Byte
    OptionsSize As Byte
    OptionsData As Long
End Type

Public Const IP_FLAG_REVERSE = &H1
Public Const IP_FLAG_DF = &H2

Public Type ICMP_ECHO_REPLY
    Address As Long
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    Data As Long
    Options As IP_OPTION_INFORMATION
End Type

Public Const IP_STATUS_BASE = 11000

Public Const IP_SUCCESS = 0
Public Const IP_BUF_TOO_SMALL = IP_STATUS_BASE + 1
Public Const IP_DEST_NET_UNREACHABLE = IP_STATUS_BASE + 2
Public Const IP_DEST_HOST_UNREACHABLE = IP_STATUS_BASE + 3
Public Const IP_DEST_PROT_UNREACHABLE = IP_STATUS_BASE + 4
Public Const IP_DEST_PORT_UNREACHABLE = IP_STATUS_BASE + 5
Public Const IP_NO_RESOURCES = IP_STATUS_BASE + 6
Public Const IP_BAD_OPTION = IP_STATUS_BASE + 7
Public Const IP_HW_ERROR = IP_STATUS_BASE + 8
Public Const IP_PACKET_TOO_BIG = IP_STATUS_BASE + 9
Public Const IP_REQ_TIMED_OUT = IP_STATUS_BASE + 10
Public Const IP_BAD_REQ = IP_STATUS_BASE + 11
Public Const IP_BAD_ROUTE = IP_STATUS_BASE + 12
Public Const IP_TTL_EXPIRED_TRANSIT = IP_STATUS_BASE + 13
Public Const IP_TTL_EXPIRED_REASSEM = IP_STATUS_BASE + 14
Public Const IP_PARAM_PROBLEM = IP_STATUS_BASE + 15
Public Const IP_SOURCE_QUENCH = IP_STATUS_BASE + 16
Public Const IP_OPTION_TOO_BIG = IP_STATUS_BASE + 17
Public Const IP_BAD_DESTINATION = IP_STATUS_BASE + 18
Public Const IP_GENERAL_FAILURE = IP_STATUS_BASE + 50

Public Declare Function IcmpParseReplies Lib "iphlpapi" _
    (ByVal ReplyBuffer As Long, _
     ByVal ReplySize As Long) As Long

