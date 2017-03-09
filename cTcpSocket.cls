VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTcpSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const RECV_BUFFER_SIZE As Long = 8192

Public hSocket As Long
Dim BindSA As SOCKADDR
Dim BindSAl As Long
Dim TargetSA As SOCKADDR
Dim TargetSAl As Long
Dim RecvBuffer() As Byte
Public lpRecvBuffer As Long

Dim RetSA As SOCKADDR
Dim RetSAl As Long
Dim RecvDataSize As Long

Private Sub Class_Initialize()
    ReDim RecvBuffer(RECV_BUFFER_SIZE - 1)
    lpRecvBuffer = VarPtr(RecvBuffer(0))
End Sub

Private Sub Class_Terminate()
    Call Free
    Erase RecvBuffer
End Sub

Public Sub SetSocketHandle(ByVal fromSocket As Long)
    hSocket = fromSocket
End Sub

Public Function Create() As Boolean
    hSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    Create = (INVALID_SOCKET <> hSocket)
End Function

Public Sub Free()
    If 0 <> hSocket Then
        shutdown hSocket, SD_BOTH
        closesocket hSocket
    End If
    hSocket = 0
End Sub

Public Function GetRecvBufferSize() As Long
    GetRecvBufferSize = RECV_BUFFER_SIZE
End Function

Public Sub SetRecvTimeout(ByVal mSecond As Long)
    setsockopt hSocket, SOL_SOCKET, SO_RCVTIMEO, mSecond, 4
End Sub

Public Sub SetNoDelay()
    Dim V As Long
    
    V = 1
    setsockopt hSocket, IPPROTO_TCP, TCP_NODELAY, V, 1
End Sub

Private Sub BuildSA(ByVal sInternet As String, ByVal sPort As String, tSA As SOCKADDR, tSAl As Long)
    Dim nPort32 As Long
    Dim nPort16 As Integer
    
    nPort32 = CLng(sPort)
    CopyMemory VarPtr(nPort16), VarPtr(nPort32), 2
    
    With tSA
        .sin_family = AF_INET
        .sin_addr.S_addr = LookupDomainName(sInternet)
        .sin_port = htons(nPort16)
    End With
    tSAl = Len(tSA)
End Sub

Public Sub SetBindInternetAndPort(ByVal sInternet As String, ByVal sPort As String)
    Call BuildSA(sInternet, sPort, BindSA, BindSAl)
End Sub

Public Sub SetTargetInternetAndPort(ByVal sInternet As String, ByVal sPort As String)
    Call BuildSA(sInternet, sPort, TargetSA, TargetSAl)
End Sub

Public Function GetTargetAddr() As Long
    GetTargetAddr = TargetSA.sin_addr.S_addr
End Function

Public Function GetRecvAddr() As Long
    GetRecvAddr = RetSA.sin_addr.S_addr
End Function

Public Function GetRemoteAddr() As Long
    RetSAl = Len(RetSA)
    getpeername hSocket, RetSA, RetSAl
    GetRemoteAddr = RetSA.sin_addr.S_addr
End Function

Public Function GetRecvDataSize() As Long
    GetRecvDataSize = RecvDataSize
End Function

Public Function GetRecvDataOpCode() As Long
    CopyMemory VarPtr(GetRecvDataOpCode), lpRecvBuffer, 4
End Function

Public Function SendData(ByVal lpData As Long, ByVal nDataSize As Long) As Boolean
    SendData = (send(hSocket, lpData, nDataSize, 0) > 0)
End Function

Public Function RecvData() As Boolean
    RecvBuffer(3) = 0
    RecvDataSize = recv(hSocket, lpRecvBuffer, RECV_BUFFER_SIZE, 0)
    RecvData = (RecvDataSize > 0)
End Function

Public Function RecvDataWithSize(ByVal nSize As Long) As Boolean
    RecvBuffer(3) = 0
    RecvDataSize = recv(hSocket, lpRecvBuffer, nSize, 0)
    RecvDataWithSize = (RecvDataSize > 0)
End Function

Public Sub CopyRecvBufferData(ByVal lpData As Long, ByVal nDataSize As Long)
    CopyMemory lpData, lpRecvBuffer, nDataSize
End Sub

Public Function BindToListen() As Boolean
    BindToListen = (SOCKET_ERROR <> bind(hSocket, BindSA, BindSAl))
End Function

Public Function StartListen() As Boolean
    StartListen = (SOCKET_ERROR <> listen(hSocket, SOMAXCONN))
End Function

Public Function ConnectTo() As Boolean
    ConnectTo = (SOCKET_ERROR <> connect(hSocket, TargetSA, TargetSAl))
End Function

Public Function AcceptFrom(newSocket As Long) As Boolean
    RetSAl = Len(RetSA)
    newSocket = accept(hSocket, RetSA, RetSAl)
    AcceptFrom = (INVALID_SOCKET <> newSocket)
End Function

Public Function SetAsyncSelect(ByVal hWnd As Long, ByVal wMessage As Long) As Boolean
    SetAsyncSelect = (SOCKET_ERROR <> WSAAsyncSelect(hSocket, hWnd, wMessage, FD_ACCEPT Or FD_READ Or FD_CLOSE))
End Function

Public Function SetAsyncSelectAccept(ByVal hWnd As Long, ByVal wMessage As Long) As Boolean
    SetAsyncSelectAccept = (SOCKET_ERROR <> WSAAsyncSelect(hSocket, hWnd, wMessage, FD_ACCEPT))
End Function

Public Sub CancelAsyncSelect(ByVal hWnd As Long)
    WSAAsyncSelect hSocket, hWnd, 0, 0
End Sub

Public Sub SetNonBlocking(ByVal B As Boolean)
    Dim V As Long
    
    If B Then
        V = 1
    Else
        V = 0
    End If
    ioctlsocket hSocket, FIONBIO, V
End Sub

Public Function LookupDomainName(sInternet) As Long
    Dim nAddr As Long
    Dim lpRetHE As Long
    Dim HE As HOSTENT
    
    LookupDomainName = INADDR_NONE
    
    nAddr = inet_addr(sInternet)
    If INADDR_NONE = nAddr Then
        lpRetHE = gethostbyname(sInternet)
        If 0 <> lpRetHE Then
            CopyMemory VarPtr(HE), lpRetHE, Len(HE)
            If 0 <> HE.hAddrList Then
                CopyMemory VarPtr(nAddr), HE.hAddrList, 4
                CopyMemory VarPtr(nAddr), nAddr, 4
                LookupDomainName = nAddr
            End If
        End If
    Else
        LookupDomainName = nAddr
    End If
End Function

