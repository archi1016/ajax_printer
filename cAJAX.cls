VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAJAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ToDevice As Object

Private Declare Sub GetSystemTime Lib "Kernel32" _
    (lpSystemTime As SYSTEMTIME)
    
Private Sub Class_Initialize()
    '
End Sub

Public Sub SetPreview(toDev As Object)
    Set ToDevice = toDev
End Sub
        
Public Sub DecodeHttpHeader(ByVal fromSocket As Long, ByVal hWnd As Long)
    Dim TSK As New cTcpSocket
    Dim Act As String
    Dim Url As String
    Dim XmlStr As String
    Dim H() As String
    Dim F() As String

    TSK.SetSocketHandle fromSocket
    TSK.CancelAsyncSelect hWnd
    TSK.SetNonBlocking False

    If RecvAllBytes(TSK, Act, Url, XmlStr) Then
        Select Case Act
            Case "POST"
                Select Case Url
                    Case "/?q=list_printer"
                        OperationListPrinter TSK
                     
                    Case "/?q=print"
                        OperationPrint XmlStr, TSK
                        
                    Case "/?q=preview"
                        OperationPreview XmlStr, TSK
                        
                    'Case "/launch"
                    '    ResponseLaunchApplication CMD_LAUNCH, XmlStr, TSK
                        
                    'Case "/install"
                    '    ResponseLaunchApplication CMD_INSTALL, XmlStr, TSK
                        
                    Case Else
                        ResponseNotFound TSK
                        
                End Select
            
            Case "OPTIONS"
                ResponseOptions TSK
                
            Case Else
                ResponseNotFound TSK
                
        End Select
    End If
    TSK.Free
    
    Set TSK = Nothing
    Erase F
    Erase H
End Sub

Private Function OperationPreview(ByVal XmlStr As String, TSK As cTcpSocket) As Boolean
    Dim Xml As Object
    Dim XmlRoot As Object
    Dim Pnt As String
    Dim PR As New cPaper
    Dim S As String
    Dim FP As String

    Set Xml = GetXmlObject
    If Xml.LoadXML(XmlStr) Then
        Set XmlRoot = Xml.documentElement
        If XmlRoot.NodeName = "paper" Then
            S = SelectPrinter(SafeGetXmlNodeAttr(XmlRoot, "printer"))
            If "" = S Then
                With ToDevice
                    .Cls
                    .ForeColor = Printer.ForeColor
                    .Font = Printer.Font
                    .DrawStyle = Printer.DrawStyle
                    .ScaleMode = Printer.ScaleMode
                    .Move .Left, 0, Printer.ScaleWidth, Printer.ScaleHeight
                End With
                S = PR.PrintFromXmlStr(True, XmlRoot, ToDevice)
            End If
        Else
            S = "not paper document"
        End If
    Else
        S = "not xml format"
    End If
    
    If "" <> S Then
        PR.PrintErrorMessage S, ToDevice
    End If
    
    FP = App.Path + "\" + PREVIEW_PNG
    PR.SavePng FP, ToDevice
    OperationPreview = ResponseFile("image/png", FP, TSK)
    
    Set PR = Nothing
    Set XmlRoot = Nothing
    Set Xml = Nothing
End Function

Private Function OperationPrint(ByVal XmlStr As String, TSK As cTcpSocket) As Boolean
    Dim Xml As Object
    Dim XmlRoot As Object
    Dim Pnt As String
    Dim PR As New cPaper
    Dim S As String
    
    Set Xml = GetXmlObject
    If Xml.LoadXML(XmlStr) Then
        Set XmlRoot = Xml.documentElement
        If XmlRoot.NodeName = "paper" Then
            S = SelectPrinter(SafeGetXmlNodeAttr(XmlRoot, "printer"))
            If "" = S Then
                S = PR.PrintFromXmlStr(False, XmlRoot, Printer)
            End If
        Else
            S = "not paper document"
        End If
    Else
        S = "not xml format"
    End If
    OperationPrint = ResponseMsg("{""error"": ""0"",""message"": """ + S + """}", TSK)
    
    Set PR = Nothing
    Set XmlRoot = Nothing
    Set Xml = Nothing
End Function

Private Function OperationListPrinter(TSK As cTcpSocket) As Boolean
    Dim P As Printer
    Dim Json As String
    Dim S As String
    Dim T() As String
    Dim C As Long
    
    Json = "{""error"": ""0"",""message"": """",""items"": ["
    C = 0
    ReDim T(63)
    
    For Each P In Printers
        Set Printer = P
        
        With Printer
            .ScaleMode = vbMillimeters
    
            S = "{"
            S = S + """name"": " + """" + .DeviceName + """"
            S = S + ",""width"": " + """" + CStr(.ScaleWidth) + """"
            S = S + ",""height"": " + """" + CStr(.ScaleHeight) + """"
            S = S + "}"
            
            T(C) = S
            C = C + 1
        End With
    Next
    
    ReDim Preserve T(C - 1)
    Json = Json + Join(T, ",") + "]}"

    OperationListPrinter = ResponseMsg(Json, TSK)
End Function

Private Function RecvAllBytes(TSK As cTcpSocket, retAction As String, retUrl As String, retXmlStr As String) As Boolean
    Dim Buffer() As Byte
    Dim Bin() As Byte
    Dim H() As String
    Dim F() As String
    Dim S As String
    Dim I As Long
    Dim NeedDownloadLength As Long
    Dim Countdown As Long
    Dim lpBuffer As Long
    Dim RetLength As Long
    
    RecvAllBytes = False
    
    If TSK.RecvData Then
        RetLength = TSK.GetRecvDataSize
        ReDim Bin(RetLength - 1)
        TSK.CopyRecvBufferData VarPtr(Bin(0)), RetLength

        H = Split(StrConv(Bin, vbUnicode), vbCrLf + vbCrLf)
        NeedDownloadLength = Len(H(0)) + 4
        
        H = Split(H(0), vbCrLf)
        If INVALID_HANDLE_VALUE <> FindFieldRowIndex("Content-Length", H, S) Then
            NeedDownloadLength = NeedDownloadLength + CLng(S)
        End If
        Countdown = NeedDownloadLength
        
        F = Split(H(0), " ")
        retAction = F(0)
        retUrl = F(1)
        
        ReDim Buffer(NeedDownloadLength - 1)
        lpBuffer = VarPtr(Buffer(0))
        TSK.CopyRecvBufferData lpBuffer, RetLength
        lpBuffer = lpBuffer + RetLength
        Countdown = Countdown - RetLength

        Do While (Countdown > 0)
            If Countdown >= TSK.GetRecvBufferSize Then
                RetLength = TSK.GetRecvBufferSize
            Else
                RetLength = Countdown
            End If
            
            If TSK.RecvDataWithSize(RetLength) Then
                RetLength = TSK.GetRecvDataSize
                TSK.CopyRecvBufferData lpBuffer, RetLength
                lpBuffer = lpBuffer + RetLength
                Countdown = Countdown - RetLength
            Else
                Exit Do
            End If
        Loop
        
        If 0 = Countdown Then
            S = ConvUtf8ToUnicode(Buffer, UBound(Buffer) + 1)
            I = InStr(S, vbCrLf + vbCrLf) + 4
            retXmlStr = Right$(S, Len(S) - (I - 1))
            RecvAllBytes = True
        End If
    End If
    
    Erase H
    Erase F
    Erase Bin
    Erase Buffer
End Function

Private Function ResponseMsg(ByVal M As String, TSK As cTcpSocket) As Boolean
    ResponseMsg = Response("text/html", M, TSK)
End Function

Private Function ResponseXml(ByVal M As String, TSK As cTcpSocket) As Boolean
    ResponseXml = Response("text/xml", M, TSK)
End Function

Private Function Response(ByVal MimeType As String, ByVal M As String, TSK As cTcpSocket) As Boolean
    Dim T As String
    Dim Bin() As Byte

    T = "HTTP/1.1 200 OK" + vbCrLf
    Call AppendHeader(T, "Date", GetGMT)
    Call AppendHeader(T, "Server", GetServerInfo)
    Call AppendHeader(T, "Connection", "close")
    Call AppendHeader(T, "Access-Control-Allow-Origin", "*")
    Call AppendHeader(T, "Access-Control-Allow-Headers", ACCESS_CONTROL_ALLOW_HEADERS)
    Call AppendHeader(T, "Content-Length", CStr(ConvStringToUtf8(M, Bin)))
    Call AppendHeader(T, "Content-Type", MimeType + "; charset=utf-8")
    T = T + vbCrLf + M
    
    Response = SendHeaderFields(T, TSK)
    
    Erase Bin
End Function

Private Function ResponseFile(ByVal MimeType As String, ByVal FP As String, TSK As cTcpSocket) As Boolean
    Dim flBin() As Byte
    Dim flSize As Long
    Dim txBin() As Byte
    Dim txLength As Long
    Dim T As String
    
    If LoadFileToMemory(FP, flBin, flSize) Then
        T = "HTTP/1.1 200 OK" + vbCrLf
        Call AppendHeader(T, "Date", GetGMT)
        Call AppendHeader(T, "Server", GetServerInfo)
        Call AppendHeader(T, "Connection", "close")
        Call AppendHeader(T, "Access-Control-Allow-Origin", "*")
        Call AppendHeader(T, "Access-Control-Allow-Headers", ACCESS_CONTROL_ALLOW_HEADERS)
        Call AppendHeader(T, "Content-Length", CStr(flSize))
        Call AppendHeader(T, "Content-Type", MimeType)
        T = T + vbCrLf
        
        If 0 < flSize Then
            txLength = ConvStringToUtf8(T, txBin)
            ReDim Preserve txBin(txLength + flSize - 1)
            CopyMemory VarPtr(txBin(txLength)), VarPtr(flBin(0)), flSize
            
            ResponseFile = TSK.SendData(VarPtr(txBin(0)), txLength + flSize)
        Else
            ResponseFile = SendHeaderFields(T, TSK)
        End If
    Else
        ResponseFile = ResponseNotFound(TSK)
    End If
    
    Erase flBin
    Erase txBin
End Function

Private Function ResponseOptions(TSK As cTcpSocket) As Boolean
    Dim T As String
    
    T = "HTTP/1.1 200 OK" + vbCrLf
    Call AppendHeader(T, "Date", GetGMT)
    Call AppendHeader(T, "Server", GetServerInfo)
    Call AppendHeader(T, "Connection", "close")
    Call AppendHeader(T, "Access-Control-Allow-Origin", "*")
    Call AppendHeader(T, "Access-Control-Allow-Headers", ACCESS_CONTROL_ALLOW_HEADERS)
    T = T + vbCrLf
    
    ResponseOptions = SendHeaderFields(T, TSK)
End Function

Private Function ResponseNotFound(TSK As cTcpSocket) As Boolean
    Dim T As String
    
    T = "HTTP/1.1 404 Not Found" + vbCrLf
    Call AppendHeader(T, "Date", GetGMT)
    Call AppendHeader(T, "Server", GetServerInfo)
    Call AppendHeader(T, "Connection", "close")
    Call AppendHeader(T, "Access-Control-Allow-Origin", "*")
    Call AppendHeader(T, "Access-Control-Allow-Headers", ACCESS_CONTROL_ALLOW_HEADERS)
    Call AppendHeader(T, "Content-Length", CStr(Len(CONTENT_ERROR_404)))
    Call AppendHeader(T, "Content-type", "text/html")
    T = T + vbCrLf + CONTENT_ERROR_404
    
    ResponseNotFound = SendHeaderFields(T, TSK)
End Function

Private Function SendHeaderFields(ByVal T As String, TSK As cTcpSocket) As Boolean
    Dim Bin() As Byte
    Dim L As Long
    
    L = ConvStringToUtf8(T, Bin)
    SendHeaderFields = TSK.SendData(VarPtr(Bin(0)), L)
    
    Erase Bin
End Function

Private Sub AppendHeader(T As String, ByVal sField As String, ByVal sValue As String)
    T = T + sField + ": " + sValue + vbCrLf
End Sub

Private Function GetServerInfo() As String
    GetServerInfo = "hys httpd"
End Function

Private Function GetGMT() As String
    Dim ST As SYSTEMTIME
    
    Call GetSystemTime(ST)
    GetGMT = ConvSystemTimeToGMT(ST)
End Function

Private Function ConvSystemTimeToGMT(ST As SYSTEMTIME) As String
    Dim wDay() As String
    Dim monh() As String

    wDay = Split("Sun/Mon/Tue/Wed/Thu/Fri/Sat", "/")
    monh = Split("/Jan/Feb/Mar/Apr/May/Jun/Jul/Aug/Sep/Oct/Nov/Dec", "/")
    
    With ST
        ConvSystemTimeToGMT = wDay(.wDayOfWeek) + ", " + ConvIntegerTo2Digi(.wDay) + " " + monh(.wMonth)
        ConvSystemTimeToGMT = ConvSystemTimeToGMT + " " + ConvIntegerTo2Digi(.wHour) + ":" + ConvIntegerTo2Digi(.wMinute) + ":" + ConvIntegerTo2Digi(.wSecond)
        ConvSystemTimeToGMT = ConvSystemTimeToGMT + " GMT"
    End With
    
    Erase wDay
    Erase monh
End Function

Private Function ConvIntegerTo2Digi(ByVal V As Integer) As String
    ConvIntegerTo2Digi = Right$("0" + CStr(V), 2)
End Function

Private Function FindFieldRowIndex(ByVal sField As String, H() As String, retV As String) As Long
    Dim I As Long
    Dim L As Long
    
    FindFieldRowIndex = INVALID_HANDLE_VALUE
    
    sField = sField + ": "
    L = Len(sField)
    
    For I = 1 To UBound(H)
        If sField = Left$(H(I), L) Then
            FindFieldRowIndex = I
            
            retV = Right$(H(I), Len(H(I)) - L)
            Exit For
        End If
    Next
End Function

Private Function SelectPrinter(ByVal PrinterName As String) As String
    Dim P As Printer
   
    SelectPrinter = """" + PrinterName + """ not exist"
    If "" <> PrinterName Then
        For Each P In Printers
            If P.DeviceName = PrinterName Then
                Set Printer = P
                With Printer
                    .ForeColor = 0
                    .FontName = "Verdana"
                    '.FontBold = True
                    .DrawStyle = vbInsideSolid
                    .ScaleMode = vbTwips
                End With
                
                SelectPrinter = ""
                Exit For
            End If
        Next
    End If
End Function


