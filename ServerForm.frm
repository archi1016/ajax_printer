VERSION 5.00
Begin VB.Form ServerForm 
   Caption         =   "錯誤"
   ClientHeight    =   7965
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   12435
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ServerForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   531
   ScaleMode       =   3  '像素
   ScaleWidth      =   829
   StartUpPosition =   3  '系統預設值
   WindowState     =   2  '最大化
   Begin VB.PictureBox PagePanel 
      Appearance      =   0  '平面
      BackColor       =   &H8000000C&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   2520
      ScaleHeight     =   4095
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.PictureBox PreviewBox 
         Appearance      =   0  '平面
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         DrawStyle       =   6  '內實線
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   240
         ScaleHeight     =   975
         ScaleWidth      =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.Timer TimerForUnload 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   1920
   End
   Begin VB.Timer TimerForHide 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   600
   End
   Begin VB.Menu menuFile 
      Caption         =   "檔案(&F)"
      Begin VB.Menu menuFileExit 
         Caption         =   "真正的關閉本服務 (&X)"
      End
   End
End
Attribute VB_Name = "ServerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TSK As cTcpSocket
Dim AJX As cAJAX

Dim TYI As cTrayIcon

Dim hGdiplusToken As Long
Dim IsCanUnload As Boolean

Private Sub Form_Load()
    SendMessageW Me.hWnd, WM_SETICON, ICON_BIG, LoadResPicture(132, vbResIcon)
    
    Call LoadMemory
    Call LoadControl
    Call LoadSocket
    
    IsCanUnload = False
    
    Me.Caption = App.ProductName
    App.Title = Me.Caption
    
    Call CreateTrayIcon

    OldServerFormProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, ReturnAddressOfFunction(AddressOf NewServerFormProc))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsCanUnload Then
        SetWindowLong Me.hWnd, GWL_WNDPROC, OldServerFormProc
        
        Call FreeSocket
        Call FreeControl
        Call FreeMemory
    Else
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    Dim L As Long
    Dim T As Long
    Dim W As Long
    Dim H As Long

    If vbMinimized <> Me.WindowState Then
        PagePanel.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        PreviewBox.Left = PagePanel.ScaleWidth
    End If
End Sub

Private Sub LoadMemory()
    Set AJX = New cAJAX
    
    AJX.SetPreview PreviewBox
    
    IsCanUnload = True
End Sub

Private Sub FreeMemory()
    Set AJX = Nothing
End Sub

Private Sub LoadControl()
    Dim GISI As GdiplusStartupInput
    
    GISI.GdiplusVersion = 1
    GdiplusStartup hGdiplusToken, GISI, 0
    
    Set TYI = New cTrayIcon
    
    TimerForHide.Enabled = True
End Sub

Private Sub FreeControl()
    TYI.Remove Me.hWnd
    
    GdiplusShutdown hGdiplusToken
        
    Set TYI = Nothing
End Sub

Private Sub LoadSocket()
    Dim WD As WSAData

    WSAStartup &H202, WD
    
    Set TSK = New cTcpSocket
    
    If TSK.Create Then
        TSK.SetBindInternetAndPort "0.0.0.0", CStr(AJAX_SERVICE_LISTEN_PORT)
        'TSK.SetBindInternetAndPort "0.0.0.0", CStr(AJAX_SERVICE_LISTEN_PORT)
        If TSK.BindToListen Then
            If TSK.StartListen Then
                If TSK.SetAsyncSelect(Me.hWnd, WM_TCP_SERVICE) Then
                End If
            End If
        End If
    End If
End Sub

Private Sub FreeSocket()
    TSK.Free
    
    Set TSK = Nothing
    
    WSACleanup
End Sub

Public Sub CreateTrayIcon()
    TYI.Add Me.hWnd, Me.Icon.Handle, Me.Caption
End Sub

Public Sub TcpEventAccept()
    Dim newSocket As Long

    If TSK.AcceptFrom(newSocket) Then
        'If INADDR_LOOPBACK <> TSK.GetRecvAddr Then
        '    closesocket newSocket
        'End If
    End If
End Sub

Public Sub TcpEventClose(ByVal fromSocket As Long)
    shutdown fromSocket, SD_BOTH
    closesocket fromSocket
End Sub

Public Sub TcpEventError(ByVal fromSocket As Long)
    shutdown fromSocket, SD_BOTH
    closesocket fromSocket
End Sub

Public Sub TcpEventRead(ByVal fromSocket As Long)
    AJX.DecodeHttpHeader fromSocket, Me.hWnd
End Sub

Private Sub menuFileExit_Click()
    IsCanUnload = True
    Unload Me
End Sub

Private Sub TimerForHide_Timer()
    TimerForHide.Enabled = False
    Me.Hide
End Sub

Private Sub TimerForUnload_Timer()
    TimerForUnload.Enabled = False
    Call menuFileExit_Click
End Sub

