VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Shell_NotifyIconW Lib "shell32.dll" _
    (ByVal dwMessage As Long, _
     lpData As NOTIFYICONDATA) As Long

Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIM_SETFOCUS = &H3
Const NIM_SETVERSION = &H4

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip(127) As Byte
    dwState As Long
    dwStateMask As Long
    szInfo(511) As Byte
    uTimeout As Long    'uVersion
    szInfoTitle(127) As Byte
    dwInfoFlags As Long
    guidItem As GUID
    hBalloonIcon As Long
End Type

Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4
Const NIF_STATE = &H8
Const NIF_INFO = &H10
Const NIF_GUID = &H20
Const NIF_REALTIME = &H40
Const NIF_SHOWTIP = &H80

Const NIS_HIDDEN = &H1
Const NIS_SHAREDICON = &H2

Const NIIF_NONE = &H0
Const NIIF_INFO = &H1
Const NIIF_WARNING = &H2
Const NIIF_ERROR = &H3
Const NIIF_USER = &H4
Const NIIF_NOSOUND = &H10
Const NIIF_LARGE_ICON = &H20
Const NIIF_RESPECT_QUIET_TIME = &H80
Const NIIF_ICON_MASK = &HF

Const TRAY_ICON_ID = 58

Public Sub Add(ByVal hWnd As Long, ByVal hIcon As Long, ByVal sMsg As String)
    Dim nid As NOTIFYICONDATA
    
    sMsg = sMsg + vbNullChar
    With nid
        .cbSize = Len(nid)
        .hIcon = hIcon
        .hWnd = hWnd
        CopyMemory VarPtr(.szTip(0)), StrPtr(sMsg), Len(sMsg) * 2
        .uCallbackMessage = WM_TRAYICONCLICK
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uID = TRAY_ICON_ID
     End With
   
    Shell_NotifyIconW NIM_ADD, nid
End Sub

Public Sub Remove(ByVal hWnd As Long)
    Dim nid As NOTIFYICONDATA
    
    With nid
        .cbSize = Len(nid)
        .hWnd = hWnd
        .uID = TRAY_ICON_ID
    End With
    
    Shell_NotifyIconW NIM_DELETE, nid
End Sub

