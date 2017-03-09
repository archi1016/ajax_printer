Attribute VB_Name = "subclass"
Option Explicit

Public OldServerFormProc As Long

Public Function NewServerFormProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim FD As Long
    
    Select Case uMsg
        Case WM_TCP_SERVICE
            CopyMemory VarPtr(FD), VarPtr(lParam) + 2, 2
            If FD = 0 Then
                FD = lParam And 65535
                Select Case FD
                    Case FD_READ
                        ServerForm.TcpEventRead wParam
                        
                    Case FD_ACCEPT
                        ServerForm.TcpEventAccept
                        
                    Case FD_CLOSE
                        ServerForm.TcpEventClose wParam
                        
                End Select
            Else
                ServerForm.TcpEventError wParam
            End If
            
        Case WM_TRAYICONCLICK
            If lParam = WM_RBUTTONUP Then
                ShowWindow hWnd, SW_SHOWMAXIMIZED
                SetForegroundWindow hWnd
            End If
            
        Case WM_TASKBARCREATED
            ServerForm.CreateTrayIcon
            
    End Select
    
    NewServerFormProc = CallWindowProc(OldServerFormProc, hWnd, uMsg, wParam, lParam)
End Function

