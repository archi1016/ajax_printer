Attribute VB_Name = "func"
Option Explicit

Sub Main()
    Dim C As String
    
    Call InitCommonControls
    Call InitWindowsVersion
    
    C = Trim$(Command)
    If "" = C Then
        WaitingForm.Show 1
        ServerForm.Show
    Else
        Call DecodeCommand(C)
    End If
End Sub

Private Sub DecodeCommand(ByVal C As String)
    Dim A() As String
    Dim U As Long
    
    A = Split(C, CMD_SPLIT_CHAR)
    U = UBound(A)
    Select Case UCase$(A(CMD_NAME))
        Case CMD_INSTALL
            Call CmdInstall
            
        Case CMD_UNINSTALL
            Call CmdUninstall
    End Select
End Sub

Private Sub CmdInstall()
    Dim WS As Object
    Dim SC As Object
    Dim T As String

    T = GetShortcutFilename
    Set WS = CreateObject("WScript.Shell")
    Set SC = WS.CreateShortcut(GetSpecialFolderPath(&H7) + "\" + T + ".lnk")
    With SC
        .TargetPath = App.Path + "\" + App.EXEName + ".exe"
        .Description = T
        .WorkingDirectory = App.Path
        .Save
    End With
    
    Set SC = Nothing
    Set WS = Nothing
End Sub

Private Sub CmdUninstall()
    DeleteFile GetSpecialFolderPath(&H7) + "\" + GetShortcutFilename + ".lnk"
End Sub

Public Function GetShortcutFilename() As String
    GetShortcutFilename = App.ProductName + " (" + APPLICATION_ID + ")"
End Function

Public Function SafeGetNumbersFromString(ByVal S As String) As String
    Dim L As Long
    Dim I As Long
    Dim V As Long
    
    SafeGetNumbersFromString = ""
    If "" <> S Then
        L = Len(S)
        For I = 1 To L
            V = Asc(Mid$(S, I, 1))
            If V >= &H30 Then
                If V <= &H39 Then
                    SafeGetNumbersFromString = SafeGetNumbersFromString + ChrW$(V)
                End If
            End If
        Next
    End If
    
    If "" = SafeGetNumbersFromString Then SafeGetNumbersFromString = "0"
End Function

Public Sub ShellConsole(ByVal sExeFile As String, ByVal sExeArgs As String)
    Dim SEI As SHELLEXECUTEINFO

    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_FLAG_NO_UI Or SEE_MASK_NOZONECHECKS
        .lpVerb = 0
        .lpFile = StrPtr(sExeFile)
        .lpParameters = StrPtr(sExeArgs)
        .nShow = SW_HIDE
    End With
    
    If ShellExecuteExW(SEI) <> 0 Then
        If SEI.hProcess <> 0 Then
            CloseHandle SEI.hProcess
        End If
    End If
    
End Sub

Public Function ConvStringToUtf8(ByVal S As String, Bin() As Byte) As Long
    Dim Ret As Long

    Ret = Len(S) * 3 + 3
    ReDim Bin(Ret - 1)
    ConvStringToUtf8 = WideCharToMultiByte(CP_UTF8, 0, StrPtr(S), Len(S), Bin(0), Ret, 0, 0)
End Function

Public Function SaveImageToPngFile(ByVal hImage As Long, ByVal FP As String) As Boolean
    Dim EncodeGuid As GUID

    CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), EncodeGuid
    DeleteFileW StrPtr(FP)
    SaveImageToPngFile = (GdipSaveImageToFile(hImage, StrPtr(FP), EncodeGuid, 0) = GpStatus_Ok)
End Function

Public Function SaveImageToJpgFile(ByVal hImage As Long, ByVal FP As String, ByVal nQuality As Long) As Boolean
    Dim EncodeGuid As GUID
    Dim EncodeParams As EncoderParameters

    CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), EncodeGuid
    With EncodeParams
        .Count = 1
        With .Parameter
            CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GID
            .NumberOfValues = 1
            .Type = EncoderParameterValueTypeLong
            .Value = VarPtr(nQuality)
        End With
    End With
    SaveImageToJpgFile = (GdipSaveImageToFile(hImage, StrPtr(FP), EncodeGuid, VarPtr(EncodeParams)) = GpStatus_Ok)
End Function

Public Function IsPrinterOnline(ByVal PrinterName As String) As Boolean
    Dim hPrinter As Long
    Dim Buffer() As Byte
    Dim nNeeded As Long
    Dim PI_2 As PRINTER_INFO_2
    Dim devm As DEVMODE
    
    IsPrinterOnline = False
    If 0 <> OpenPrinterW(StrPtr(PrinterName), hPrinter, 0) Then
        ReDim Buffer(4095)
        GetPrinterW hPrinter, 2, VarPtr(Buffer(0)), 0, nNeeded
        ReDim Buffer(nNeeded - 1)
        If 0 <> GetPrinterW(hPrinter, 2, VarPtr(Buffer(0)), nNeeded, nNeeded) Then
            CopyMemory VarPtr(PI_2), VarPtr(Buffer(0)), Len(PI_2)
      '      CopyMemory VarPtr(devm), PI_2.pDevMode, Len(devm)
      'Debug.Print devm.dmPaperWidth & " x " & devm.dmPaperLength
            IsPrinterOnline = (0 = PI_2.Status)
        End If
    
        ClosePrinter hPrinter
    End If
End Function
