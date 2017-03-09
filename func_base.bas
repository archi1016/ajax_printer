Attribute VB_Name = "func_base"
Option Explicit

Public CURRENT_WINDOWS_VERSION As Long

Public Sub MsgError(ByVal M As String)
    MsgBox M, vbExclamation, "提示"
End Sub

Public Sub MsgInfo(ByVal M As String)
    MsgBox M, vbInformation, "訊息"
End Sub

Public Function MsgQuestion(ByVal M As String) As Boolean
    MsgQuestion = (MsgBox(M, vbQuestion Or vbYesNo, "詢問") = vbYes)
End Function

Public Sub MsgErrCode(ByVal M As String)
    M = M + vbCrLf + "錯誤代碼: 0x" + ConvLongToHex(Err.LastDllError, 8)
    MsgBox M, vbExclamation, "提示"
End Sub

Public Sub MsgErrWithCode(ByVal M As String, ByVal E As Long)
    M = M + vbCrLf + "錯誤代碼: 0x" + ConvLongToHex(E, 8)
    MsgBox M, vbExclamation, "提示"
End Sub

Public Sub InitWindowsVersion()
    Dim OSVI As OSVERSIONINFOEX
    
    CURRENT_WINDOWS_VERSION = WINDOWS_VERSION_XP

    OSVI.dwOSVersionInfoSize = Len(OSVI)
    If GetVersionExW(OSVI) <> 0 Then
        If OSVI.dwMajorVersion = 6 Then
            Select Case OSVI.dwMinorVersion
                Case 1
                    CURRENT_WINDOWS_VERSION = WINDOWS_VERSION_7
             
                Case 2
                    CURRENT_WINDOWS_VERSION = WINDOWS_VERSION_8
 
                Case Else
                    CURRENT_WINDOWS_VERSION = WINDOWS_VERSION_VISTA
                
            End Select
        End If
    End If
End Sub

Public Function ConvLongToHex(ByVal V As Long, ByVal L As String) As String
    ConvLongToHex = Right$("0000000" + Hex$(V), L)
End Function

Public Function ConvLongToString(ByVal V As Long, ByVal L As String) As String
    ConvLongToString = Right$("0000000" + CStr(V), L)
End Function

Public Function IsHTTP(ByVal S As String) As Boolean
    IsHTTP = ("http" = LCase$(Left$(S, 4)))
End Function

Public Function GetNowTimeMinutes() As Long
    Dim ST As SYSTEMTIME
    
    Call GetLocalTime(ST)
    GetNowTimeMinutes = CLng(Date) - 40458
    GetNowTimeMinutes = GetNowTimeMinutes * 24 * 60
    GetNowTimeMinutes = GetNowTimeMinutes + CLng(ST.wHour * 60) + CLng(ST.wMinute)
End Function

Public Sub LaunchProgram(ByVal IsWait As Boolean, ByVal ExeFile As String, ByVal ExeArgs As String, ByVal nWindowState As Long)
    Dim SEI As SHELLEXECUTEINFO
    Dim ExePath As String

    If Mid$(ExeFile, 2, 2) = ":\" Then
        ExePath = Left$(ExeFile, InStrRev(ExeFile, "\") - 1)
    Else
        ExePath = ""
    End If
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_FLAG_NO_UI Or SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_NOZONECHECKS
        .lpFile = StrPtr(ExeFile)
        .lpParameters = StrPtr(ExeArgs)
        .lpDirectory = StrPtr(ExePath)
        .nShow = nWindowState
    End With
    If ShellExecuteExW(SEI) <> 0 Then
        If IsWait Then
            WaitForSingleObject SEI.hProcess, INFINITE
        End If
        CloseHandle SEI.hProcess
    Else
        Call MsgErrCode("""" + ExeFile + """ - 程序啟動失敗！")
    End If
End Sub

Public Function CreateProgram(ByVal ExeFile As String, ByVal ExeArgs As String, ByVal nWindowState As Long) As Long
    Dim SEI As SHELLEXECUTEINFO
    Dim ExePath As String

    CreateProgram = 0
    
    If Mid$(ExeFile, 2, 2) = ":\" Then
        ExePath = Left$(ExeFile, InStrRev(ExeFile, "\") - 1)
    Else
        ExePath = ""
    End If
    
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_FLAG_NO_UI Or SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_NOZONECHECKS
        .lpFile = StrPtr(ExeFile)
        .lpParameters = StrPtr(ExeArgs)
        .lpDirectory = StrPtr(ExePath)
        .nShow = nWindowState
    End With
    If ShellExecuteExW(SEI) <> 0 Then
        CreateProgram = SEI.hProcess
    End If
End Function

Public Function CreateAnyId() As Long
    Dim I As Long
    Dim L(3) As Long
    Dim lpL As Long
    Dim lpCreateAnyId As Long
    
    Randomize Timer
    For I = 0 To 3
        L(I) = CLng(Rnd * 1023)
    Next
    
    lpL = VarPtr(L(0))
    lpCreateAnyId = VarPtr(CreateAnyId)
    CopyMemory lpCreateAnyId + 3, lpL, 1
    CopyMemory lpCreateAnyId + 2, lpL + 4, 1
    CopyMemory lpCreateAnyId + 1, lpL + 8, 1
    CopyMemory lpCreateAnyId, lpL + 12, 1
End Function

Public Function FindFilesAtFolder(ByVal fromFolder As String, ByVal Filters As String, FFS() As String, FFC As Long) As Boolean
    Dim hFind As Long
    Dim WFD As WIN32_FIND_DATA
    Dim FN As String
    
    FindFilesAtFolder = False
    FFC = 0
    
    fromFolder = fromFolder + "\" + Filters
    hFind = FindFirstFileW(StrPtr(fromFolder), WFD)
    If hFind <> INVALID_HANDLE_VALUE Then
        ReDim FFS(511)
        Do
            With WFD
                .dwFileAttributes = .dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
                If .dwFileAttributes = 0 Then
                    FN = String$(MAX_PATH, vbNullChar)
                    CopyMemory StrPtr(FN), VarPtr(.cFileName(0)), MAX_PATH * 2
                    FN = StrCutNull(FN)
                    FFS(FFC) = FN
                    FFC = FFC + 1
                End If
            End With
        Loop Until (FindNextFileW(hFind, WFD) = 0)
        FindClose hFind
        FindFilesAtFolder = (FFC > 0)
    End If
End Function

Public Function FindFoldersAtFolder(ByVal fromFolder As String, FFS() As String, FFC As Long) As Boolean
    Dim hFind As Long
    Dim WFD As WIN32_FIND_DATA
    Dim FN As String
    
    FindFoldersAtFolder = False
    FFC = 0
    
    fromFolder = fromFolder + "\*"
    hFind = FindFirstFileW(StrPtr(fromFolder), WFD)
    If hFind <> INVALID_HANDLE_VALUE Then
        ReDim FFS(511)
        Do
            With WFD
                .dwFileAttributes = .dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY
                If .dwFileAttributes <> 0 Then
                    FN = String$(MAX_PATH, vbNullChar)
                    CopyMemory StrPtr(FN), VarPtr(.cFileName(0)), MAX_PATH * 2
                    FN = StrCutNull(FN)
                    If FN <> "." Then
                        If FN <> ".." Then
                            FFS(FFC) = FN
                            FFC = FFC + 1
                        End If
                    End If
                End If
            End With
        Loop Until (FindNextFileW(hFind, WFD) = 0)
        FindClose hFind
        
        FindFoldersAtFolder = (FFC > 0)
    End If
End Function

Public Function StrCutNull(ByVal S As String) As String
    Dim I As Long
    
    If Len(S) > 0 Then
        I = InStr(S, vbNullChar)
        If I > 0 Then
            StrCutNull = Left$(S, I - 1)
        Else
            StrCutNull = S
        End If
    Else
        StrCutNull = ""
    End If
End Function

Public Function ReadStrFromIniFile(ByVal IniFile As String, ByVal KeyName As String) As String
    ReadStrFromIniFile = ReadStrFromIniFileEx(IniFile, "CONFIG", KeyName)
End Function

Public Function ReadStrFromIniFileEx(ByVal IniFile As String, ByVal SectionName As String, ByVal KeyName As String) As String
    Dim Ret As Long
    
    If Mid$(IniFile, 2, 2) <> ":\" Then IniFile = App.Path + IniFile
    ReadStrFromIniFileEx = String$(1024, vbNullChar)
    Ret = GetPrivateProfileStringW(StrPtr(SectionName), StrPtr(KeyName), 0, StrPtr(ReadStrFromIniFileEx), 1024, StrPtr(IniFile))
    ReadStrFromIniFileEx = Left$(ReadStrFromIniFileEx, Ret)
End Function

Public Sub WriteStrToIniFile(ByVal IniFile As String, ByVal KeyName As String, ByVal ValueStr As String)
    Call WriteStrToIniFileEx(IniFile, "CONFIG", KeyName, ValueStr)
End Sub

Public Sub WriteStrToIniFileEx(ByVal IniFile As String, ByVal SectionName As String, ByVal KeyName As String, ByVal ValueStr As String)
    Dim Ret As Long
    
    If Mid$(IniFile, 2, 2) <> ":\" Then IniFile = App.Path + IniFile
    If ValueStr <> "" Then
        WritePrivateProfileStringW StrPtr(SectionName), StrPtr(KeyName), StrPtr(ValueStr), StrPtr(IniFile)
    Else
        WritePrivateProfileStringW StrPtr(SectionName), StrPtr(KeyName), 0, StrPtr(IniFile)
    End If
End Sub

Public Function IsFolderExist(ByVal FP As String) As Boolean
    Dim dwAttr As Long
    
    IsFolderExist = False
    dwAttr = GetFileAttributesW(StrPtr(FP))
    If dwAttr <> INVALID_FILE_ATTRIBUTES Then
        dwAttr = dwAttr And FILE_ATTRIBUTE_DIRECTORY
        IsFolderExist = (dwAttr <> 0)
    End If
End Function

Public Function IsFileExist(ByVal FP As String) As Boolean
    Dim dwAttr As Long
    
    IsFileExist = False
    dwAttr = GetFileAttributesW(StrPtr(FP))
    If dwAttr <> INVALID_FILE_ATTRIBUTES Then
        dwAttr = dwAttr And FILE_ATTRIBUTE_DIRECTORY
        IsFileExist = (dwAttr = 0)
    End If
End Function

Public Sub CheckAndCreateFolder(ByVal FP As String)
    If Not IsFolderExist(FP) Then
        SHCreateDirectory 0, StrPtr(FP)
    End If
End Sub

Public Function CopyFile(ByVal srcFile As String, ByVal desFile As String) As Boolean
    SetFileAttributesW StrPtr(desFile), FILE_ATTRIBUTE_NORMAL
    CopyFile = (0 <> CopyFileW(StrPtr(srcFile), StrPtr(desFile), 0))
End Function

Public Function MoveFile(ByVal srcFile As String, ByVal desFile As String) As Boolean
    MoveFile = False
    
    If DeleteFile(desFile) Then
        MoveFile = (0 <> MoveFileW(StrPtr(srcFile), StrPtr(desFile)))
    End If
End Function

Public Function ConvIpToStr(ByVal nAddr As Long) As String
    Dim B(3) As Byte
    
    CopyMemory VarPtr(B(0)), VarPtr(nAddr), 4
    ConvIpToStr = CStr(B(0)) + "." + CStr(B(1)) + "." + CStr(B(2)) + "." + CStr(B(3))
End Function

Public Function ConsoleLaunch(ByVal sExeFile As String, ByVal sExeArgs As String) As Boolean
    Dim SEI As SHELLEXECUTEINFO

    ConsoleLaunch = False
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_DOENVSUBST
        .lpVerb = 0
        .lpFile = StrPtr(sExeFile)
        .lpParameters = StrPtr(sExeArgs)
        .nShow = SW_HIDE
    End With
    
    If ShellExecuteExW(SEI) <> 0 Then
        If SEI.hInstApp > 32 Then
            If SEI.hProcess <> 0 Then
                WaitForSingleObject SEI.hProcess, INFINITE
                CloseHandle SEI.hProcess
                ConsoleLaunch = True
            End If
        End If
    End If
    
End Function

Public Function GetTempFolder() As String
    Dim Ret As Long
    
    GetTempFolder = String$(MAX_PATH, vbNullChar)
    Ret = GetTempPathW(MAX_PATH, StrPtr(GetTempFolder))
    GetTempFolder = Left$(GetTempFolder, Ret - 1)
End Function

Public Function LaunchProcessAndReturnHandle(ByVal ExeFile As String, ByVal ExeArgs As String) As Long
    Dim si As STARTUPINFO
    Dim ExePath As String
    Dim pi As PROCESS_INFORMATION
    Dim lpExeArgs As Long
    Dim lpExePath As Long
    
    LaunchProcessAndReturnHandle = 0
    
    si.cb = Len(si)
    GetStartupInfoW si
    si.dwFlags = si.dwFlags Or STARTF_USESHOWWINDOW
    si.wShowWindow = SW_SHOWNORMAL
    
    If Mid$(ExeFile, 2, 2) = ":\" Then
        ExePath = ReturnParentDirectory(ExeFile)
        lpExePath = StrPtr(ExePath)
    Else
        lpExePath = 0
    End If

    If InStr(ExeFile, " ") > 0 Then
        ExeFile = """" + ExeFile + """"
    End If

    If ExeArgs <> "" Then
        ExeFile = ExeFile + " " + ExeArgs
    End If

    If CreateProcessW(0, StrPtr(ExeFile), 0, 0, &HFFFFFFFF, 0, 0, lpExePath, si, pi) <> 0 Then
        CloseHandle pi.hThread
        LaunchProcessAndReturnHandle = pi.hProcess
    End If
End Function

Public Sub LaunchProcess(ByVal IsWait As Boolean, ByVal CpuIndex As Long, ByVal ExeFile As String, ByVal ExeArgs As String, ByVal nWindowState As Long)
    Dim SysInfo As SYSTEM_INFO
    Dim CpuMask As Long
    Dim si As STARTUPINFO
    Dim ExePath As String
    Dim pi As PROCESS_INFORMATION
    Dim lpExeArgs As Long
    Dim lpExePath As Long
    
    If CpuIndex <> INVALID_HANDLE_VALUE Then
        GetSystemInfo SysInfo
        If CpuIndex >= SysInfo.dwNumberOfProcessors Then
            CpuIndex = 0
        End If
        CpuMask = 2 ^ CpuIndex
        SetProcessAffinityMask GetCurrentProcess, CpuMask
    End If
    
    si.cb = Len(si)
    GetStartupInfoW si
    si.dwFlags = si.dwFlags Or STARTF_USESHOWWINDOW
    si.wShowWindow = nWindowState
    
    If Mid$(ExeFile, 2, 2) = ":\" Then
        ExePath = ReturnParentDirectory(ExeFile)
        lpExePath = StrPtr(ExePath)
    Else
        lpExePath = 0
    End If

    If InStr(ExeFile, " ") > 0 Then
        ExeFile = """" + ExeFile + """"
    End If

    If ExeArgs <> "" Then
        ExeFile = ExeFile + " " + ExeArgs
    End If

    If CreateProcessW(0, StrPtr(ExeFile), 0, 0, &HFFFFFFFF, 0, 0, lpExePath, si, pi) <> 0 Then
        If IsWait Then
            WaitForSingleObject pi.hProcess, INFINITE
        End If
        CloseHandle pi.hThread
        CloseHandle pi.hProcess
    Else
        Call MsgErrCode("""" + ExeFile + """ - 程序啟動失敗！")
    End If
End Sub

Public Sub ShellAndWaitProgram(ByVal ExeFile As String, ByVal ExeArgs As String)
    Call LaunchProgram(True, ExeFile, ExeArgs, SW_SHOWNORMAL)
End Sub

Public Sub ShellProgram(ByVal ExeFile As String, ByVal ExeArgs As String)
    Call LaunchProgram(False, ExeFile, ExeArgs, SW_SHOWNORMAL)
End Sub

Public Sub ShellProgramMax(ByVal ExeFile As String, ByVal ExeArgs As String)
    Call LaunchProgram(False, ExeFile, ExeArgs, SW_SHOWMAXIMIZED)
End Sub

Public Sub LaunchAnotherProcess(ByVal sArguments As String)
    Dim sExeFile As String
    
    sExeFile = App.Path + "\" + App.EXEName + ".exe"
    Call ShellProgram(sExeFile, sArguments)
End Sub

Public Function ReturnLastDirectoryName(ByVal FP As String) As String
    ReturnLastDirectoryName = Right$(FP, Len(FP) - InStrRev(FP, "\"))
End Function

Public Function ReturnParentDirectory(ByVal FP As String) As String
    ReturnParentDirectory = Left$(FP, InStrRev(FP, "\") - 1)
End Function

Public Function GetSpecialFolderPath(ByVal csidl As Long) As String
    GetSpecialFolderPath = String$(MAX_PATH, vbNullChar)
    SHGetFolderPathW 0, csidl, 0, 0, StrPtr(GetSpecialFolderPath)
    GetSpecialFolderPath = StrCutNull(GetSpecialFolderPath)
End Function

Public Function GetSelectedFolder(ByVal hWnd As Long, ByVal sTitle As String) As String
    Dim BI As BROWSEINFO
    Dim DN As String
    Dim Ret As Long
    
    GetSelectedFolder = ""
    DN = String$(MAX_PATH, vbNullChar)
    With BI
        .hWndOwner = hWnd
        .pszDisplayName = StrPtr(DN)
        .lpszTitle = StrPtr(sTitle)
        .ulFlags = BIF_USENEWUI Or BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_RETURNFSANCESTORS
    End With
    
    Ret = SHBrowseForFolderW(BI)
    If Ret <> 0 Then
        If SHGetPathFromIDListW(Ret, StrPtr(DN)) <> 0 Then
            GetSelectedFolder = StrCutNull(DN)
        End If
        CoTaskMemFree Ret
    End If
End Function

Public Function GetSelectedFile(ByVal hWnd As Long, ByVal sTitle As String, ByVal sFileReadme As String, ByVal sFilter As String) As String
    Dim CFP As String
    Dim OFN As OPENFILENAME
    Dim lpstrFilter As String
    Dim lpstrFile As String
    
    CFP = App.Path
    GetSelectedFile = ""
    lpstrFilter = sFileReadme + " (" + sFilter + ")" + vbNullChar + sFilter + vbNullChar + vbNullChar
    lpstrFile = String$(MAX_PATH, vbNullChar)
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = hWnd
        .lpstrFilter = StrPtr(lpstrFilter)
        .lpstrFile = StrPtr(lpstrFile)
        .nMaxFile = Len(lpstrFile)
        .lpstrTitle = StrPtr(sTitle)
        .flags = OFN_EXPLORER Or OFN_DONTADDTORECENT Or OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_NODEREFERENCELINKS Or OFN_NONETWORKBUTTON Or OFN_PATHMUSTEXIST
    End With
    If GetOpenFileNameW(OFN) <> 0 Then
        GetSelectedFile = StrCutNull(lpstrFile)
    End If
    SetCurrentDirectoryW StrPtr(CFP)
End Function

Public Function GetSelectedFiles(ByVal hWnd As Long, ByVal sTitle As String, ByVal sFileReadme As String, ByVal sFilter As String, FFS() As String, FFC As Long) As Boolean
    Dim CFP As String
    Dim OFN As OPENFILENAME
    Dim lpstrFilter As String
    Dim lpstrFile As String
    Dim I As Long
    Dim S() As String
    
    CFP = App.Path
    FFC = 0
    lpstrFilter = sFileReadme + " (" + sFilter + ")" + vbNullChar + sFilter + vbNullChar + vbNullChar
    lpstrFile = String$(32768, vbNullChar)
    With OFN
        .lStructSize = Len(OFN)
        .hWndOwner = hWnd
        .lpstrFilter = StrPtr(lpstrFilter)
        .lpstrFile = StrPtr(lpstrFile)
        .nMaxFile = Len(lpstrFile)
        .lpstrTitle = StrPtr(sTitle)
        .flags = OFN_EXPLORER Or OFN_DONTADDTORECENT Or OFN_ENABLESIZING Or OFN_FILEMUSTEXIST Or OFN_NODEREFERENCELINKS Or OFN_NONETWORKBUTTON Or OFN_PATHMUSTEXIST Or OFN_ALLOWMULTISELECT Or OFN_FORCESHOWHIDDEN
    End With
    If GetOpenFileNameW(OFN) <> 0 Then
        S = Split(lpstrFile, vbNullChar)
        If S(1) = "" Then
            ReDim FFS(0)
            FFS(0) = S(0)
            FFC = 1
        Else
            S = Split(lpstrFile, vbNullChar + vbNullChar)
            S = Split(S(0), vbNullChar)
            lpstrFile = S(0)
            If Right$(lpstrFile, 1) <> "\" Then lpstrFile = lpstrFile + "\"
            
            FFC = UBound(S)
            ReDim FFS(FFC - 1)
            For I = 1 To FFC
                FFS(I - 1) = lpstrFile + S(I)
            Next
        End If
    End If
    SetCurrentDirectoryW StrPtr(CFP)
    GetSelectedFiles = (FFC > 0)
    
    Erase S
End Function

Public Function DeleteFile(ByVal FP As String) As Boolean
    Dim dwAttr As Long
    
    DeleteFile = True
    dwAttr = GetFileAttributesW(StrPtr(FP))
    If dwAttr <> INVALID_FILE_ATTRIBUTES Then
        SetFileAttributesW StrPtr(FP), FILE_ATTRIBUTE_NORMAL
        DeleteFile = (DeleteFileW(StrPtr(FP)) <> 0)
    End If
End Function

Public Function RemoveDirectory(ByVal FP As String) As Boolean
    Dim dwAttr As Long
    
    RemoveDirectory = True
    dwAttr = GetFileAttributesW(StrPtr(FP))
    If dwAttr <> INVALID_FILE_ATTRIBUTES Then
        SetFileAttributesW StrPtr(FP), FILE_ATTRIBUTE_NORMAL
        RemoveDirectory = (RemoveDirectoryW(StrPtr(FP)) <> 0)
    End If
End Function

Public Function GetComputerName() As String
    Dim Ret As Long
    
    Ret = MAX_PATH
    GetComputerName = String$(Ret, vbNullChar)
    GetComputerNameW StrPtr(GetComputerName), Ret
    GetComputerName = StrCutNull(GetComputerName)
End Function

Public Function GetWindowsDirectory() As String
    Dim Ret As Long
    
    GetWindowsDirectory = String$(MAX_PATH, vbNullChar)
    Ret = GetWindowsDirectoryW(StrPtr(GetWindowsDirectory), MAX_PATH)
    GetWindowsDirectory = StrCutNull(GetWindowsDirectory)
End Function

Public Function GetSystemDirectory() As String
    Dim Ret As Long
    
    GetSystemDirectory = String$(MAX_PATH, vbNullChar)
    Ret = GetSystemDirectoryW(StrPtr(GetSystemDirectory), MAX_PATH)
    GetSystemDirectory = StrCutNull(GetSystemDirectory)
End Function

Public Function GetSystemWow64Directory() As String
    Dim Ret As Long
    
    GetSystemWow64Directory = String$(MAX_PATH, vbNullChar)
    Ret = GetSystemWow64DirectoryW(StrPtr(GetSystemWow64Directory), MAX_PATH)
    GetSystemWow64Directory = StrCutNull(GetSystemWow64Directory)
    
    If "" = GetSystemWow64Directory Then GetSystemWow64Directory = GetSystemDirectory
End Function

Public Function ConvFileSizeToStr(ByVal nSize As Currency) As String
    If nSize >= 1073741824 Then
        ConvFileSizeToStr = CStr(CSng(CLng((nSize * 10) / 1073741824)) / 10) + "GiB"
    Else
        If nSize >= 1048576 Then
            ConvFileSizeToStr = CStr(CSng(CLng((nSize * 10) / 1048576)) / 10) + "MiB"
        Else
            If nSize >= 1024 Then
                ConvFileSizeToStr = CStr(CSng(CLng((nSize * 10) / 1024)) / 10) + "KiB"
            Else
                ConvFileSizeToStr = CStr(nSize)
            End If
        End If
    End If
End Function

Public Function ReturnAddressOfFunction(ByVal A As Long) As Long
    ReturnAddressOfFunction = A
End Function

Public Function IsSpaceEnough(ByVal FP As String, ByVal V As Currency) As Boolean
    Dim F As Currency
    
    IsSpaceEnough = False
    
    V = V * 1.05
    If Mid$(FP, 2, 2) = ":\" Then
        FP = Left$(FP, 3)
        If GetDiskFreeSpaceExW(StrPtr(FP), VarPtr(F), 0, 0) <> 0 Then
            F = F * 10000
            IsSpaceEnough = (F > V)
        End If
    End If
End Function

Public Function RemoveFolder(ByVal hWnd As Long, ByVal sTitle As String, ByVal FP As String) As Boolean
    Dim SFOP As SHFILEOPSTRUCT
    
    FP = FP + vbNullChar
    With SFOP
        .hWnd = hWnd
        .wFunc = FO_DELETE
        .pFrom = StrPtr(FP)
        .fFlags = FOF_NOCONFIRMATION
        .lpszProgressTitle = StrPtr(sTitle)
    End With
    RemoveFolder = (0 = SHFileOperationW(SFOP))
End Function

Public Function RemoveFolderSilent(ByVal hWnd As Long, ByVal sTitle As String, ByVal FP As String) As Boolean
    Dim SFOP As SHFILEOPSTRUCT
    
    FP = FP + vbNullChar
    With SFOP
        .hWnd = hWnd
        .wFunc = FO_DELETE
        .pFrom = StrPtr(FP)
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
        .lpszProgressTitle = StrPtr(sTitle)
    End With
    RemoveFolderSilent = (0 = SHFileOperationW(SFOP))
End Function

Public Function DumpMemoryToFile(ByVal lpBuffer As Long, ByVal nWriteSize As Long, ByVal FP As String) As Boolean
    Dim hFile As Long
    Dim Ret As Long
    
    DumpMemoryToFile = False
    
    DeleteFile FP
    hFile = CreateFileW(StrPtr(FP), GENERIC_WRITE, FILE_SHARE_READ, 0, CREATE_NEW, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If INVALID_HANDLE_VALUE <> hFile Then
        WriteFile hFile, lpBuffer, nWriteSize, Ret, 0
        CloseHandle hFile
        
        DumpMemoryToFile = True
    End If
End Function

Public Function LoadFileToMemory(ByVal FP As String, Buffer() As Byte, nReadSize As Long) As Boolean
    Dim hFile As Long
    Dim Ret As Long
    
    LoadFileToMemory = False
    
    hFile = CreateFileW(StrPtr(FP), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If INVALID_HANDLE_VALUE <> hFile Then
        nReadSize = GetFileSize(hFile, Ret)
        If nReadSize > 0 Then
            ReDim Buffer(nReadSize - 1)
            ReadFile hFile, VarPtr(Buffer(0)), nReadSize, Ret, 0
        End If
        CloseHandle hFile
        
        LoadFileToMemory = True
    End If
End Function

Public Function LoadDataFromFile(ByVal FP As String, ByVal lpBuffer As Long, ByVal nReadSize As Long) As Boolean
    Dim hFile As Long
    Dim Ret As Long
    
    LoadDataFromFile = False
    
    hFile = CreateFileW(StrPtr(FP), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If INVALID_HANDLE_VALUE <> hFile Then
        ReadFile hFile, lpBuffer, nReadSize, Ret, 0
        CloseHandle hFile
        
        LoadDataFromFile = True
    End If
End Function

Public Function GetSelectedColor(ByVal hWnd As Long, ByVal CurrentColor As Long, CustomColors() As Long) As Long
    Dim CSC As ChooseColor
    
    With CSC
        .lStructSize = Len(CSC)
        .hWndOwner = hWnd
        .rgbResult = CurrentColor
        .lpCustColors = VarPtr(CustomColors(0))
        .flags = CC_FULLOPEN Or CC_RGBINIT
    End With
    If 0 <> ChooseColorW(CSC) Then
        GetSelectedColor = CSC.rgbResult
    Else
        GetSelectedColor = CurrentColor
    End If
End Function

Public Function ConvHSLtoRGB(ByVal H As Single, ByVal S As Single, ByVal L As Single) As Long
    Dim RGB(2) As Single
    Dim v1 As Single
    Dim v2 As Single
    Dim ColorRGB(2) As Long
    Dim I As Long
    Dim t1 As Single
    Dim t2 As Single
    
    H = H / 360
    S = S / 100
    L = L / 100
    
    
    If S = 0 Then
        ColorRGB(0) = CLng(L * 255)
        ColorRGB(1) = ColorRGB(0)
        ColorRGB(2) = ColorRGB(0)
    Else
        If L < 0.5 Then
            t2 = L * (1 + S)
        Else
            t2 = L + S - (L * S)
        End If
        t1 = 2 * L - t2
        v1 = 1 / 3
        v2 = 2 / 3
        RGB(0) = H + v1
        RGB(1) = H
        RGB(2) = H - v1
        For I = 0 To 2
            If RGB(I) < 0 Then RGB(I) = RGB(I) + 1
            If RGB(I) > 1 Then RGB(I) = RGB(I) - 1
            If RGB(I) * 6 < 1 Then
                RGB(I) = t1 + (t2 - t1) * 6 * RGB(I)
            Else
                If RGB(I) * 2 < 1 Then
                    RGB(I) = t2
                Else
                    If RGB(I) * 3 < 2 Then
                       RGB(I) = t1 + (t2 - t1) * (v2 - RGB(I)) * 6
                    Else
                        RGB(I) = t1
                    End If
                End If
            End If
            RGB(I) = RGB(I) * 255
            If RGB(I) > 255 Then RGB(I) = 255
            If RGB(I) < 0 Then RGB(I) = 0
            ColorRGB(I) = CLng(RGB(I))
        Next
    End If
    
    I = VarPtr(ConvHSLtoRGB)
    CopyMemory I, VarPtr(ColorRGB(0)), 1
    CopyMemory I + 1, VarPtr(ColorRGB(1)), 1
    CopyMemory I + 2, VarPtr(ColorRGB(2)), 1
End Function

Public Function ConvColorGdiToGdiplus(ByVal nColor As Long) As Long
    Dim lpS As Long
    Dim lpD As Long
    
    ConvColorGdiToGdiplus = nColor Or &HFF000000
    lpS = VarPtr(nColor)
    lpD = VarPtr(ConvColorGdiToGdiplus)
    CopyMemory lpD, lpS + 2, 1
    CopyMemory lpD + 2, lpS, 1
End Function

Public Function ConvColorGdiplusToGdi(ByVal nColor As Long) As Long
    Dim lpS As Long
    Dim lpD As Long
    
    ConvColorGdiplusToGdi = nColor And &HFFFFFF
    lpS = VarPtr(nColor)
    lpD = VarPtr(ConvColorGdiplusToGdi)
    CopyMemory lpD, lpS + 2, 1
    CopyMemory lpD + 2, lpS, 1
End Function

Public Function ReadTextFromFile(ByVal FP As String, Str As String) As Boolean
    Dim nType As Long
    
    ReadTextFromFile = ReadTextFromFileRT(FP, Str, nType)
End Function

Public Function ReadTextFromFileRT(ByVal FP As String, Str As String, EncodeType As Long) As Boolean
    Dim hFile As Long
    Dim Bin() As Byte
    Dim nSize As Long
    Dim Ret As Long
    
    ReadTextFromFileRT = False
    hFile = CreateFileW(StrPtr(FP), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        nSize = GetFileSize(hFile, Ret)
        If nSize > 0 Then
            ReDim Bin(nSize - 1)
            ReadFile hFile, VarPtr(Bin(0)), nSize, Ret, 0
            EncodeType = GetByteOrderMark(Bin)
            Select Case EncodeType
                Case ByteOrderMark_UNICODE
                    nSize = nSize - 2
                    Str = String$(nSize \ 2, vbNullChar)
                    CopyMemory StrPtr(Str), VarPtr(Bin(2)), nSize
                    
                Case ByteOrderMark_UTF8
                    Str = ConvUtf8ToUnicode(Bin, nSize)
                
                Case Else
                    Str = StrConv(Bin, vbUnicode)
                    
            End Select
            ReadTextFromFileRT = True
        End If
        CloseHandle hFile
    End If
    
    Erase Bin
End Function

Private Function GetByteOrderMark(Bin() As Byte) As Long
    If Bin(0) = &HFF Then
        If Bin(1) = &HFE Then
            GetByteOrderMark = ByteOrderMark_UNICODE
            Exit Function
        End If
    End If
    If Bin(0) = &HEF Then
        If Bin(1) = &HBB Then
            If Bin(2) = &HBF Then
                GetByteOrderMark = ByteOrderMark_UTF8
                Exit Function
            End If
        End If
    End If
    GetByteOrderMark = ByteOrderMark_ANSI
End Function

Public Function ConvUtf8ToUnicode(BinaryData() As Byte, ByVal BinarySize As Long) As String
    Dim Ret As Long
    Dim BaseAddr As Long
    
    ConvUtf8ToUnicode = ""
    
    BaseAddr = 0
    If UBound(BinaryData) >= 2 Then
        If BinaryData(0) = &HEF Then
            If BinaryData(1) = &HBB Then
                If BinaryData(2) = &HBF Then
                    BaseAddr = 3
                    BinarySize = BinarySize - 3
                End If
            End If
        End If
    End If
    
    Ret = MultiByteToWideChar(CP_UTF8, 0, BinaryData(BaseAddr), BinarySize, 0, 0)
    If Ret > 0 Then
        ConvUtf8ToUnicode = String$(Ret, vbNullChar)
        MultiByteToWideChar CP_UTF8, 0, BinaryData(BaseAddr), BinarySize, StrPtr(ConvUtf8ToUnicode), Ret
    End If
End Function

Public Function ReadUtf8TextFromFile(ByVal FP As String, Str As String) As Boolean
    Dim hFile As Long
    Dim Bin() As Byte
    Dim nSize As Long
    Dim Ret As Long
    
    ReadUtf8TextFromFile = False
    hFile = CreateFileW(StrPtr(FP), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        nSize = GetFileSize(hFile, Ret)
        If nSize > 0 Then
            ReDim Bin(nSize - 1)
            ReadFile hFile, VarPtr(Bin(0)), nSize, Ret, 0
            Str = ConvUtf8ToUnicode(Bin, nSize)
            ReadUtf8TextFromFile = True
        End If
        CloseHandle hFile
    End If
    
    Erase Bin
End Function


Public Function WriteAnsiTextToFile(ByVal FP As String, ByVal S As String) As Boolean
    Dim hFile As Long
    Dim Ret As Long
    Dim Bin() As Byte
    
    WriteAnsiTextToFile = False
    
    DeleteFileW StrPtr(FP)
    hFile = CreateFileW(StrPtr(FP), GENERIC_WRITE, 0, 0, CREATE_NEW, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        If S <> "" Then
            Bin = StrConv(S, vbFromUnicode)
            WriteFile hFile, VarPtr(Bin(0)), UBound(Bin) + 1, Ret, 0
        End If
        CloseHandle hFile
        
        WriteAnsiTextToFile = True
    End If
    
    Erase Bin
End Function

Public Function WriteUnicodeTextToFile(ByVal FP As String, ByVal S As String, ByVal IsBom As Boolean) As Boolean
    Dim hFile As Long
    Dim Ret As Long
    
    WriteUnicodeTextToFile = False
    
    DeleteFileW StrPtr(FP)
    hFile = CreateFileW(StrPtr(FP), GENERIC_WRITE, 0, 0, CREATE_NEW, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        If "" <> S Then
            If IsBom Then
                S = ChrW$(&HFEFF) + S
            End If
            WriteFile hFile, StrPtr(S), Len(S) * 2, Ret, 0
        End If
        CloseHandle hFile
        
        WriteUnicodeTextToFile = True
    End If
End Function

Public Function WriteUtf8TextToFile(ByVal FP As String, ByVal S As String, ByVal IsBom As Boolean) As Boolean
    Dim hFile As Long
    Dim Ret As Long
    Dim Bin() As Byte
    Dim L As Long
    
    WriteUtf8TextToFile = False
    
    DeleteFileW StrPtr(FP)
    hFile = CreateFileW(StrPtr(FP), GENERIC_WRITE, 0, 0, CREATE_NEW, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        If S <> "" Then
            Ret = Len(S) * 3 + 3
            ReDim Bin(Ret - 1)
            L = WideCharToMultiByte(CP_UTF8, 0, StrPtr(S), Len(S), Bin(3), Ret, 0, 0)
            If IsBom Then
                Bin(0) = &HEF
                Bin(1) = &HBB
                Bin(2) = &HBF
                WriteFile hFile, VarPtr(Bin(0)), L + 3, Ret, 0
            Else
                WriteFile hFile, VarPtr(Bin(3)), L, Ret, 0
            End If
        End If
        CloseHandle hFile
        
        WriteUtf8TextToFile = True
    End If
    
    Erase Bin
End Function

Public Sub GetMousePosition(ByVal hWnd As Long, retX As Long, retY As Long)
    Dim pt As POINTAPI
    
    GetCursorPos pt
    ScreenToClient hWnd, pt
    retX = pt.nX
    retY = pt.nY
End Sub

Public Function GetRegValueString(ByVal hKey As Long, ByVal ValueName As String) As String
    Dim nType As Long
    Dim Ret As Long
    
    GetRegValueString = String$(MAX_PATH, vbNullChar)
    Ret = MAX_PATH * 2
    If ERROR_SUCCESS = RegQueryValueExW(hKey, StrPtr(ValueName), 0, nType, StrPtr(GetRegValueString), Ret) Then
        GetRegValueString = StrCutNull(GetRegValueString)
    Else
        GetRegValueString = ""
    End If
End Function

Public Function SetRegValueString(ByVal hKey As Long, ByVal ValueName As String, ByVal Str As String) As Boolean
    Dim nLen As Long
    
    Str = Str + vbNullChar
    nLen = Len(Str) * 2
    SetRegValueString = (ERROR_SUCCESS = RegSetValueExW(hKey, StrPtr(ValueName), 0, REG_SZ, StrPtr(Str), nLen))
End Function

Public Function GetRegValueLong(ByVal hKey As Long, ByVal ValueName As String) As Long
    Dim nType As Long
    Dim Ret As Long

    GetRegValueLong = 0
    Ret = 4
    RegQueryValueExW hKey, StrPtr(ValueName), 0, nType, VarPtr(GetRegValueLong), Ret
End Function

Public Function SetRegValueLong(ByVal hKey As Long, ByVal ValueName As String, ByVal V As Long) As Boolean
    SetRegValueLong = (ERROR_SUCCESS = RegSetValueExW(hKey, StrPtr(ValueName), 0, REG_DWORD, VarPtr(V), 4))
End Function

Public Function LoadFileLastWriteTime(ByVal FP As String, FT As FILETIME) As Boolean
    Dim hFile As Long
    
    LoadFileLastWriteTime = False
    
    hFile = CreateFileW(StrPtr(FP), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If INVALID_HANDLE_VALUE <> hFile Then
        GetFileTime hFile, 0, 0, VarPtr(FT)
        CloseHandle hFile
        
        LoadFileLastWriteTime = True
    End If
End Function

