Attribute VB_Name = "api_shell32"
Option Explicit

Public Declare Function ShellExecuteW Lib "shell32" _
    (ByVal hWnd As Long, _
     ByVal lpOperation As Long, _
     ByVal lpFile As Long, _
     ByVal lpParameters As Long, _
     ByVal lpDirectory As Long, _
     ByVal nShowCmd As Long) As Long

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As Long
    lpFile As Long
    lpParameters As Long
    lpDirectory As Long
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As Long
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Const SEE_MASK_CLASSNAME = &H1
Public Const SEE_MASK_CLASSKEY = &H3
Public Const SEE_MASK_IDLIST = &H4
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_ICON = &H10
Public Const SEE_MASK_HOTKEY = &H20
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_CONNECTNETDRV = &H80
Public Const SEE_MASK_FLAG_DDEWAIT = &H100
Public Const SEE_MASK_DOENVSUBST = &H200
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const SEE_MASK_UNICODE = &H4000
Public Const SEE_MASK_NO_CONSOLE = &H8000
Public Const SEE_MASK_ASYNCOK = &H100000
Public Const SEE_MASK_HMONITOR = &H200000
Public Const SEE_MASK_NOZONECHECKS = &H800000

Public Declare Function ShellExecuteExW Lib "shell32" _
    (lpExecInfo As SHELLEXECUTEINFO) As Long

Public Declare Function SHGetFolderPathW Lib "shell32" _
    (ByVal hWndOwner As Long, _
     ByVal nFolder As Long, _
     ByVal hToken As Long, _
     ByVal dwFlags As Long, _
     ByVal lppszPath As Long) As Long
     
Public Declare Function SHGetSpecialFolderPathW Lib "shell32" _
    (ByVal hWndOwner As Long, _
     ByVal lpszPath As Long, _
     ByVal csidl As Long, _
     ByVal fCreate As Long) As Long

Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_LOCAL_APPDATA = &H1C
Public Const CSIDL_COMMON_APPDATA = &H23
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_COMMON_DOCUMENTS = &H2E
Public Const CSIDL_PROGRAM_FILES = &H26
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_PROGRAM_FILESX86 = &H2A
Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C


Public Declare Function SHFileOperationW Lib "shell32" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
     
Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4
     
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_SILENT = &H4
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_WANTMAPPINGHANDLE = &H20
Public Const FOF_CREATEPROGRESSDLG = &H0
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_FILESONLY = &H80
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_NOERRORUI = &H400

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As Long
    pTo As Long
    fFlags As Integer
    fAborted As Long
    hNameMaps As Long
    lpszProgressTitle As Long
End Type


Public Declare Function SHCreateDirectory Lib "shell32" _
    (ByVal hWnd As Long, _
     ByVal pszPath As Long) As Long

Public Declare Function SHBrowseForFolderW Lib "shell32" _
    (lpbi As BROWSEINFO) As Long

Public Type BROWSEINFO
    hWndOwner As Long
    pidlRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_EDITBOX = &H10
Public Const BIF_VALIDATE = &H20
Public Const BIF_NEWDIALOGSTYLE = &H40
Public Const BIF_BROWSEINCLUDEURLS = &H80
Public Const BIF_USENEWUI = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
Public Const BIF_UAHINT = &H100
Public Const BIF_NONEWFOLDERBUTTON = &H200
Public Const BIF_NOTRANSLATETARGETS = &H400
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const BIF_BROWSEINCLUDEFILES = &H4000
Public Const BIF_SHAREABLE As Long = &H8000


Public Declare Function CoTaskMemFree Lib "ole32" _
    (ByVal pv As Long) As Long

Public Declare Function SHGetPathFromIDListW Lib "shell32" _
    (ByVal pidl As Long, _
     ByVal pszPath As Long) As Long



Public Declare Sub DragAcceptFiles Lib "shell32" _
    (ByVal hWnd As Long, _
     ByVal fAccept As Long)

Public Declare Sub DragFinish Lib "shell32" _
    (ByVal hDrop As Long)

Public Declare Function DragQueryFileW Lib "shell32" _
    (ByVal hDrop As Long, _
     ByVal iFile As Long, _
     ByVal lpszFile As Long, _
     ByVal cch As Long) As Long

Public Declare Function DragQueryPoint Lib "shell32" _
    (ByVal hDrop As Long, _
     lppt As POINTAPI) As Long

