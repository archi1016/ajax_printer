Attribute VB_Name = "api_window"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32" ()

Public Declare Function InitCommonControlsEx Lib "comctl32" _
    (lpInitCtrls As InitCommonControlsEx) As Long

Public Type InitCommonControlsEx
    dwSize As Long
    dwICC As Long
End Type

Public Const ICC_LISTVIEW_CLASSES = &H1           'listview,   header
Public Const ICC_TREEVIEW_CLASSES = &H2           'treeview,   tooltips
Public Const ICC_BAR_CLASSES = &H4                     'toolbar,   statusbar,   trackbar,   tooltips
Public Const ICC_TAB_CLASSES = &H8                     'tab,   tooltips
Public Const ICC_UPDOWN_CLASS = &H10                 'updown
Public Const ICC_PROGRESS_CLASS = &H20             'progress
Public Const ICC_HOTKEY_CLASS = &H40                 'hotkey
Public Const ICC_ANIMATE_CLASS = &H80               'animate
Public Const ICC_WIN95_CLASSES = &HFF               'loads   everything   above
Public Const ICC_DATE_CLASSES = &H100               'month   picker,   date   picker,   time   picker,   updown
Public Const ICC_USEREX_CLASSES = &H200           'ComboEx
Public Const ICC_COOL_CLASSES = &H400               'Rebar   =coolbar   control

Public Declare Function SendMessageW Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal Msg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long
     
Public Declare Function CreateWindowExW Lib "user32" _
    (ByVal dwExStyle As Long, _
     ByVal lpClassName As Long, _
     ByVal lpWindowName As Long, _
     ByVal dwStyle As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hWndParent As Long, _
     ByVal hMenu As Long, _
     ByVal hInstance As Long, _
     ByVal lpParam As Long) As Long
     
Public Declare Function DestroyWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Const CCS_TOP = &H1
Public Const CCS_NOMOVEY = &H2
Public Const CCS_BOTTOM = &H3
Public Const CCS_NORESIZE = &H4
Public Const CCS_NOPARENTALIGN = &H8
Public Const CCS_NODIVIDER = &H40

Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_DLGMODALFRAME = &H1
Public Const WS_EX_MDICHILD = &H40
Public Const WS_EX_LEFT = &H0
Public Const WS_EX_LEFTSCROLLBAR = &H4000
Public Const WS_EX_LTRREADING = &H0
Public Const WS_EX_NOPARENTNOTIFY = &H4
Public Const WS_EX_RIGHT = &H1000
Public Const WS_EX_RIGHTSCROLLBAR = &H0
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_TOPMOST = &H8
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_ACCEPTFILES = &H10
Public Const WS_EX_RTLREADING = &H2000
Public Const WS_EX_TRANSPARENT = &H20
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_OVERLAPPED = &H0
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX) '標準的なスタイル
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) 'ポップアップ型の標準的なスタイル
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_TILED = WS_OVERLAPPED
Public Const CW_USEDEFAULT = &H80000000
Public Const BS_PUSHBUTTON = &H0

Public Const CCM_FIRST = &H2000&
Public Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Public Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)
Public Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)
Public Const CCM_GETDROPTARGET = (CCM_FIRST + 4)
Public Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
Public Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
Public Const CCM_SETVERSION = (CCM_FIRST + 7)
Public Const CCM_GETVERSION = (CCM_FIRST + 8)
Public Const CCM_SETNOTIFYWINDOW = (CCM_FIRST + 9)

Public Declare Function MoveWindow Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal bRepaint As Long) As Long

Public Declare Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal nCmdShow As Long) As Long
          
Public Declare Function SetActiveWindow Lib "user32" _
    (ByVal hWnd As Long) As Long
                    
Public Declare Function SetFocusW Lib "user32" Alias "SetFocus" _
    (ByVal hWnd As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long


Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long) As Long
     
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" _
    (ByVal hWnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long
    
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" _
    (ByVal lpPrevWndFunc As Long, _
     ByVal hWnd As Long, _
     ByVal Msg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = -4
Public Const GWL_HWNDPARENT = -8
Public Const GWL_EXSTYLE = -20
Public Const GWL_STYLE = -16


Public Declare Function SetWindowPos Lib "user32" _
   (ByVal hWnd As Long, _
   ByVal hWndInsertAfter As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal wFlags As Long) As Long
   
Public Const HWND_BOTTOM = 1
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME




Public Declare Function FindWindowW Lib "user32" _
    (ByVal lpClassName As Long, _
     ByVal lpWindowName As Long) As Long

Public Declare Function EnumWindows Lib "user32" _
    (ByVal lpEnumFunc As Long, _
     ByVal lParam As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function GetWindowTextLengthW Lib "user32" _
    (ByVal hWnd As Long) As Long
     
Public Declare Function GetWindowTextW Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal lpString As Long, _
     ByVal cch As Long) As Long

Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, _
      lpRect As RECT) As Long


Public Declare Function GetOpenFileNameW Lib "comdlg32" _
    (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetSaveFileNameW Lib "comdlg32" _
    (pOpenfilename As OPENFILENAME) As Long
    
Public Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As Long
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_FORCESHOWHIDDEN As Long = &H10000000
Public Const OFN_DONTADDTORECENT As Long = &H2000000
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_ENABLESIZING As Long = &H800000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_HIDEREADONLY As Long = &H4



Public Declare Function ChooseColorW Lib "comdlg32" _
    (Color As ChooseColor) As Long


Public Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_SHOWHELP = &H8
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_ANYCOLOR = &H100



Public Declare Function EnableWindow Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal bEnable As Long) As Long


