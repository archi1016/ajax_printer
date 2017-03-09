Attribute VB_Name = "api_user32"
Option Explicit

Public Declare Function MessageBoxW Lib "user32.dll" _
    (ByVal hWnd As Long, _
     ByVal lpText As Long, _
     ByVal lpCaption As Long, _
     ByVal uType As Long) As Long

Public Const MB_OK = &H0
Public Const MB_OKCANCEL = &H1
Public Const MB_ABORTRETRYIGNORE = &H2
Public Const MB_YESNOCANCEL = &H3
Public Const MB_YESNO = &H4
Public Const MB_RETRYCANCEL = &H5
Public Const MB_CANCELTRYCONTINUE = &H6

Public Const MB_ICONHAND = &H10
Public Const MB_ICONQUESTION = &H20
Public Const MB_ICONEXCLAMATION = &H30
Public Const MB_ICONASTERISK = &H40
Public Const MB_USERICON = &H80
Public Const MB_ICONWARNING = MB_ICONEXCLAMATION
Public Const MB_ICONERROR = MB_ICONHAND
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONSTOP = MB_ICONHAND

Public Const MB_DEFBUTTON1 = &H0
Public Const MB_DEFBUTTON2 = &H100
Public Const MB_DEFBUTTON3 = &H200
Public Const MB_DEFBUTTON4 = &H300

Public Const MB_APPLMODAL = &H0
Public Const MB_SYSTEMMODAL = &H1000
Public Const MB_TASKMODAL = &H2000

Public Const MB_HELP = &H4000
Public Const MB_NOFOCUS = &H8000
Public Const MB_SETFOREGROUND = &H10000
Public Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Public Const MB_TOPMOST = &H40000
Public Const MB_RIGHT = &H80000
Public Const MB_RTLREADING = &H100000
Public Const MB_SERVICE_NOTIFICATION = &H200000

Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7
Public Const IDCLOSE = 8
Public Const IDHELP = 9



Public Declare Function RegisterWindowMessageW Lib "user32.dll" _
    (ByVal lpString As Long) As Long
    
Public Declare Function GetParent Lib "user32" _
    (ByVal hWnd As Long) As Long
    
Public Declare Function GetSysColor Lib "user32" _
    (ByVal nIndex As Long) As Long

'Public Const CTLCOLOR_MSGBOX = 0
'Public Const CTLCOLOR_EDIT = 1
'Public Const CTLCOLOR_LISTBOX = 2
'Public Const CTLCOLOR_BTN = 3
'Public Const CTLCOLOR_DLG = 4
'Public Const CTLCOLOR_SCROLLBAR = 5
'Public Const CTLCOLOR_STATIC = 6
'Public Const CTLCOLOR_MAX = 7

Public Const COLORS_SCROLLBAR = 0
Public Const COLORS_BACKGROUND = 1
Public Const COLORS_ACTIVECAPTION = 2
Public Const COLORS_INACTIVECAPTION = 3
Public Const COLORS_MENU = 4
Public Const COLORS_WINDOW = 5
Public Const COLORS_WINDOWFRAME = 6
Public Const COLORS_MENUTEXT = 7
Public Const COLORS_WINDOWTEXT = 8
Public Const COLORS_CAPTIONTEXT = 9
Public Const COLORS_ACTIVEBORDER = 10
Public Const COLORS_INACTIVEBORDER = 11
Public Const COLORS_APPWORKSPACE = 12
Public Const COLORS_HIGHLIGHT = 13
Public Const COLORS_HIGHLIGHTTEXT = 14
Public Const COLORS_BTNFACE = 15
Public Const COLORS_BTNSHADOW = 16
Public Const COLORS_GRAYTEXT = 17
Public Const COLORS_BTNTEXT = 18
Public Const COLORS_INACTIVECAPTIONTEXT = 19
Public Const COLORS_BTNHIGHLIGHT = 20
Public Const COLORS_3DDKSHADOW = 21
Public Const COLORS_3DLIGHT = 22
Public Const COLORS_INFOTEXT = 23
Public Const COLORS_INFOBK = 24
Public Const COLORS_GRADIENTACTIVECAPTION = 27
Public Const COLORS_GRADIENTINACTIVECAPTION = 28
Public Const COLORS_MENUHILIGHT = 29
Public Const COLORS_MENUBAR = 30

Public Declare Function PrivateExtractIconsW Lib "user32" _
    (ByVal lpszFile As Long, _
     ByVal nIconIndex As Long, _
     ByVal cxIcon As Long, _
     ByVal cyIcon As Long, _
     phicon As Long, _
     piconid As Long, _
     ByVal nIcons As Long, _
     ByVal flags As Long) As Long
     
Public Declare Function CreateIcon Lib "user32" _
    (ByVal hInstance As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal cPlanes As Long, _
     ByVal cBitsPixel As Long, _
     ByVal lpbANDbits As Long, _
     ByVal lpbXORbits As Long) As Long
          
Public Declare Function DestroyIcon Lib "user32" _
    (ByVal hIcon As Long) As Long

Public Declare Function DrawIconEx Lib "user32" _
    (ByVal hDC As Long, _
     ByVal xLeft As Long, _
     ByVal yTop As Long, _
     ByVal hIcon As Long, _
     ByVal cxWidth As Long, _
     ByVal cyWidth As Long, _
     ByVal istepIfAniCur As Long, _
     ByVal hbrFickerFreeDraw As Long, _
     ByVal diFlags As Long) As Long

Public Const DI_COMPAT = 4
Public Const DI_DEFAULTSIZE = 8
Public Const DI_IMAGE = 2
Public Const DI_MASK = 1
Public Const DI_NORMAL = 3
Public Const DI_APPBANDING = 1


Public Declare Function EnumDisplaySettingsW Lib "user32" _
    (ByVal lpszDeviceName As Long, _
     ByVal iModeNum As Long, _
     lpDevMode As DEVMODE) As Long

Public Declare Function ChangeDisplaySettingsW Lib "user32" _
    (lpDevMode As DEVMODE, _
     ByVal dwFlags As Long) As Long
     
Public Const ENUM_CURRENT_SETTINGS = -1
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Type DEVMODE
    dmDeviceName(CCDEVICENAME * 2 - 1) As Byte
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName(CCHFORMNAME * 2 - 1) As Byte
    'dmLogPixels As Long
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    
    'dmICMMethod As Long
    'dmICMIntent As Long
    'dmMediaType As Long
    'dmDitherType As Long
    'dmReserved1 As Long
    'dmReserved2 As Long
    'dmPanningWidth As Long
    'dmPanningHeight As Long
End Type

Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFREQUENCY = &H400000

Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H2

Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Const BITSPIXEL = 12

Public Declare Function SystemParametersInfoW Lib "user32" _
    (ByVal uAction As Long, _
     ByVal uParam As Long, _
     ByVal lpvParam As Long, _
     ByVal fuWinIni As Long) As Long

Public Const SPI_GETBEEP = &H1
Public Const SPI_SETBEEP = &H2
Public Const SPI_GETMOUSE = &H3
Public Const SPI_SETMOUSE = &H4
Public Const SPI_GETBORDER = &H5
Public Const SPI_SETBORDER = &H6
Public Const SPI_GETKEYBOARDSPEED = &HA
Public Const SPI_SETKEYBOARDSPEED = &HB
Public Const SPI_LANGDRIVER = &HC
Public Const SPI_ICONHORIZONTALSPACING = &HD
Public Const SPI_GETSCREENSAVETIMEOUT = &HE
Public Const SPI_SETSCREENSAVETIMEOUT = &HF
Public Const SPI_GETSCREENSAVEACTIVE = &H10
Public Const SPI_SETSCREENSAVEACTIVE = &H11
Public Const SPI_GETGRIDGRANULARITY = &H12
Public Const SPI_SETGRIDGRANULARITY = &H13
Public Const SPI_SETDESKWALLPAPER = &H14
Public Const SPI_SETDESKPATTERN = &H15
Public Const SPI_GETKEYBOARDDELAY = &H16
Public Const SPI_SETKEYBOARDDELAY = &H17
Public Const SPI_ICONVERTICALSPACING = &H18
Public Const SPI_GETICONTITLEWRAP = &H19
Public Const SPI_SETICONTITLEWRAP = &H1A
Public Const SPI_GETMENUDROPALIGNMENT = &H1B
Public Const SPI_SETMENUDROPALIGNMENT = &H1C
Public Const SPI_SETDOUBLECLKWIDTH = &H1D
Public Const SPI_SETDOUBLECLKHEIGHT = &H1E
Public Const SPI_GETICONTITLELOGFONT = &H1F
Public Const SPI_SETDOUBLECLICKTIME = &H20
Public Const SPI_SETMOUSEBUTTONSWAP = &H21
Public Const SPI_SETICONTITLELOGFONT = &H22
Public Const SPI_GETFASTTASKSWITCH = &H23
Public Const SPI_SETFASTTASKSWITCH = &H24
Public Const SPI_SETDRAGFULLWINDOWS = &H25
Public Const SPI_GETDRAGFULLWINDOWS = &H26
Public Const SPI_GETNONCLIENTMETRICS = &H29
Public Const SPI_SETNONCLIENTMETRICS = &H2A
Public Const SPI_GETMINIMIZEDMETRICS = &H2B
Public Const SPI_SETMINIMIZEDMETRICS = &H2C
Public Const SPI_GETICONMETRICS = &H2D
Public Const SPI_SETICONMETRICS = &H2E
Public Const SPI_SETWORKAREA = &H2F
Public Const SPI_GETWORKAREA = &H30
Public Const SPI_SETPENWINDOWS = &H31
Public Const SPI_GETHIGHCONTRAST = &H42
Public Const SPI_SETHIGHCONTRAST = &H43
Public Const SPI_GETKEYBOARDPREF = &H44
Public Const SPI_SETKEYBOARDPREF = &H45
Public Const SPI_GETSCREENREADER = &H46
Public Const SPI_SETSCREENREADER = &H47
Public Const SPI_GETANIMATION = &H48
Public Const SPI_SETANIMATION = &H49
Public Const SPI_GETFONTSMOOTHING = &H4A
Public Const SPI_SETFONTSMOOTHING = &H4B
Public Const SPI_SETDRAGWIDTH = &H4C
Public Const SPI_SETDRAGHEIGHT = &H4D
Public Const SPI_SETHANDHELD = &H4E
Public Const SPI_GETLOWPOWERTIMEOUT = &H4F
Public Const SPI_GETPOWEROFFTIMEOUT = &H50
Public Const SPI_SETLOWPOWERTIMEOUT = &H51
Public Const SPI_SETPOWEROFFTIMEOUT = &H52
Public Const SPI_GETLOWPOWERACTIVE = &H53
Public Const SPI_GETPOWEROFFACTIVE = &H54
Public Const SPI_SETLOWPOWERACTIVE = &H55
Public Const SPI_SETPOWEROFFACTIVE = &H56
Public Const SPI_SETCURSORS = &H57
Public Const SPI_SETICONS = &H58
Public Const SPI_GETDEFAULTINPUTLANG = &H59
Public Const SPI_SETDEFAULTINPUTLANG = &H5A
Public Const SPI_SETLANGTOGGLE = &H5B
Public Const SPI_GETWINDOWSEXTENSION = &H5C
Public Const SPI_SETMOUSETRAILS = &H5D
Public Const SPI_GETMOUSETRAILS = &H5E
Public Const SPI_SETSCREENSAVERRUNNING = &H61
Public Const SPI_SCREENSAVERRUNNING = SPI_SETSCREENSAVERRUNNING
Public Const SPI_GETFILTERKEYS = &H32
Public Const SPI_SETFILTERKEYS = &H33
Public Const SPI_GETTOGGLEKEYS = &H34
Public Const SPI_SETTOGGLEKEYS = &H35
Public Const SPI_GETMOUSEKEYS = &H36
Public Const SPI_SETMOUSEKEYS = &H37
Public Const SPI_GETSHOWSOUNDS = &H38
Public Const SPI_SETSHOWSOUNDS = &H39
Public Const SPI_GETSTICKYKEYS = &H3A
Public Const SPI_SETSTICKYKEYS = &H3B
Public Const SPI_GETACCESSTIMEOUT = &H3C
Public Const SPI_SETACCESSTIMEOUT = &H3D
Public Const SPI_GETSERIALKEYS = &H3E
Public Const SPI_SETSERIALKEYS = &H3F
Public Const SPI_GETSOUNDSENTRY = &H40
Public Const SPI_SETSOUNDSENTRY = &H41
Public Const SPI_GETSNAPTODEFBUTTON = &H5F
Public Const SPI_SETSNAPTODEFBUTTON = &H60
Public Const SPI_GETMOUSEHOVERWIDTH = &H62
Public Const SPI_SETMOUSEHOVERWIDTH = &H63
Public Const SPI_GETMOUSEHOVERHEIGHT = &H64
Public Const SPI_SETMOUSEHOVERHEIGHT = &H65
Public Const SPI_GETMOUSEHOVERTIME = &H66
Public Const SPI_SETMOUSEHOVERTIME = &H67
Public Const SPI_GETWHEELSCROLLLINES = &H68
Public Const SPI_SETWHEELSCROLLLINES = &H69
Public Const SPI_GETMENUSHOWDELAY = &H6A
Public Const SPI_SETMENUSHOWDELAY = &H6B
Public Const SPI_GETWHEELSCROLLCHARS = &H6C
Public Const SPI_SETWHEELSCROLLCHARS = &H6D
Public Const SPI_GETSHOWIMEUI = &H6E
Public Const SPI_SETSHOWIMEUI = &H6F
Public Const SPI_GETMOUSESPEED = &H70
Public Const SPI_SETMOUSESPEED = &H71
Public Const SPI_GETSCREENSAVERRUNNING = &H72
Public Const SPI_GETDESKWALLPAPER = &H73
Public Const SPI_GETAUDIODESCRIPTION = &H74
Public Const SPI_SETAUDIODESCRIPTION = &H75
Public Const SPI_GETSCREENSAVESECURE = &H76
Public Const SPI_SETSCREENSAVESECURE = &H77
Public Const SPI_GETHUNGAPPTIMEOUT = &H78
Public Const SPI_SETHUNGAPPTIMEOUT = &H79
Public Const SPI_GETWAITTOKILLTIMEOUT = &H7A
Public Const SPI_SETWAITTOKILLTIMEOUT = &H7B
Public Const SPI_GETWAITTOKILLSERVICETIMEOUT = &H7C
Public Const SPI_SETWAITTOKILLSERVICETIMEOUT = &H7D
Public Const SPI_GETMOUSEDOCKTHRESHOLD = &H7E
Public Const SPI_SETMOUSEDOCKTHRESHOLD = &H7F
Public Const SPI_GETPENDOCKTHRESHOLD = &H80
Public Const SPI_SETPENDOCKTHRESHOLD = &H81
Public Const SPI_GETWINARRANGING = &H82
Public Const SPI_SETWINARRANGING = &H83
Public Const SPI_GETMOUSEDRAGOUTTHRESHOLD = &H84
Public Const SPI_SETMOUSEDRAGOUTTHRESHOLD = &H85
Public Const SPI_GETPENDRAGOUTTHRESHOLD = &H86
Public Const SPI_SETPENDRAGOUTTHRESHOLD = &H87
Public Const SPI_GETMOUSESIDEMOVETHRESHOLD = &H88
Public Const SPI_SETMOUSESIDEMOVETHRESHOLD = &H89
Public Const SPI_GETPENSIDEMOVETHRESHOLD = &H8A
Public Const SPI_SETPENSIDEMOVETHRESHOLD = &H8B
Public Const SPI_GETDRAGFROMMAXIMIZE = &H8C
Public Const SPI_SETDRAGFROMMAXIMIZE = &H8D
Public Const SPI_GETSNAPSIZING = &H8E
Public Const SPI_SETSNAPSIZING = &H8F
Public Const SPI_GETDOCKMOVING = &H90
Public Const SPI_SETDOCKMOVING = &H91
Public Const SPI_GETACTIVEWINDOWTRACKING = &H1000
Public Const SPI_SETACTIVEWINDOWTRACKING = &H1001
Public Const SPI_GETMENUANIMATION = &H1002
Public Const SPI_SETMENUANIMATION = &H1003
Public Const SPI_GETCOMBOBOXANIMATION = &H1004
Public Const SPI_SETCOMBOBOXANIMATION = &H1005
Public Const SPI_GETLISTBOXSMOOTHSCROLLING = &H1006
Public Const SPI_SETLISTBOXSMOOTHSCROLLING = &H1007
Public Const SPI_GETGRADIENTCAPTIONS = &H1008
Public Const SPI_SETGRADIENTCAPTIONS = &H1009
Public Const SPI_GETKEYBOARDCUES = &H100A
Public Const SPI_SETKEYBOARDCUES = &H100B
Public Const SPI_GETMENUUNDERLINES = SPI_GETKEYBOARDCUES
Public Const SPI_SETMENUUNDERLINES = SPI_SETKEYBOARDCUES
Public Const SPI_GETACTIVEWNDTRKZORDER = &H100C
Public Const SPI_SETACTIVEWNDTRKZORDER = &H100D
Public Const SPI_GETHOTTRACKING = &H100E
Public Const SPI_SETHOTTRACKING = &H100F
Public Const SPI_GETMENUFADE = &H1012
Public Const SPI_SETMENUFADE = &H1013
Public Const SPI_GETSELECTIONFADE = &H1014
Public Const SPI_SETSELECTIONFADE = &H1015
Public Const SPI_GETTOOLTIPANIMATION = &H1016
Public Const SPI_SETTOOLTIPANIMATION = &H1017
Public Const SPI_GETTOOLTIPFADE = &H1018
Public Const SPI_SETTOOLTIPFADE = &H1019
Public Const SPI_GETCURSORSHADOW = &H101A
Public Const SPI_SETCURSORSHADOW = &H101B
Public Const SPI_GETMOUSESONAR = &H101C
Public Const SPI_SETMOUSESONAR = &H101D
Public Const SPI_GETMOUSECLICKLOCK = &H101E
Public Const SPI_SETMOUSECLICKLOCK = &H101F
Public Const SPI_GETMOUSEVANISH = &H1020
Public Const SPI_SETMOUSEVANISH = &H1021
Public Const SPI_GETFLATMENU = &H1022
Public Const SPI_SETFLATMENU = &H1023
Public Const SPI_GETDROPSHADOW = &H1024
Public Const SPI_SETDROPSHADOW = &H1025
Public Const SPI_GETBLOCKSENDINPUTRESETS = &H1026
Public Const SPI_SETBLOCKSENDINPUTRESETS = &H1027
Public Const SPI_GETUIEFFECTS = &H103E
Public Const SPI_SETUIEFFECTS = &H103F
Public Const SPI_GETDISABLEOVERLAPPEDCONTENT = &H1040
Public Const SPI_SETDISABLEOVERLAPPEDCONTENT = &H1041
Public Const SPI_GETCLIENTAREAANIMATION = &H1042
Public Const SPI_SETCLIENTAREAANIMATION = &H1043
Public Const SPI_GETCLEARTYPE = &H1048
Public Const SPI_SETCLEARTYPE = &H1049
Public Const SPI_GETSPEECHRECOGNITION = &H104A
Public Const SPI_SETSPEECHRECOGNITION = &H104B
Public Const SPI_GETFOREGROUNDLOCKTIMEOUT = &H2000
Public Const SPI_SETFOREGROUNDLOCKTIMEOUT = &H2001
Public Const SPI_GETACTIVEWNDTRKTIMEOUT = &H2002
Public Const SPI_SETACTIVEWNDTRKTIMEOUT = &H2003
Public Const SPI_GETFOREGROUNDFLASHCOUNT = &H2004
Public Const SPI_SETFOREGROUNDFLASHCOUNT = &H2005
Public Const SPI_GETCARETWIDTH = &H2006
Public Const SPI_SETCARETWIDTH = &H2007
Public Const SPI_GETMOUSECLICKLOCKTIME = &H2008
Public Const SPI_SETMOUSECLICKLOCKTIME = &H2009
Public Const SPI_GETFONTSMOOTHINGTYPE = &H200A
Public Const SPI_SETFONTSMOOTHINGTYPE = &H200B
Public Const SPI_GETFONTSMOOTHINGCONTRAST = &H200C
Public Const SPI_SETFONTSMOOTHINGCONTRAST = &H200D
Public Const SPI_GETFOCUSBORDERWIDTH = &H200E
Public Const SPI_SETFOCUSBORDERWIDTH = &H200F
Public Const SPI_GETFOCUSBORDERHEIGHT = &H2010
Public Const SPI_SETFOCUSBORDERHEIGHT = &H2011
Public Const SPI_GETFONTSMOOTHINGORIENTATION = &H2012
Public Const SPI_SETFONTSMOOTHINGORIENTATION = &H2013
Public Const SPI_GETMINIMUMHITRADIUS = &H2014
Public Const SPI_SETMINIMUMHITRADIUS = &H2015
Public Const SPI_GETMESSAGEDURATION = &H2016
Public Const SPI_SETMESSAGEDURATION = &H2017

Public Const SPIF_SENDWININICHANGE = &H2
Public Const SPIF_UPDATEINIFILE = &H1

Public Const WHEEL_PAGESCROLL = &HFFFFFFFF


Public Declare Function GetCursorPos Lib "user32" _
    (lpPoint As POINTAPI) As Long
    
Public Declare Function SetCapture Lib "user32" _
    (ByVal hWnd As Long) As Long
    
Public Declare Function ReleaseCapture Lib "user32" () As Long
    
Public Declare Function WindowFromPoint Lib "user32" _
    (ByVal nX As Long, _
     ByVal nY As Long) As Long
     
Public Declare Function ClientToScreen Lib "user32" _
    (ByVal hWnd As Long, _
     lpPoint As POINTAPI) As Long

Public Declare Function ScreenToClient Lib "user32" _
    (ByVal hWnd As Long, _
     lpPoint As POINTAPI) As Long
     
Public Declare Function GetWindowDC Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function GetDC Lib "user32" _
    (ByVal hWnd As Long) As Long
     
Public Declare Function ReleaseDC Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal hDC As Long) As Long
    
Public Declare Sub GetClientRect Lib "user32" _
    (ByVal hWnd As Long, _
     lpRect As RECT)

Public Declare Function FillRect Lib "user32" _
    (ByVal hDC As Long, _
     lpRect As RECT, _
     ByVal hBrush As Long) As Long

Public Declare Function WindowFromDC Lib "user32" _
    (ByVal hDC As Long) As Long

Public Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Public Declare Function CreateIconIndirect Lib "user32" _
    (piconinfo As ICONINFO) As Long

Public Declare Function CreateIconFromResource Lib "user32" _
    (ByVal presbits As Long, _
     ByVal dwResSize As Long, _
     ByVal fIcon As Long, _
     ByVal dwVer As Long) As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hWnd As Long, _
     ByVal crKey As Long, _
     ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) As Long

Public Const WS_EX_LAYERED = &H80000
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1

Public Declare Function ShowCursor Lib "user32" _
    (ByVal bShow As Long) As Long

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, _
     ByVal lpfn As Long, _
     ByVal hmod As Long, _
     ByVal dwThreadId As Long) As Long
     
Public Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long
  
Public Declare Function CallNextHookEx Lib "user32" _
    (ByVal hHook As Long, _
     ByVal nCode As Long, _
     ByVal wParam As Long, _
     lParam As Any) As Long

Public Const WH_KEYBOARD_LL = 13

Public Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer

Public Const LLKHF_ALTDOWN = &H20
Public Const VK_LWIN = &H5B
Public Const VK_RWIN = &H5C
Public Const HC_ACTION = 0
Public Const HC_GETNEXT = 1
Public Const HC_SKIP = 2

Public Declare Function GetKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer

Public Const VK_CONTROL = &H11

Public Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type


Public Declare Function UpdateWindow Lib "user32.dll" _
    (ByVal hWnd As Long) As Long

Public Declare Function LoadIconW Lib "user32.dll" _
    (ByVal hInstance As Long, _
     ByVal lpIconName As Long) As Long

Public Declare Function LoadImageW Lib "user32.dll" _
    (ByVal hinst As Long, _
     ByVal lpszName As Long, _
     ByVal uType As Long, _
     ByVal cxDesired As Long, _
     ByVal cyDesired As Long, _
     ByVal fuLoad As Long) As Long
     
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2

Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_VGA_COLOR = &H80
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H18000 - &H10000

Public Declare Function DrawTextW Lib "user32" _
    (ByVal hDC As Long, _
     ByVal lpString As Long, _
     ByVal nCount As Long, _
     lpRect As RECT, _
     ByVal uFormat As Long) As Long

Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000
Public Const DT_NOFULLWIDTHCHARBREAK = &H80000
Public Const DT_HIDEPREFIX = &H100000
Public Const DT_PREFIXONLY = &H200000


Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long
    
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYVTHUMB = 9
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXCURSOR = 13
Public Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_MOUSEPRESENT = 19
Public Const SM_CYVSCROLL = 20
Public Const SM_CXHSCROLL = 21
Public Const SM_DEBUG = 22
Public Const SM_SWAPBUTTON = 23
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_CXMIN = 28
Public Const SM_CYMIN = 29
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXICONSPACING = 38
Public Const SM_CYICONSPACING = 39
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_PENWINDOWS = 41
Public Const SM_DBCSENABLED = 42
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_SECURE = 44
Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46
Public Const SM_CXMINSPACING = 47
Public Const SM_CYMINSPACING = 48
Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50
Public Const SM_CYSMCAPTION = 51
Public Const SM_CXSMSIZE = 52
Public Const SM_CYSMSIZE = 53
Public Const SM_CXMENUSIZE = 54
Public Const SM_CYMENUSIZE = 55
Public Const SM_ARRANGE = 56
Public Const SM_CXMINIMIZED = 57
Public Const SM_CYMINIMIZED = 58
Public Const SM_CXMAXTRACK = 59
Public Const SM_CYMAXTRACK = 60
Public Const SM_CXMAXIMIZED = 61
Public Const SM_CYMAXIMIZED = 62
Public Const SM_NETWORK = 63
Public Const SM_CLEANBOOT = 67
Public Const SM_CXDRAG = 68
Public Const SM_CYDRAG = 69
Public Const SM_SHOWSOUNDS = 70
Public Const SM_CXMENUCHECK = 71
Public Const SM_CYMENUCHECK = 72
Public Const SM_SLOWMACHINE = 73
Public Const SM_MIDEASTENABLED = 74
Public Const SM_MOUSEWHEELPRESENT = 75
Public Const SM_CMETRICS = 76

Public Declare Function LoadKeyboardLayoutW Lib "user32" _
    (ByVal pwszKLID As Long, _
     ByVal flags As Long) As Long

Public Const KLF_ACTIVATE = &H1
Public Const KLF_NOTELLSHELL = &H80
Public Const KLF_REORDER = &H8
Public Const KLF_REPLACELANG = &H10
Public Const KLF_SUBSTITUTE_OK = &H2
Public Const KLF_SETFORPROCESS = &H100

Public Declare Function UnloadKeyboardLayout Lib "user32" _
    (ByVal hkl As Long) As Long

Public Declare Function GetKeyboardLayoutList Lib "user32" _
    (ByVal nBuff As Long, _
     lpList As Long) As Long

Public Declare Function ExitWindowsEx Lib "user32" _
    (ByVal uFlags As Long, _
     ByVal dwReason As Long) As Long

Public Const EWX_HYBRID_SHUTDOWN = &H400000
Public Const EWX_LOGOFF = &H0
Public Const EWX_POWEROFF = &H8
Public Const EWX_REBOOT = &H2
Public Const EWX_RESTARTAPPS = &H40
Public Const EWX_SHUTDOWN = &H1
Public Const EWX_FORCE = &H4
Public Const EWX_FORCEIFHUNG = &H10



