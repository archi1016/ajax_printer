Attribute VB_Name = "api_gdi"
Option Explicit

Public Declare Function GetCurrentObject Lib "gdi32" _
    (ByVal fnPenStyle As Long, _
     ByVal uObjectType As Long) As Long
     
Public Const OBJ_BRUSH = 2
Public Const OBJ_PEN = 1
Public Const OBJ_PAL = 5
Public Const OBJ_FONT = 6
Public Const OBJ_BITMAP = 7
     
Public Declare Function CreateSolidBrush Lib "gdi32" _
    (ByVal crColor As Long) As Long

Public Declare Function CreatePen Lib "gdi32" _
    (ByVal fnPenStyle As Long, _
     ByVal nWidth As Long, _
     ByVal crColor As Long) As Long

Public Const PS_SOLID = 0
Public Const PS_DASH = 1
Public Const PS_DOT = 2
Public Const PS_DASHDOT = 3
Public Const PS_DASHDOTDOT = 4
Public Const PS_NULL = 5
Public Const PS_INSIDEFRAME = 6

Public Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal hgdiobj As Long) As Long
    
Public Declare Function CreateCompatibleDC Lib "gdi32" _
    (ByVal hdc As Long) As Long
    
Public Declare Function DeleteDC Lib "gdi32" _
    (ByVal hdc As Long) As Long
    
Public Declare Function GetTextColor Lib "gdi32" _
    (ByVal hdc As Long) As Long
     
Public Declare Function SetTextColor Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal crColor As Long) As Long
    
Public Declare Function GetBkColor Lib "gdi32" _
    (ByVal hdc As Long) As Long
     
Public Declare Function SetBkColor Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal crColor As Long) As Long
     
Public Declare Function GetPixel Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal nX As Long, _
     ByVal nY As Long) As Long
     
Public Declare Function SetPixel Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal nX As Long, _
     ByVal nY As Long, _
     ByVal crColor As Long) As Long
     
Public Declare Function MoveToEx Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal nX As Long, _
     ByVal nY As Long, _
     ByVal lpPoint As Long) As Long
     
Public Declare Function LineTo Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal nXEnd As Long, _
     ByVal nYEnd As Long) As Long
     
Public Declare Function Polygon Lib "gdi32" _
    (ByVal hdc As Long, _
     lpPoints As POINTAPI, _
     ByVal nCount As Long) As Long
     
Public Declare Function Ellipse Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal nLeftRect As Long, _
     ByVal nTopRect As Long, _
     ByVal nRightRect As Long, _
     ByVal nBottomRect As Long) As Long
     
Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long) As Long

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type CIEXYZ
    ciexyzX As Long
    ciexyzY As Long
    ciexyzZ As Long
End Type

Public Type ICEXYZTRIPLE
    ciexyzRed As CIEXYZ
    ciexyzGreen As CIEXYZ
    ciexyzBlue As CIEXYZ
End Type

Public Type BITMAPV5HEADER
    bV5Size As Long
    bV5Width As Long
    bV5Height As Long
    bV5Planes As Integer
    bV5BitCount As Integer
    bV5Compression As Long
    bV5SizeImage As Long
    bV5XPelsPerMeter As Long
    bV5YPelsPerMeter As Long
    bV5ClrUsed As Long
    bV5ClrImportant As Long
    bV5RedMask As Long
    bV5GreenMask As Long
    bV5BlueMask As Long
    bV5AlphaMask As Long
    bV5CSType As Long
    bV5Endpoints As ICEXYZTRIPLE
    bV5GammaRed As Long
    bV5GammaGreen As Long
    bV5GammaBlue As Long
    bV5Intent As Long
    bV5ProfileData As Long
    bV5ProfileSize As Long
    bV5Reserved As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
End Type

Public Const BI_RGB = 0

Public Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
     pbmi As BITMAPINFO, _
     ByVal iUsage As Long, _
     ppvBits As Long, _
     ByVal hSection As Long, _
     ByVal dwOffset As Long) As Long

Public Declare Function CreateDIBBitmap Lib "gdi32" _
    (ByVal hdc As Long, _
     lpbmih As BITMAPINFOHEADER, _
     ByVal fdwInit As Long, _
     lpbInit As Long, _
     pbmi As BITMAPINFO, _
     ByVal fuUsage As Long) As Long
     
Public Const DIB_PAL_COLORS = 1
Public Const DIB_RGB_COLORS = 0

Public Declare Function BitBlt Lib "gdi32" _
    (ByVal hdcDest As Long, _
     ByVal nXDest As Long, _
     ByVal nYDest As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hdcSrc As Long, _
     ByVal nXSrc As Long, _
     ByVal nYSrc As Long, _
     ByVal dwRop As Long) As Long

Public Const SRCCOPY As Long = &HCC0020
Public Const CAPTUREBLT As Long = &H40000000

Public Declare Function PlgBlt Lib "gdi32" _
    (ByVal hdcDest As Long, _
     lpPoint As POINTAPI, _
     ByVal hdcSrc As Long, _
     ByVal nXSrc As Long, _
     ByVal nYSrc As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hbmMask As Long, _
     ByVal xMask As Long, _
     ByVal yMask As Long) As Long

Public Declare Function AlphaBlend Lib "msimg32" _
    (ByVal hdcDest As Long, _
     ByVal xoriginDest As Long, _
     ByVal yoriginDest As Long, _
     ByVal wDest As Long, _
     ByVal hDest As Long, _
     ByVal hdcSrc As Long, _
     ByVal xoriginSrc As Long, _
     ByVal yoriginSrc As Long, _
     ByVal wSrc As Long, _
     ByVal hSrc As Long, _
     ByVal ftn As Long) As Long

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Const AC_SRC_OVER = 0
Public Const AC_SRC_ALPHA = 1

Public Declare Function SetMapMode Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal fnMapMode As Long) As Long

Public Declare Function GetMapMode Lib "gdi32" _
    (ByVal hdc As Long) As Long
     
Public Const MM_ANISOTROPIC = 8
Public Const MM_HIENGLISH = 5
Public Const MM_HIMETRIC = 3
Public Const MM_ISOTROPIC = 7
Public Const MM_LOENGLISH = 4
Public Const MM_LOMETRIC = 2
Public Const MM_TEXT = 1
Public Const MM_TWIPS = 6

Public Declare Function SetGraphicsMode Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal iMode As Long) As Long
     
Public Declare Function GetGraphicsMode Lib "gdi32" _
    (ByVal hdc As Long) As Long

Public Const GM_ADVANCED = 2

Public Declare Function SetWorldTransform Lib "gdi32" _
    (ByVal hdc As Long, _
     lpXform As XFORM) As Long
 
Public Declare Function GetWorldTransform Lib "gdi32" _
    (ByVal hdc As Long, _
     lpXform As XFORM) As Long
     
Public Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

  
Public Declare Function EnumFontFamiliesExW Lib "gdi32" _
    (ByVal hdc As Long, _
     lpLogFont As LOGFONT, _
     ByVal lpEnumFontFamExProc As Long, _
     ByVal lParam As Long, _
     ByVal dwFlags As Long) As Long

Public Declare Function CreateFontIndirectW Lib "gdi32" _
    (lpLogFont As LOGFONT) As Long

Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE * 2 - 1) As Byte
End Type

Public Type ENUMLOGFONTEX
    elfLogFont As LOGFONT
    elfFullName(LF_FULLFACESIZE * 2 - 1) As Byte
    elfStyle(LF_FACESIZE * 2 - 1) As Byte
    elfScript(LF_FACESIZE * 2 - 1) As Byte
End Type

Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const OEM_CHARSET = 255

Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_ULTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_REGULAR = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_DEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_ULTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_BLACK = 900

Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_PRECIS = 4
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_OUTLINE_PRECIS = 8

Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_MASK = 15
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_TT_ALWAYS = 32
Public Const CLIP_EMBEDDED = 128

Public Const DEFAULT_QUALITY = 0
Public Const DRAFT_QUALITY = 1
Public Const PROOF_QUALITY = 2
Public Const NONANTIALIASED_QUALITY = 3
Public Const ANTIALIASED_QUALITY = 4

Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2

Public Const FF_DECORATIVE = 80
Public Const FF_DONTCARE = 0
Public Const FF_MODERN = 48
Public Const FF_ROMAN = 16
Public Const FF_SCRIPT = 64
Public Const FF_SWISS = 32

Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

     
Public Declare Function GetDeviceGammaRamp Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal lpRamp As Long) As Long

Public Declare Function SetDeviceGammaRamp Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal lpRamp As Long) As Long

Public Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hdc As Long, _
     ByVal nIndex As Long) As Long

Public Const HORZSIZE = 4
Public Const VERTSIZE = 6
Public Const HORZRES = 8
Public Const VERTRES = 10
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const BITSPIXEL = 12
Public Const PLANES = 14
Public Const NUMBRUSHES = 16
Public Const NUMPENS = 18
Public Const NUMFONTS = 22
Public Const NUMCOLORS = 24
Public Const NUMMARKERS = 20
Public Const ASPECTX = 40
Public Const ASPECTY = 42
Public Const ASPECTXY = 44
Public Const PDEVICESIZE = 26
Public Const CLIPCAPS = 36
Public Const SIZEPALETTE = 104
Public Const NUMRESERVED = 106
Public Const COLORRES = 108
Public Const PHYSICALWIDTH = 110
Public Const PHYSICALHEIGHT = 111
Public Const PHYSICALOFFSETX = 112
Public Const PHYSICALOFFSETY = 113
Public Const SCALINGFACTORX = 114
Public Const SCALINGFACTORY = 115
Public Const VREFRESH = 116
Public Const DESKTOPHORZRES = 118
Public Const DESKTOPVERTRES = 117
Public Const BLTALIGNMENT = 119
Public Const RASTERCAPS = 38
Public Const RC_BANDING = 2
Public Const RC_BITBLT = 1
Public Const RC_BITMAP64 = 8
Public Const RC_DI_BITMAP = 128
Public Const RC_DIBTODEV = 512
Public Const RC_FLOODFILL = 4096
Public Const RC_GDI20_OUTPUT = 16
Public Const RC_PALETTE = 256
Public Const RC_SCALING = 4
Public Const RC_STRETCHBLT = 2048
Public Const RC_STRETCHDIB = 8192
Public Const RC_DEVBITS As Long = &H8000
Public Const RC_OP_DX_OUTPUT = &H4000
Public Const CURVECAPS = 28
Public Const CC_NONE = 0
Public Const CC_CIRCLES = 1
Public Const CC_PIE = 2
Public Const CC_CHORD = 4
Public Const CC_ELLIPSES = 8
Public Const CC_WIDE = 16
Public Const CC_STYLED = 32
Public Const CC_WIDESTYLED = 64
Public Const CC_INTERIORS = 128
Public Const CC_ROUNDRECT = 256
Public Const LINECAPS = 30
Public Const LC_NONE = 0
Public Const LC_POLYLINE = 2
Public Const LC_MARKER = 4
Public Const LC_POLYMARKER = 8
Public Const LC_WIDE = 16
Public Const LC_STYLED = 32
Public Const LC_WIDESTYLED = 64
Public Const LC_INTERIORS = 128
Public Const POLYGONALCAPS = 32
Public Const RC_BIGFONT = 1024
Public Const RC_GDI20_STATE = 32
Public Const RC_NONE = 0
Public Const RC_SAVEBITMAP = 64
Public Const PC_NONE = 0
Public Const PC_POLYGON = 1
Public Const PC_POLYPOLYGON = 256
Public Const PC_PATHS = 512
Public Const PC_RECTANGLE = 2
Public Const PC_WINDPOLYGON = 4
Public Const PC_SCANLINE = 8
Public Const PC_TRAPEZOID = 4
Public Const PC_WIDE = 16
Public Const PC_STYLED = 32
Public Const PC_WIDESTYLED = 64
Public Const PC_INTERIORS = 128
Public Const TEXTCAPS = 34
Public Const TC_OP_CHARACTER = 1
Public Const TC_OP_STROKE = 2
Public Const TC_CP_STROKE = 4
Public Const TC_CR_90 = 8
Public Const TC_CR_ANY = 16
Public Const TC_SF_X_YINDEP = 32
Public Const TC_SA_DOUBLE = 64
Public Const TC_SA_INTEGER = 128
Public Const TC_SA_CONTIN = 256
Public Const TC_EA_DOUBLE = 512
Public Const TC_IA_ABLE = 1024
Public Const TC_UA_ABLE = 2048
Public Const TC_SO_ABLE = 4096
Public Const TC_RA_ABLE = 8192
Public Const TC_VA_ABLE = 16384
Public Const TC_RESERVED = 32768
Public Const TC_SCROLLBLT = 65536
