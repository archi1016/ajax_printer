Attribute VB_Name = "api_gdiplus"
Option Explicit

Public Const ImageLockModeRead = 1
Public Const ImageLockModeWrite = 2
Public Const ImageLockModeUserInputBuf = 4

Public Const PixelFormatIndexed = &H10000
Public Const PixelFormatGDI = &H20000
Public Const PixelFormatAlpha = &H40000
Public Const PixelFormatPAlpha = &H80000
Public Const PixelFormatExtended = &H100000
Public Const PixelFormatCanonical = &H200000

Public Const PixelFormat32bppARGB = 10 Or 32 * 256 Or PixelFormatAlpha Or PixelFormatGDI Or PixelFormatCanonical

Public Type BitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    Scan0 As Long
    Reserved As Long
End Type



Public Const WrapModeTile = 0
Public Const WrapModeTileFlipX = 1
Public Const WrapModeTileFlipY = 2
Public Const WrapModeTileFlipXY = 3
Public Const WrapModeClamp = 4

Public Const MatrixOrderPrepend = 0
Public Const MatrixOrderAppend = 1

Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Declare Function GdiplusStartup Lib "gdiplus" _
    (hToken As Long, _
     InputBuf As GdiplusStartupInput, _
     ByVal OutputBuf As Long) As Long
     
Public Declare Function GdiplusShutdown Lib "gdiplus" _
    (ByVal hToken As Long) As Long

Public Const GpStatus_Ok = 0
Public Const GpStatus_GenericError = 1
Public Const GpStatus_InvalidParameter = 2
Public Const GpStatus_OutOfMemory = 3
Public Const GpStatus_ObjectBusy = 4
Public Const GpStatus_InsufficientBuffer = 5
Public Const GpStatus_NotImplemented = 6
Public Const GpStatus_Win32Error = 7
Public Const GpStatus_WrongState = 8
Public Const GpStatus_Aborted = 9
Public Const GpStatus_FileNotFound = 10
Public Const GpStatus_ValueOverflow = 11
Public Const GpStatus_AccessDenied = 12
Public Const GpStatus_UnknownImageFormat = 13
Public Const GpStatus_FontFamilyNotFound = 14
Public Const GpStatus_FontStyleNotFound = 15
Public Const GpStatus_NotTrueTypeFont = 16
Public Const GpStatus_UnsupportedGdiplusVersion = 17
Public Const GpStatus_GdiplusNotInitialized = 18
Public Const GpStatus_PropertyNotFound = 19
Public Const GpStatus_PropertyNotSupported = 20

Public Declare Function GdipCreateFromHDC Lib "gdiplus" _
    (ByVal hdc As Long, _
     hGraphics As Long) As Long

Public Declare Function GdipDeleteGraphics Lib "gdiplus" _
    (ByVal hGraphics As Long) As Long
    
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" _
    (ByVal unicodeFilename As Long, _
     hImage As Long) As Long

Public Declare Function GdipDisposeImage Lib "gdiplus" _
    (ByVal hImage As Long) As Long
    
Public Declare Function GdipCloneImage Lib "gdiplus" _
    (ByVal hImage As Long, _
     hCloneImage As Long) As Long
     
Public Declare Function GdipGetImageWidth Lib "gdiplus" _
    (ByVal hImage As Long, _
     iWidth As Long) As Long
     
Public Declare Function GdipGetImageHeight Lib "gdiplus" _
    (ByVal hImage As Long, _
     iHeight As Long) As Long

Public Declare Function GdipGetImagePixelFormat Lib "gdiplus" _
    (ByVal hImage As Long, _
     iFormat As Long) As Long
     
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" _
    (ByVal hImage As Long, _
     hGraphics As Long) As Long

Public Declare Function GdipCreateTexture Lib "gdiplus" _
    (ByVal hImage As Long, _
     ByVal wrapMode As Long, _
     hTextureBrush As Long) As Long

Public Declare Function GdipCreatePen1 Lib "gdiplus" _
    (ByVal argbColor As Long, _
     ByVal nWdith As Single, _
     ByVal srcUnit As Long, _
     hPen As Long) As Long
     
Public Declare Function GdipDeletePen Lib "gdiplus" _
    (ByVal hPen As Long) As Long
    
Public Declare Function GdipSetPenWidth Lib "gdiplus" _
    (ByVal hPen As Long, _
     ByVal nWidth As Single) As Long
     
Public Declare Function GdipSetPenLineJoin Lib "gdiplus" _
    (ByVal hPen As Long, _
     ByVal nLineJoin As Long) As Long
    
Public Const LineJoinMiter = 0
Public Const LineJoinBevel = 1
Public Const LineJoinRound = 2
Public Const LineJoinMiterClipped = 3
    
Public Declare Function GdipCreateSolidFill Lib "gdiplus" _
    (ByVal argbColor As Long, _
     hBrush As Long) As Long
     
Public Type POINTF
    pX As Single
    pY As Single
End Type
     
Public Declare Function GdipCreateLineBrush Lib "gdiplus" _
    (point1 As POINTF, _
     point2 As POINTF, _
     ByVal color1 As Long, _
     ByVal color2 As Long, _
     ByVal wrapMode As Long, _
     hLineGradientBrush As Long) As Long
     
Public Declare Function GdipDeleteBrush Lib "gdiplus" _
    (ByVal hBrush As Long) As Long

Public Declare Function GdipGetTextureImage Lib "gdiplus" _
    (ByVal hTextureBrush As Long, _
     hImage As Long) As Long
    
Public Declare Function GdipResetTextureTransform Lib "gdiplus" _
    (ByVal hTextureBrush As Long) As Long
    
Public Declare Function GdipTranslateTextureTransform Lib "gdiplus" _
    (ByVal hTextureBrush As Long, _
     ByVal dX As Single, _
     ByVal dy As Single, _
     ByVal order As Long) As Long
    
'Public Declare Function GdipSetTextureWrapMode Lib "GDIPlus" _
    (ByVal hTextureBrush As Long, _
     ByVal wrapmode As Long) As Long

Public Declare Function GdipFillRectangleI Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal hBrush As Long, _
     ByVal lX As Long, _
     ByVal lY As Long, _
     ByVal lWidth As Long, _
     ByVal lHeight As Long) As Long

Public Declare Function GdipDrawImageI Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal hImage As Long, _
     ByVal lX As Long, _
     ByVal lY As Long) As Long
     
Public Declare Function GdipDrawImageRectI Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal hImage As Long, _
     ByVal lX As Long, _
     ByVal lY As Long, _
     ByVal bWidth As Long, _
     ByVal bHeight As Long) As Long

Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal hImage As Long, _
     ByVal dstX As Long, _
     ByVal dstY As Long, _
     ByVal dstWidth As Long, _
     ByVal dstHeight As Long, _
     ByVal srcX As Long, _
     ByVal srcY As Long, _
     ByVal srcWidth As Long, _
     ByVal srcHeight As Long, _
     ByVal srcUnit As Long, _
     ByVal hImageAttributes As Long, _
     ByVal callback As Long, _
     ByVal callbackData As Long) As Long
     
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" _
    (ByVal unicodeFilename As Long, _
     hBitmap As Long) As Long
    
Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" _
    (ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal stride As Long, _
     ByVal PixelFormat As Long, _
     ByVal lpScan0 As Long, _
     hBitmap As Long) As Long
     
Public Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" _
    (ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hGraphics As Long, _
     hBitmap As Long) As Long
     
Public Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" _
    (ByVal hBitmap As Long, _
     hIcon As Long) As Long

Public Declare Function GdipBitmapLockBits Lib "gdiplus" _
    (ByVal hBitmap As Long, _
     RC As RECT, _
     ByVal flags As Long, _
     ByVal PixelFormat As Long, _
     LockedBitmapData As BitmapData) As Long

Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" _
    (ByVal hBitmap As Long, _
     LockedBitmapData As BitmapData) As Long

Public Declare Function GdipCreateImageAttributes Lib "gdiplus" _
    (hImageAttributes As Long) As Long

Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" _
    (ByVal hImageAttributes As Long) As Long

Public Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" _
    (ByVal hImageAttributes As Long, _
     ByVal wrap As Long, _
     ByVal argb As Long, _
     ByVal clamp As Long) As Long
     
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal pixelOffsetMode As Long) As Long

Public Const HighQuality = 2

Public Declare Function GdipSetInterpolationMode Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal interpolationMode As Long) As Long

Public Declare Function GdipSetCompositingMode Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal compositingMode As Long) As Long
     
Public Declare Function GdipSetCompositingQuality Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal compositingQuality As Long) As Long

Public Declare Function GdipSetSmoothingMode Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal nSmoothingMode As Long) As Long

Public Declare Function GdipSetPageUnit Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal nUnit As Long) As Long

Public Const QualityModeInvalid = -1
Public Const QualityModeDefault = 0
Public Const QualityModeLow = 1
Public Const QualityModeHigh = 2

Public Const PixelOffsetModeInvalid = QualityModeInvalid
Public Const PixelOffsetModeDefault = QualityModeDefault
Public Const PixelOffsetModeHighSpeed = QualityModeLow
Public Const PixelOffsetModeHighQuality = QualityModeHigh
Public Const PixelOffsetModeNone = QualityModeHigh + 1
Public Const PixelOffsetModeHalf = QualityModeHigh + 2

Public Const InterpolationModeInvalid = QualityModeInvalid
Public Const InterpolationModeDefault = QualityModeDefault
Public Const InterpolationModeLowQuality = QualityModeLow
Public Const InterpolationModeHighQuality = QualityModeHigh
Public Const InterpolationModeBilinear = QualityModeHigh + 1
Public Const InterpolationModeBicubic = QualityModeHigh + 2
Public Const InterpolationModeNearestNeighbor = QualityModeHigh + 3
Public Const InterpolationModeHighQualityBilinear = QualityModeHigh + 4
Public Const InterpolationModeHighQualityBicubic = QualityModeHigh + 5

Public Const SmoothingModeInvalid = QualityModeInvalid
Public Const SmoothingModeDefault = QualityModeDefault
Public Const SmoothingModeHighSpeed = QualityModeLow
Public Const SmoothingModeHighQuality = QualityModeHigh
Public Const SmoothingModeNone = 3
Public Const SmoothingModeAntiAlias = 4

Public Const CompositingModeSourceOver = 0
Public Const CompositingModeSourceCopy = 1


Public Const CompositingQualityDefault = QualityModeDefault
Public Const CompositingQualityHighSpeed = QualityModeLow
Public Const CompositingQualityHighQuality = QualityModeHigh
Public Const CompositingQualityGammaCorrected = 3
Public Const CompositingQualityAssumeLinear = 4

Public Const UnitWorld = 0
Public Const UnitDisplay = 1
Public Const UnitPixel = 2
Public Const UnitPoint = 3
Public Const UnitInch = 4
Public Const UnitDocument = 5
Public Const UnitMillimeter = 6

Public Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type




Public Declare Function GdipFillPath Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal hBrush As Long, _
     ByVal hPath As Long) As Long

Public Declare Function GdipDrawPath Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal hPen As Long, _
     ByVal hPath As Long) As Long


Public Const LANG_NEUTRAL = 0

Public Declare Function GdipCreateStringFormat Lib "gdiplus" _
    (ByVal formatAttributes As Long, _
     ByVal language As Long, _
     hFormat As Long) As Long
     
Public Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" _
    (hFormat As Long) As Long
    
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" _
    (ByVal hFormat As Long) As Long
    
     
Public Const FillModeAlternate = 0
Public Const FillModeWinding = 1

Public Declare Function GdipCreatePath Lib "gdiplus" _
    (ByVal brushMode As Long, _
     hPath As Long) As Long

Public Declare Function GdipDeletePath Lib "gdiplus" _
    (ByVal hPath As Long) As Long
     
Public Declare Function GdipAddPathString Lib "gdiplus" _
    (ByVal hPath As Long, _
     ByVal WCHAR As Long, _
     ByVal length As Long, _
     ByVal hFontFamily As Long, _
     ByVal style As Long, _
     ByVal emSize As Single, _
     layoutRECT As RECTF, _
     ByVal hStringFormat As Long) As Long
     
Public Declare Function GdipGetPathWorldBounds Lib "gdiplus" _
 (ByVal hPath As Long, _
  bounds As RECTF, _
  ByVal hMatrix As Long, _
  ByVal hPen As Long) As Long



Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" _
    (ByVal WCHAR As Long, _
     ByVal fontCollection As Long, _
     hFontFamily As Long) As Long
    
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" _
    (ByVal hFontFamily As Long) As Long

Public Declare Function GdipCreateFontFromDC Lib "gdiplus" _
    (ByVal hdc As Long, _
     hFont As Long) As Long

Public Declare Function GdipCreateFont Lib "gdiplus" _
    (ByVal hFontFamily As Long, _
     ByVal emSize As Single, _
     ByVal style As Long, _
     ByVal unit As Long, _
     hFont As Long) As Long
     
Public Const FontStyleRegular = &H0
Public Const FontStyleBold = &H1
Public Const FontStyleItalic = &H2
Public Const FontStyleUnderline = &H4
Public Const FontStyleStrikeout = &H8
     
Public Declare Function GdipDeleteFont Lib "gdiplus" _
    (ByVal hFont As Long) As Long
     
Public Declare Function GdipMeasureString Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal WCHAR As Long, _
     ByVal length As Long, _
     ByVal hFont As Long, _
     layoutRECT As RECTF, _
     ByVal hStringFormat As Long, _
     boundingBox As RECTF, _
     codepointsFitted As Long, _
     linesFilled As Long) As Long

Public Declare Function GdipDrawString Lib "gdiplus" _
    (ByVal hGraphics As Long, _
     ByVal WCHAR As Long, _
     ByVal length As Long, _
     ByVal hFont As Long, _
     layoutRECT As RECTF, _
     ByVal hStringFormat As Long, _
     ByVal hBrush As Long) As Long

Public Declare Function GdipSaveImageToFile Lib "gdiplus" _
    (ByVal hImage As Long, _
     ByVal unicodeFilename As Long, _
     clsidEncoder As GUID, _
     ByVal encoderParams As Long) As Long
     
Public Type EncoderParameter
    GID   As GUID
    NumberOfValues   As Long
    Type   As Long
    Value   As Long
End Type
    
Public Type EncoderParameters
    Count   As Long
    Parameter   As EncoderParameter
End Type

Public Const EncoderParameterValueTypeByte = 1
Public Const EncoderParameterValueTypeASCII = 2
Public Const EncoderParameterValueTypeShort = 3
Public Const EncoderParameterValueTypeLong = 4
Public Const EncoderParameterValueTypeRational = 5
Public Const EncoderParameterValueTypeLongRange = 6
Public Const EncoderParameterValueTypeUndefined = 7
Public Const EncoderParameterValueTypeRationalRange = 8


