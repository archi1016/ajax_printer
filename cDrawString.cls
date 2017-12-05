VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DRAW_TEXT_FORMAT = DT_WORDBREAK Or DT_HIDEPREFIX Or DT_NOPREFIX

Dim TwipsPerPixelX As Single
Dim TwipsPerPixelY As Single

Dim NewFontName As String
Dim NewFontSize As String
Dim PrevFontName As String
Dim PrevFontSize As Single
Dim TextTopAlign As String
Dim TextAngle As Long
Dim hDCcompatible As Long
Dim hFontVertical As Long
Dim hFontOld As Long

Private Sub Class_Initialize()
    '
End Sub

Private Sub Class_Terminate()
    '
End Sub

Public Sub Init(TppX As Single, TppY As Single)
    TwipsPerPixelX = TppX
    TwipsPerPixelY = TppY
End Sub

Public Function t2pX(ByVal V As Single) As Long
    t2pX = CLng(V / TwipsPerPixelX)
End Function

Public Function t2pY(ByVal V As Single) As Long
    t2pY = CLng(V / TwipsPerPixelY)
End Function

Public Function p2tX(ByVal V As Long) As Single
    p2tX = V * TwipsPerPixelX
End Function

Public Function p2tY(ByVal V As Long) As Single
    p2tY = V * TwipsPerPixelY
End Function

Public Sub ChangeFont(XmlNode As Object, desObj As Object)
    NewFontName = SafeGetXmlNodeAttr(XmlNode, "font")
    NewFontSize = SafeGetXmlNodeAttr(XmlNode, "size")
    
    If "" <> NewFontName Then
        PrevFontName = desObj.FontName
        desObj.FontName = NewFontName
    End If
    If "" <> NewFontSize Then
        PrevFontSize = desObj.FontSize
        desObj.FontSize = CSng(NewFontSize)
    End If
End Sub

Public Sub RestoreFont(desObj As Object)
    If "" <> NewFontName Then
        desObj.FontName = PrevFontName
    End If
    If "" <> NewFontSize Then
        desObj.FontSize = PrevFontSize
    End If
End Sub

Public Function CreateVerticalFont(XmlNode As Object, desObj As Object) As Long
    Dim LF As LOGFONT
    
    If 0 <> hFontVertical Then
        DeleteObject hFontVertical
        hFontVertical = 0
    End If
    If 0 <> hDCcompatible Then
        DeleteDC hDCcompatible
        hDCcompatible = 0
    End If
    
    PrevFontName = desObj.FontName
    PrevFontSize = desObj.FontSize
    NewFontName = SafeGetXmlNodeAttr(XmlNode, "font")
    NewFontSize = SafeGetXmlNodeAttr(XmlNode, "size")
    TextTopAlign = UCase$(SafeGetXmlNodeAttr(XmlNode, "align"))
    
    If "" <> NewFontName Then PrevFontName = NewFontName
    If "" <> NewFontSize Then PrevFontSize = CSng(NewFontSize)
    
    Select Case TextTopAlign
        Case "LEFT"
            TextAngle = 90
        
        Case "RIGHT"
            TextAngle = 270
        
        Case "BOTTOM"
            TextAngle = 180
        
    End Select

    hDCcompatible = CreateCompatibleDC(desObj.hdc)
    If 0 <> hDCcompatible Then
        With LF
            .lfHeight = CLng((PrevFontSize * CSng(GetDeviceCaps(hDCcompatible, LOGPIXELSY))) / 72)
            .lfWidth = 0
            .lfEscapement = TextAngle * 10
            .lfWeight = FW_NORMAL
            .lfCharSet = DEFAULT_CHARSET
            .lfOutPrecision = OUT_DEFAULT_PRECIS
            .lfClipPrecision = CLIP_DEFAULT_PRECIS
            .lfQuality = DEFAULT_QUALITY
            .lfPitchAndFamily = DEFAULT_PITCH
            CopyMemory VarPtr(.lfFaceName(0)), StrPtr(PrevFontName), Len(PrevFontName) * 2
        End With
        
        hFontVertical = CreateFontIndirectW(LF)
        If 0 <> hFontVertical Then
            hFontOld = SelectObject(hDCcompatible, hFontVertical)
        End If
    End If
    
    CreateVerticalFont = hDCcompatible
End Function

Public Sub DeleteVerticalFont(desObj As Object)
    If 0 <> hDCcompatible Then
        If 0 <> hFontVertical Then
            SelectObject hDCcompatible, hFontOld
            DeleteObject hFontVertical
            hFontVertical = 0
        End If
        
        DeleteDC hDCcompatible
        hDCcompatible = 0
    End If
End Sub

Public Sub CountSize(ByVal hdc As Long, ByVal T As String, ByVal drawWidth As Single, retWidth As Single, retHeight As Single)
    Dim RC As RECT
    
    RC.Right = t2pX(drawWidth)
    retHeight = p2tY(DrawTextW(hdc, StrPtr(T), -1, RC, DT_CALCRECT Or DRAW_TEXT_FORMAT))
    retWidth = p2tX(RC.Right - RC.Left)
End Sub

Public Sub Draw(ByVal hdc As Long, ByVal T As String, ByVal nFormat As Long, ByVal drawX As Single, ByVal drawY As Single, ByVal drawWidth As Single, ByVal drawHeight As Single)
    Dim RC As RECT
    
    With RC
        .Left = t2pX(drawX)
        .Top = t2pY(drawY)
        .Right = .Left + t2pX(drawWidth)
        .Bottom = .Top + t2pY(drawHeight)
    End With
    DrawTextW hdc, StrPtr(T), -1, RC, nFormat Or DT_NOCLIP Or DRAW_TEXT_FORMAT
End Sub

Public Function DrawVertical(ByVal hdc As Long, ByVal T As String, ByVal locX As Single, ByVal locY As Single, drawWidth As Single, drawHeight As Single)
    Dim drawX As Single
    Dim drawY As Single
    Dim nW As Long
    Dim nH As Long
    Dim hBitmap As Long
    Dim hBitmapOld As Long
    Dim RC As RECT
    
    If 0 <> hDCcompatible Then
        drawX = 0
        drawY = 0
        Select Case TextAngle
            Case 90
                Call SwitchSingleValue(drawWidth, drawHeight)
                drawY = drawX + drawHeight
                
            Case 180
                drawX = drawX + drawWidth
                drawY = drawX + drawHeight
            
            Case 270
                Call SwitchSingleValue(drawWidth, drawHeight)
                drawX = drawX + drawWidth
                
        End Select
        
        nW = t2pX(drawWidth)
        nH = t2pY(drawHeight)
        hBitmap = CreateCompatibleBitmap(hDCcompatible, nW, nH)
        If 0 <> hBitmap Then
            hBitmapOld = SelectObject(hDCcompatible, hBitmap)
            
            With RC
                .Left = t2pX(drawX)
                .Top = t2pY(drawY)
                .Right = .Left + nW
                .Bottom = .Top + nH
            End With
            DrawTextW hDCcompatible, StrPtr(T), -1, RC, DT_NOCLIP Or DRAW_TEXT_FORMAT
            BitBlt hdc, t2pX(locX), t2pY(locY), nW, nH, hDCcompatible, 0, 0, SRCCOPY
            
            SelectObject hDCcompatible, hBitmapOld
            DeleteObject hBitmap
        End If
    End If
End Function

Public Sub SwitchSingleValue(v1 As Single, v2 As Single)
    Dim V As Single
    
    V = v1
    v1 = v2
    v2 = V
End Sub
