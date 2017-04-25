VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim IsPreviewMode As Boolean
Dim ToDevice As Object

Dim PaddingTop As Single
Dim PaddingBottom As Single
Dim PaddingLeft As Single
Dim PaddingRight As Single

Dim PageWidth As Single
Dim PageHeight As Single

Dim FieldFontHeight As Single
Dim FieldFontMaxHeight As Single
Dim FieldFontMinHeight As Single

Dim stFontName(127) As String
Dim stFontNameIndex As Long

Dim stFontSize(127) As Single
Dim stFontSizeIndex As Long

Dim stLocationX(127) As Single
Dim stLocationY(127) As Single
Dim stLocationIndex As Long

Dim stColor(127) As Long
Dim stColorIndex As Long

Private Sub Class_Initialize()
    PaddingTop = 0
    PaddingBottom = 0
    PaddingLeft = 0
    PaddingRight = 0
    
    stFontNameIndex = 0
    stLocationIndex = 0
    stColorIndex = 0
End Sub

Private Sub Class_Terminate()
    Set ToDevice = Nothing
End Sub

Public Sub Init(toDev As Object)
    Set ToDevice = toDev
End Sub

Public Sub SavePng(ByVal FP As String)
    Dim BmpFile As String
    Dim hImage As Long
    
    BmpFile = App.Path + "\" + PREVIEW_BMP
    DeleteFile BmpFile
    DeleteFile FP
    
    SavePicture ToDevice.Image, BmpFile
    If GpStatus_Ok = GdipLoadImageFromFile(StrPtr(BmpFile), hImage) Then
        Call SaveImageToPngFile(hImage, FP)
        GdipDisposeImage hImage
    End If
    
End Sub

Public Sub PrintErrorMessage(ByVal M As String)
    With ToDevice
        .ForeColor = vbRed
        .FontName = "Verdana"
        .FontSize = 16
        .Width = CountTextWidth(M) * Screen.TwipsPerPixelX
        .Height = CountTextHeight(M) * Screen.TwipsPerPixelY
        .Cls
        .CurrentX = 0
        .CurrentY = 0
    End With

    ToDevice.Print M
    ToDevice.Refresh
End Sub

Public Function PrintFromXmlStr(ByVal IsPreview As Boolean, XmlRoot As Object) As String
    Dim XmlNode As Object
    Dim S As String
    Dim T As String

    IsPreviewMode = IsPreview
    
    PageWidth = ToDevice.Width
    PageHeight = ToDevice.Height
    
    For Each XmlNode In XmlRoot.selectNodes("l")
         PrintFromXmlStr = RunScript(XmlNode, XmlNode.Text)
         If "" <> PrintFromXmlStr Then Exit For
    Next
        
    Set XmlNode = Nothing
    
End Function

Private Function RunScript(XmlNode As Object, ByVal T As String) As String
    Dim C As String
    Dim XmlNode_2 As Object

    RunScript = ""
    C = SafeGetXmlNodeAttr(XmlNode, "c")
    Select Case C
        Case "new_page"
            If IsPreviewMode Then
                ToDevice.Cls
            Else
                ToDevice.NewPage
            End If
            
        Case "done"
            If IsPreviewMode Then
                ToDevice.Refresh
            Else
                ToDevice.EndDoc
            End If
        
           
        Case "padding_top"
            PaddingTop = SafeGetDistance(T, PageWidth)
            PageHeight = PageHeight - PaddingTop - PaddingBottom
            ToDevice.CurrentY = PaddingTop
        
        Case "padding_bottom"
            PaddingBottom = SafeGetDistance(T, PageWidth)
            PageHeight = PageHeight - PaddingTop - PaddingBottom
        
        Case "padding_left"
            PaddingLeft = SafeGetDistance(T, PageWidth)
            PageWidth = PageWidth - PaddingLeft - PaddingRight
        
        Case "padding_right"
            PaddingRight = SafeGetDistance(T, PageWidth)
            PageWidth = PageWidth - PaddingLeft - PaddingRight
        
        Case "push_xy"
            stLocationX(stLocationIndex) = ToDevice.CurrentX
            stLocationY(stLocationIndex) = ToDevice.CurrentY
            stLocationIndex = stLocationIndex + 1

        Case "pop_xy"
            If 0 < stLocationIndex Then
                stLocationIndex = stLocationIndex - 1
                ToDevice.CurrentX = stLocationX(stLocationIndex)
                ToDevice.CurrentY = stLocationY(stLocationIndex)
            End If
            
        Case "absolute_x"
            ToDevice.CurrentX = PaddingLeft + SafeGetDistance(T, PageWidth)
            
        Case "absolute_y"
            ToDevice.CurrentY = PaddingTop + SafeGetDistance(T, PageHeight)
            
        Case "relative_x"
            ToDevice.CurrentX = ToDevice.CurrentX + SafeGetDistance(T, PageWidth)
            
        Case "relative_y"
            ToDevice.CurrentY = ToDevice.CurrentY + SafeGetDistance(T, PageHeight)
            
        Case "scroll_up"
            ToDevice.CurrentY = ToDevice.CurrentY - SafeGetDistance(T, PageHeight)
            
        Case "scroll_down"
            ToDevice.CurrentY = ToDevice.CurrentY + SafeGetDistance(T, PageHeight)
            
        Case "line"
            Call DrawLine(T)
        
        Case "print"
            ToDevice.CurrentX = PaddingLeft
            Call PrintText(XmlNode, T, False)
            
        Case "print_i"
            ToDevice.CurrentX = PaddingLeft
            Call PrintText(XmlNode, T, True)
            
        Case "print_center"
            ToDevice.CurrentX = PaddingLeft + (PageWidth - GetTextWidth(XmlNode, T)) / 2
            Call PrintText(XmlNode, T, False)
            
        Case "print_center_i"
            ToDevice.CurrentX = PaddingLeft + (PageWidth - GetTextWidth(XmlNode, T)) / 2
            Call PrintText(XmlNode, T, True)
            
        Case "print_right"
            ToDevice.CurrentX = PaddingLeft + (PageWidth - GetTextWidth(XmlNode, T))
            Call PrintText(XmlNode, T, False)
    
        Case "print_right_i"
            ToDevice.CurrentX = PaddingLeft + (PageWidth - GetTextWidth(XmlNode, T))
            Call PrintText(XmlNode, T, True)
            
            
        Case "barcode"
            ToDevice.CurrentX = PaddingLeft
            Call PrintBarcode(XmlNode, T)
            
        Case "barcode_center"
            ToDevice.CurrentX = PaddingLeft + (PageWidth - GetBarcodeWidth(XmlNode, T)) / 2
            Call PrintBarcode(XmlNode, T)
            
        Case "barcode_right"
            ToDevice.CurrentX = PaddingLeft + (PageWidth - GetBarcodeWidth(XmlNode, T))
            Call PrintBarcode(XmlNode, T)
            
            
        Case "fields"
            RunScript = PrintFields(XmlNode.selectNodes("f"))
        
        Case Else
            RunScript = RunCommonCommand(C, T)
            
    End Select
    
    Set XmlNode_2 = Nothing
End Function

Private Function RunCommonCommand(ByVal C As String, ByVal T As String) As String
    RunCommonCommand = ""
    Select Case C
        Case "set_font"
            ToDevice.FontName = T
        
         Case "push_font"
            stFontName(stFontNameIndex) = ToDevice.FontName
            stFontNameIndex = stFontNameIndex + 1
            
        Case "pop_font"
            If 0 < stFontNameIndex Then
                stFontNameIndex = stFontNameIndex - 1
                ToDevice.FontName = stFontName(stFontNameIndex)
            End If
            
            
        Case "set_size"
            ToDevice.FontSize = CSng(T)
            
        Case "push_size"
            stFontSize(stFontSizeIndex) = ToDevice.FontSize
            stFontSizeIndex = stFontSizeIndex + 1
            
        Case "pop_size"
            If 0 < stFontSizeIndex Then
                stFontSizeIndex = stFontSizeIndex - 1
                ToDevice.FontSize = stFontSize(stFontSizeIndex)
            End If
            
            
        Case "set_color"
            ToDevice.ForeColor = SafeGetColor(T)
        
        Case "push_color"
            stColor(stColorIndex) = ToDevice.ForeColor
            stColorIndex = stColorIndex + 1
            
        Case "pop_color"
            If 0 < stColorIndex Then
                stColorIndex = stColorIndex - 1
                ToDevice.ForeColor = stColor(stColorIndex)
            End If
           
        Case Else
            RunCommonCommand = "unknow command \""" + C + "\"""
            
    End Select
End Function

Private Function GetBarcodeWidth(XmlNode As Object, ByVal T As String) As Single
    Dim BR As New cBarcode
    Dim barW As Single
    
    barW = CSng(SafeGetXmlNodeAttr(XmlNode, "bw")) * 20!
    GetBarcodeWidth = BR.GetWidth(T, barW)
    
    Set BR = Nothing
End Function

Private Sub PrintBarcode(XmlNode As Object, ByVal T As String)
    Dim BR As New cBarcode
    Dim barW As Single
    Dim barH As Single
    
    barW = CSng(SafeGetXmlNodeAttr(XmlNode, "bw")) * 20!
    barH = CSng(SafeGetXmlNodeAttr(XmlNode, "h")) * 20!
    If 0 < barH Then
        BR.Draw ToDevice, T, barW, barH
    End If
    
    Set BR = Nothing
End Sub

Private Function PrintFields(XmlNodeList As Object) As String
    Dim XmlNode As Object
    Dim C As String
    Dim T As String
    Dim LastY As Single
    Dim fL As Single
    Dim fW As Single

    PrintFields = ""
    LastY = ToDevice.CurrentY
    FieldFontMaxHeight = 0
    FieldFontMinHeight = 99999
    fL = PaddingLeft
    For Each XmlNode In XmlNodeList
        C = SafeGetXmlNodeAttr(XmlNode, "c")
        T = XmlNode.Text
    
        If "print" = Left$(C, 5) Then
            fW = SafeGetDistance(SafeGetXmlNodeAttr(XmlNode, "w"), PageWidth)

            Select Case C
                Case "print"
                    ToDevice.CurrentX = fL
                    ToDevice.CurrentY = LastY
                    Call PrintText(XmlNode, T, False)
                
                Case "print_i"
                    ToDevice.CurrentX = fL
                    ToDevice.CurrentY = LastY
                    Call PrintText(XmlNode, T, True)
                    
                Case "print_center"
                    ToDevice.CurrentX = fL + (fW - GetTextWidth(XmlNode, T)) / 2
                    ToDevice.CurrentY = LastY
                    Call PrintText(XmlNode, T, False)
                    
                Case "print_center_i"
                    ToDevice.CurrentX = fL + (fW - GetTextWidth(XmlNode, T)) / 2
                    ToDevice.CurrentY = LastY
                    Call PrintText(XmlNode, T, True)
                    
                Case "print_right"
                    ToDevice.CurrentX = fL + (fW - GetTextWidth(XmlNode, T))
                    ToDevice.CurrentY = LastY
                    Call PrintText(XmlNode, T, False)

                Case "print_right_i"
                    ToDevice.CurrentX = fL + (fW - GetTextWidth(XmlNode, T))
                    ToDevice.CurrentY = LastY
                    Call PrintText(XmlNode, T, True)
                    
            End Select
            fL = fL + fW
            If FieldFontHeight > FieldFontMaxHeight Then FieldFontMaxHeight = FieldFontHeight
            If FieldFontHeight < FieldFontMinHeight Then FieldFontMinHeight = FieldFontHeight
        Else
            PrintFields = RunCommonCommand(C, T)
        End If
    Next
    
    ToDevice.CurrentY = LastY + FieldFontMinHeight
    
    Set XmlNode = Nothing
End Function

Private Function GetTextWidth(XmlNode As Object, ByVal T As String) As Single
    Dim newFontName As String
    Dim newFontSize As String
    Dim prevFontName As String
    Dim prevFontSize As Single
    
    newFontName = SafeGetXmlNodeAttr(XmlNode, "font")
    newFontSize = SafeGetXmlNodeAttr(XmlNode, "size")
    If "" <> newFontName Then
        prevFontName = ToDevice.FontName
        ToDevice.FontName = newFontName
    End If
    If "" <> newFontSize Then
        prevFontSize = ToDevice.FontSize
        ToDevice.FontSize = CSng(newFontSize)
    End If
    
    GetTextWidth = ToDevice.TextWidth(T)
    
    If "" <> newFontName Then
        ToDevice.FontName = prevFontName
    End If
    If "" <> newFontSize Then
        ToDevice.FontSize = prevFontSize
    End If
End Function

Private Sub PrintText(XmlNode As Object, ByVal T As String, ByVal IsInvert As Boolean)
    Dim L() As String
    Dim U As Long
    Dim I As Long
    Dim newFontName As String
    Dim newFontSize As String
    Dim prevFontName As String
    Dim prevFontSize As Single
    Dim prevColor As Long
    Dim prevX As Single
    Dim prevY As Single


    newFontName = SafeGetXmlNodeAttr(XmlNode, "font")
    newFontSize = SafeGetXmlNodeAttr(XmlNode, "size")
    If "" <> newFontName Then
        prevFontName = ToDevice.FontName
        ToDevice.FontName = newFontName
    End If
    If "" <> newFontSize Then
        prevFontSize = ToDevice.FontSize
        ToDevice.FontSize = CSng(newFontSize)
    End If
    
    If "" = T Then T = " "
    FieldFontHeight = ToDevice.TextHeight(T)
'Debug.Print ToDevice.FontName + ": " + T + ": " + CStr(FieldFontHeight)
    
    L = Split(T, "\n")
    U = UBound(L)

    prevX = ToDevice.CurrentX
    If IsInvert Then
        prevColor = ToDevice.ForeColor
        
        prevY = ToDevice.CurrentY
        ToDevice.Line (prevX, prevY)-(prevX + ToDevice.TextWidth(T), prevY + ToDevice.TextHeight(T)), , BF
        
        ToDevice.ForeColor = vbWhite
        ToDevice.CurrentY = prevY
        
        For I = 0 To U
            ToDevice.CurrentX = prevX
            ToDevice.Print L(I)
        Next

        ToDevice.ForeColor = prevColor
    Else
        For I = 0 To U
            ToDevice.CurrentX = prevX
            ToDevice.Print L(I)
        Next
    End If
    
    If "" <> newFontName Then
        ToDevice.FontName = prevFontName
    End If
    If "" <> newFontSize Then
        ToDevice.FontSize = prevFontSize
    End If
    
    Erase L
End Sub

Private Sub DrawLine(ByVal T As String)
    Dim H As Single
    
    H = SafeGetDistance(T, PageWidth)
    ToDevice.Line (PaddingLeft, ToDevice.CurrentY)-(PaddingLeft + PageWidth, ToDevice.CurrentY + H), , BF
End Sub

Private Function CountTextWidth(ByVal T As String) As Single
    Dim RC As RECT
    
    DrawTextW ToDevice.hDC, StrPtr(T), -1, RC, DT_CALCRECT
    CountTextWidth = CSng(RC.Right - RC.Left)
End Function

Private Function CountTextHeight(ByVal T As String) As Single
    Dim RC As RECT
    
    CountTextHeight = CSng(DrawTextW(ToDevice.hDC, StrPtr(T), -1, RC, DT_CALCRECT))
End Function


Private Function SafeGetColor(ByVal T As String) As Long
    Dim S As String
    Dim B(3) As Byte
    
    SafeGetColor = 0
    If "#" = Left$(T, 1) Then
        S = T + "000000"

        B(0) = Val("&H" + Mid$(S, 2, 2))
        B(1) = Val("&H" + Mid$(S, 4, 2))
        B(2) = Val("&H" + Mid$(S, 6, 2))
        CopyMemory VarPtr(SafeGetColor), VarPtr(B(0)), 3
    End If
End Function

Private Function SafeGetDistance(ByVal T As String, ByVal D As Single) As Single
    SafeGetDistance = 0
    If "" <> T Then
        If 0 = InStr(T, "%") Then
            SafeGetDistance = CSng(T) * 20!
        Else
            SafeGetDistance = CSng(Replace(T, "%", ""))
            SafeGetDistance = SafeGetDistance / 100
            SafeGetDistance = D * SafeGetDistance
        End If
    End If
End Function


