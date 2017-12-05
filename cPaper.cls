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
Dim DS As cDrawString

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

Private Sub Class_Initialize()
    Set DS = New cDrawString

    PaddingTop = 0
    PaddingBottom = 0
    PaddingLeft = 0
    PaddingRight = 0
    
    stFontNameIndex = 0
    stLocationIndex = 0
End Sub

Private Sub Class_Terminate()
    Set DS = Nothing
    Set ToDevice = Nothing
End Sub

Public Sub Init(toDev As Object, devTwipsPerPixelX As Single, devTwipsPerPixelY As Single)
    Set ToDevice = toDev
    
    ToDevice.PSet (0, 0), vbWhite   'start printer hdc
    
    DS.Init devTwipsPerPixelX, devTwipsPerPixelY
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
        .Width = .TextWidth(M)
        .Height = .TextHeight(M)
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
         PrintFromXmlStr = RunScript(XmlNode)
         If "" <> PrintFromXmlStr Then Exit For
    Next
        
    Set XmlNode = Nothing
    
End Function

Private Function RunScript(XmlNode As Object) As String
    Dim T As String
    Dim C As String
    Dim XmlNode_2 As Object

    RunScript = ""
    T = XmlNode.Text
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
            PageHeight = ToDevice.Height - PaddingTop - PaddingBottom
            ToDevice.CurrentY = PaddingTop
        
        Case "padding_bottom"
            PaddingBottom = SafeGetDistance(T, PageWidth)
            PageHeight = ToDevice.Height - PaddingTop - PaddingBottom
        
        Case "padding_left"
            PaddingLeft = SafeGetDistance(T, PageWidth)
            PageWidth = ToDevice.Width - PaddingLeft - PaddingRight
        
        Case "padding_right"
            PaddingRight = SafeGetDistance(T, PageWidth)
            PageWidth = ToDevice.Width - PaddingLeft - PaddingRight
        
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
            
        Case "line_w"
            Call DrawLineWidth(XmlNode)
        
        
        
        Case "print"
            Call PrintText(XmlNode, PaddingLeft, PageWidth, DT_LEFT, False)
            
        Case "print_i"
            Call PrintText(XmlNode, PaddingLeft, PageWidth, DT_LEFT, True)
            
        Case "print_center"
            Call PrintText(XmlNode, PaddingLeft, PageWidth, DT_CENTER, False)
            
        Case "print_center_i"
            Call PrintText(XmlNode, PaddingLeft, PageWidth, DT_CENTER, True)
            
        Case "print_right"
            Call PrintText(XmlNode, PaddingLeft, PageWidth, DT_RIGHT, False)
    
        Case "print_right_i"
            Call PrintText(XmlNode, PaddingLeft, PageWidth, DT_RIGHT, True)
            
            
            
        Case "print_xy"
            Call PrintTextXY(XmlNode)
            
        Case "print_vt"
            Call PrintTextVertical(XmlNode)
            
 
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
            
            
        'Case "set_color"
        '    ToDevice.ForeColor = SafeGetColor(T)
        
        'Case "push_color"
        '    stColor(stColorIndex) = ToDevice.ForeColor
        '    stColorIndex = stColorIndex + 1
            
        'Case "pop_color"
        '    If 0 < stColorIndex Then
        '        stColorIndex = stColorIndex - 1
        '        ToDevice.ForeColor = stColor(stColorIndex)
        '    End If
           
        Case Else
            RunCommonCommand = "unknow command """ + C + """"
            
    End Select
End Function

Private Function GetBarcodeWidth(XmlNode As Object, ByVal T As String) As Single
    Dim BR As New cBarcode
    Dim barW As Single
    
    barW = CSng(SafeGetXmlNodeAttr(XmlNode, "bw"))
    GetBarcodeWidth = BR.GetWidth(T, barW)
    
    Set BR = Nothing
End Function

Private Sub PrintBarcode(XmlNode As Object, ByVal T As String)
    Dim BR As New cBarcode
    Dim barW As Single
    Dim barH As Single
    
    barW = CSng(SafeGetXmlNodeAttr(XmlNode, "bw"))
    barH = CSng(SafeGetXmlNodeAttr(XmlNode, "h"))
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
            ToDevice.CurrentY = LastY

            Select Case C
                Case "print"
                    Call PrintText(XmlNode, fL, fW, DT_LEFT, False)
                
                Case "print_i"
                    Call PrintText(XmlNode, fL, fW, DT_LEFT, True)
                    
                Case "print_center"
                    Call PrintText(XmlNode, fL, fW, DT_CENTER, False)
                    
                Case "print_center_i"
                    Call PrintText(XmlNode, fL, fW, DT_CENTER, True)
                    
                Case "print_right"
                    Call PrintText(XmlNode, fL, fW, DT_RIGHT, False)

                Case "print_right_i"
                    Call PrintText(XmlNode, fL, fW, DT_RIGHT, True)
                    
            End Select
            
            fL = fL + fW
            If FieldFontHeight > FieldFontMaxHeight Then FieldFontMaxHeight = FieldFontHeight
            If FieldFontHeight < FieldFontMinHeight Then FieldFontMinHeight = FieldFontHeight
        Else
            PrintFields = RunCommonCommand(C, T)
        End If
    Next
    
    ToDevice.CurrentY = LastY + FieldFontMaxHeight
    
    Set XmlNode = Nothing
End Function

Private Sub PrintTextVertical(XmlNode As Object)
    Dim T As String
    Dim TxtWidth As Single
    Dim TxtHeight As Single
    Dim prevX As Single
    Dim prevY As Single
    Dim hdc As Long
    Dim locX As Single
    Dim locY As Single
    
    Dim RC As RECT

    prevX = ToDevice.CurrentX
    prevY = ToDevice.CurrentY

    hdc = DS.CreateVerticalFont(XmlNode, ToDevice)
    
    T = XmlNode.Text
    If "" = T Then T = " "
    DS.CountSize hdc, T, PageWidth, TxtWidth, TxtHeight
    
    locX = PaddingLeft + SafeGetDistance(SafeGetXmlNodeAttr(XmlNode, "x"), PageWidth)
    locY = PaddingTop + SafeGetDistance(SafeGetXmlNodeAttr(XmlNode, "y"), PageHeight)
    DS.DrawVertical ToDevice.hdc, T, locX, locY, TxtWidth, TxtHeight
    Call DrawTextBorder(XmlNode, locX, locY, TxtWidth, TxtHeight)

    ToDevice.CurrentX = prevX
    ToDevice.CurrentY = prevY
    
    DS.DeleteVerticalFont ToDevice
End Sub

Private Sub PrintTextXY(XmlNode As Object)
    Dim T As String
    Dim TxtWidth As Single
    Dim TxtHeight As Single
    Dim prevX As Single
    Dim prevY As Single
    Dim hdc As Long
    Dim locX As Single
    Dim locY As Single
    
    Dim RC As RECT

    prevX = ToDevice.CurrentX
    prevY = ToDevice.CurrentY

    DS.ChangeFont XmlNode, ToDevice
    
    T = XmlNode.Text
    If "" = T Then T = " "
    DS.CountSize ToDevice.hdc, T, PageWidth, TxtWidth, TxtHeight
    
    locX = PaddingLeft + SafeGetDistance(SafeGetXmlNodeAttr(XmlNode, "x"), PageWidth)
    locY = PaddingTop + SafeGetDistance(SafeGetXmlNodeAttr(XmlNode, "y"), PageHeight)
    DS.Draw ToDevice.hdc, T, DT_LEFT, locX, locY, TxtWidth, TxtHeight
    Call DrawTextBorder(XmlNode, locX, locY, TxtWidth, TxtHeight)

    ToDevice.CurrentX = prevX
    ToDevice.CurrentY = prevY
    
    DS.RestoreFont ToDevice
End Sub

Private Sub PrintText(XmlNode As Object, ByVal drawLeft As Single, ByVal drawWidth As Single, ByVal nFormat As Long, ByVal IsInvert As Boolean)
    Dim T As String
    Dim TxtWidth As Single
    Dim TxtHeight As Single
    Dim prevX As Single
    Dim prevY As Single
    Dim RC As RECT

    ToDevice.CurrentX = drawLeft
    prevX = ToDevice.CurrentX
    prevY = ToDevice.CurrentY

    DS.ChangeFont XmlNode, ToDevice
    
    T = XmlNode.Text
    If "" = T Then T = " "
    DS.CountSize ToDevice.hdc, T, drawWidth, TxtWidth, TxtHeight
    FieldFontHeight = TxtHeight     'For Field Print
    
    Select Case nFormat
         Case DT_CENTER
             prevX = prevX + (drawWidth - TxtWidth) / 2

         Case DT_RIGHT
             prevX = prevX + (drawWidth - TxtWidth)

    End Select
    
    Call DrawTextBorder(XmlNode, prevX, prevY, TxtWidth, TxtHeight)
    
    If IsInvert Then
        ToDevice.Line (prevX, prevY)-(prevX + TxtWidth, prevY + TxtHeight), , BF
        ToDevice.ForeColor = vbWhite
    End If

    DS.Draw ToDevice.hdc, T, nFormat, prevX, prevY, TxtWidth, TxtHeight

    If IsInvert Then
        ToDevice.ForeColor = vbBlack
    End If
    
    ToDevice.CurrentX = drawLeft
    ToDevice.CurrentY = prevY + TxtHeight
    
    DS.RestoreFont ToDevice
End Sub

Private Sub DrawTextBorder(XmlNode As Object, ByVal dX As Single, ByVal dy As Single, ByVal TxtWidth As Single, ByVal TxtHeight As Single)
    Dim BorderT As String
    Dim BorderWidth As Single
    Dim I As Single
    
    BorderT = SafeGetXmlNodeAttr(XmlNode, "border")
    If "" <> BorderT Then
        BorderWidth = CSng(BorderT) * 20!
        If 0 < BorderWidth Then
            For I = 1 To BorderWidth
                ToDevice.Line (dX - I, dy - I)-(dX + TxtWidth + I, dy + TxtHeight + I), , B
            Next
        End If
    End If
End Sub

Private Sub DrawLine(ByVal T As String)
    Dim prevX As Single
    Dim prevY As Single
    Dim H As Single
    
    prevX = ToDevice.CurrentX
    prevY = ToDevice.CurrentY

    H = SafeGetDistance(T, PageWidth)
    ToDevice.Line (PaddingLeft, ToDevice.CurrentY)-(PaddingLeft + PageWidth, ToDevice.CurrentY + H), , BF

    ToDevice.CurrentX = prevX
    ToDevice.CurrentY = prevY + H
End Sub

Private Sub DrawLineWidth(XmlNode As Object)
    Dim prevX As Single
    Dim prevY As Single
    Dim lB As Single
    Dim lE As Single
    Dim H As Single
    
    prevX = ToDevice.CurrentX
    prevY = ToDevice.CurrentY

    lB = SafeGetDistance(SafeGetXmlNodeAttr(XmlNode, "begin"), PageWidth)
    lE = SafeGetDistance(SafeGetXmlNodeAttr(XmlNode, "end"), PageWidth)

    H = SafeGetDistance(XmlNode.Text, PageWidth)
    ToDevice.Line (PaddingLeft + lB, ToDevice.CurrentY)-(PaddingLeft + lE, ToDevice.CurrentY + H), , BF

    ToDevice.CurrentX = prevX
    ToDevice.CurrentY = prevY + H
End Sub

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


