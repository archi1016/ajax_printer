VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const NUMBER_BARS = "nnWWn,WnnnW,nWnnW,WWnnn,nnWnW,WnWnn,nWWnn,nnnWW,WnnWn,nWnWn"

Dim NumberBars() As String

Private Sub Class_Initialize()
    NumberBars = Split(NUMBER_BARS, ",")
End Sub

Private Sub Class_Terminate()
    Erase NumberBars
End Sub

Public Sub Draw(ToDevice As Object, ByVal T As String, ByVal BarWidth As Single, ByVal BarHeight As Single)
    Dim C As Long
    Dim I As Long
    Dim J As Long
    Dim tmpBar(2) As String
    Dim Bar As String
    Dim locX As Single
    Dim locY As Single
    Dim offsetX As Single
    
    'T = FillChecksum(T)
    If 0 <> (Len(T) And 1) Then T = "0" + T
    C = Len(T) \ 2
    Bar = ""
    For I = 1 To C
        tmpBar(0) = NumberBars(CLng(Mid$(T, I * 2 - 1, 1)))
        tmpBar(1) = NumberBars(CLng(Mid$(T, I * 2, 1)))
        tmpBar(2) = ""
        
        For J = 1 To 5
            If "W" = Mid$(tmpBar(0), J, 1) Then
                tmpBar(2) = tmpBar(2) + "11"
            Else
                tmpBar(2) = tmpBar(2) + "1"
            End If
            If "W" = Mid$(tmpBar(1), J, 1) Then
                tmpBar(2) = tmpBar(2) + "00"
            Else
                tmpBar(2) = tmpBar(2) + "0"
            End If
        Next
        Bar = Bar + tmpBar(2)
    Next
    Bar = "0001010" + Bar + "1101000"

    locX = ToDevice.CurrentX
    locY = ToDevice.CurrentY
    C = Len(Bar)
    For I = 1 To C
        If "1" = Mid$(Bar, I, 1) Then
            offsetX = locX + I * BarWidth
            ToDevice.Line (offsetX, locY)-(offsetX + BarWidth, locY + BarHeight), , BF
        End If
    Next
End Sub

Public Function GetWidth(ByVal T As String, ByVal BarWidth As Single) As Single
    Dim L As Long
    
    L = (CountCodeLength(T) + 2) * 7
    GetWidth = BarWidth * CSng(L)
End Function

Private Function FillChecksum(ByVal T As String) As String
    Dim L As Long
    Dim I As Long
    Dim evenSum As Long
    Dim oddSum As Long
    Dim V As Long
    
    L = Len(T)
    evenSum = 0
    oddSum = 0
    For I = 1 To L
        If 0 <> (I And 1) Then
            evenSum = evenSum + CLng(Mid$(T, I, 1))
        Else
            oddSum = oddSum + CLng(Mid$(T, I, 1))
        End If
    Next
    evenSum = evenSum * 3

    V = (evenSum + oddSum) Mod 10
    If 0 < V Then V = 10 - V
    
    FillChecksum = T + CStr(V)
    If 0 <> (Len(FillChecksum) And 1) Then FillChecksum = "0" + FillChecksum
End Function

Private Function CountCodeLength(ByVal T As String) As Long
    Dim L As Long
    
    'L = Len(T) + 1
    L = Len(T)
    If 0 <> (L And 1) Then L = L + 1
    CountCodeLength = L
End Function


