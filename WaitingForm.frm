VERSION 5.00
Begin VB.Form WaitingForm 
   BorderStyle     =   1  '單線固定
   Caption         =   "Waiting"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "微軟正黑體"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "WaitingForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   220
   ScaleMode       =   3  '像素
   ScaleWidth      =   680
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton BtnClose 
      Height          =   675
      Left            =   6660
      TabIndex        =   1
      Top             =   2280
      Width           =   2475
   End
   Begin VB.Timer TimerForCountdown 
      Interval        =   1000
      Left            =   720
      Top             =   1860
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Lab"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000AA&
      Height          =   600
      Left            =   540
      TabIndex        =   0
      Top             =   480
      Width           =   810
   End
End
Attribute VB_Name = "WaitingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CountdownValue As Long

Private Sub BtnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = App.ProductName
    
    Lab.Caption = "等候網路連線穩定"
    
    CountdownValue = 8
    Call SetButtonCaption
End Sub

Private Sub Form_Resize()
    Lab.Move DEFAULT_PANEL_MARGIN, DEFAULT_PANEL_MARGIN
    BtnClose.Move Me.ScaleWidth - DEFAULT_BUTTON_WIDTH - DEFAULT_PANEL_MARGIN, Me.ScaleHeight - DEFAULT_BUTTON_HEIGHT - DEFAULT_PANEL_MARGIN, DEFAULT_BUTTON_WIDTH, DEFAULT_BUTTON_HEIGHT
End Sub

Private Sub SetButtonCaption()
    BtnClose.Caption = "進入 (" + CStr(CountdownValue) + ")"
    DoEvents
End Sub

Private Sub TimerForCountdown_Timer()
    CountdownValue = CountdownValue - 1
    Call SetButtonCaption
    If 0 = CountdownValue Then Unload Me
End Sub

