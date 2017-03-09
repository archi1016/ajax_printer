Attribute VB_Name = "define_panel"
Option Explicit

Public Const DEFAULT_CONTENT_TOP = 80
Public Const DEFAULT_CONTROL_BORDER_WIDTH = 2
Public Const DEFAULT_FOCUS_BORDER_WIDTH = 5
Public Const DEFAULT_BUTTON_WIDTH = 144
Public Const DEFAULT_BUTTON_HEIGHT = 40
Public Const DEFAULT_INPUT_MARGIN = 10
Public Const DEFAULT_ROW_HEIGHT = 48
Public Const DEFAULT_ICON_SIZE = 32
Public Const DEFAULT_ICON_MARGIN = 8

Public Const DEFAULT_PANEL_HEIGHT = 400
Public Const DEFAULT_PANEL_CONTROL_TOP = 36
Public Const DEFAULT_PANEL_DESCRIPTION_TOP = 32
Public Const DEFAULT_PANEL_DESCRIPTION_MARGIN = 4
Public Const DEFAULT_PANEL_MARGIN = 20
Public Const DEFAULT_PANEL_TIMER_INTERVAL = 5
Public Const DEFAULT_PANEL_MOVE_STEPS = 2
Public Const DEFAULT_PANEL_MOVE_TIMES = 3

Public Const TITLE_PANEL_HEIGHT = DEFAULT_ICON_SIZE + DEFAULT_ICON_MARGIN * 2

Public Const NAVIGATION_PANEL_WIDTH = 240
Public Const NAVIGATION_PANEL_BACK_WIDTH = 60
Public Const NAVIGATION_PANEL_BACK_HEIGHT = 60
Public Const NAVIGATION_PANEL_CAPTION_LEFT = 60
Public Const NAVIGATION_PANEL_CAPTION_TOP = 20
Public Const NAVIGATION_PANEL_ROW_HEIGHT = 48
Public Const NAVIGATION_PANEL_TIMER_INTERVAL = 5
Public Const NAVIGATION_PANEL_MOVE_STEPS = 60
Public Const NAVIGATION_PANEL_MOVE_TIMES = NAVIGATION_PANEL_WIDTH \ NAVIGATION_PANEL_MOVE_STEPS

Public Const BOTTOM_PANEL_MARGIN = 10
Public Const SECTION_PANEL_WIDTH = 184
Public Const DIALOG_PANEL_WIDTH = 880
Public Const DIALOG_PANEL_MIN_WIDTH = 640
Public Const DIALOG_PANEL_HEIGHT = 620
Public Const DIALOG_PANEL_MIN_HEIGHT = 480
Public Const STATUS_PANEL_HEIGHT = 28

Public Const PROCESSING_CONTENT_TOP = 100
Public Const PROCESSING_INTERVAL = 50
Public Const PROCESSING_BAR_WIDTH = 80
Public Const PROCESSING_BAR_HEIGHT = 28
Public Const PROCESSING_BAR_STEP = 2

Public Sub PanelCountDialogWidthAndLeft(ByVal fromWidth As Long, toWidth As Long, toLeft As Long)
    toWidth = fromWidth
    If fromWidth > DIALOG_PANEL_WIDTH Then
        toWidth = DIALOG_PANEL_WIDTH
    Else
        If fromWidth < DIALOG_PANEL_MIN_WIDTH Then
            toWidth = DIALOG_PANEL_MIN_WIDTH
        End If
    End If
    toLeft = (fromWidth - toWidth) \ 2
    If toLeft <= 0 Then toLeft = 0
End Sub

Public Sub PanelCountDialogHeightAndTop(ByVal fromHeight As Long, toHeight As Long, toTop As Long)
    toHeight = fromHeight
    If fromHeight > DIALOG_PANEL_HEIGHT Then
        toHeight = DIALOG_PANEL_HEIGHT
    Else
        If fromHeight < DIALOG_PANEL_MIN_HEIGHT Then
            toHeight = DIALOG_PANEL_MIN_HEIGHT
        End If
    End If
    toTop = (fromHeight - toHeight) \ 2
    If toTop <= 0 Then toTop = 0
End Sub
