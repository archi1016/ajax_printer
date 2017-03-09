Attribute VB_Name = "define_nm"
Option Explicit

Public Const NM_FIRST = 0
Public Const NM_OUTOFMEMORY = NM_FIRST - 1
Public Const NM_CLICK = NM_FIRST - 2
Public Const NM_DBLCLK = NM_FIRST - 3
Public Const NM_RETURN = NM_FIRST - 4
Public Const NM_RCLICK = NM_FIRST - 5
Public Const NM_RDBLCLK = NM_FIRST - 6
Public Const NM_SETFOCUS = NM_FIRST - 7
Public Const NM_KILLFOCUS = NM_FIRST - 8
Public Const NM_CUSTOMDRAW = NM_FIRST - 12
Public Const NM_HOVER = NM_FIRST - 13
Public Const NM_NCHITTEST = NM_FIRST - 14
Public Const NM_KEYDOWN = NM_FIRST - 15
Public Const NM_RELEASEDCAPTURE = NM_FIRST - 16
Public Const NM_SETCURSOR = NM_FIRST - 17
Public Const NM_CHAR = NM_FIRST - 18
Public Const NM_LDOWN = NM_FIRST - 20

Public Type NMHDR
   hwndFrom As Long
   idfrom As Long
   Code As Long
End Type

Public Type NMKEY
    hdr As NMHDR
    nVKey As Long
    uFlags As Long
End Type

Public Type NMCUSTOMDRAW
    hdr As NMHDR
    dwDrawStage As Long
    hDC As Long
    RC As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
End Type

Public Const CDDS_PREPAINT = &H1
Public Const CDDS_POSTPAINT = &H2
Public Const CDDS_PREERASE = &H3
Public Const CDDS_POSTERASE = &H4
Public Const CDDS_ITEM = &H10000
Public Const CDDS_ITEMPREPAINT = CDDS_ITEM Or CDDS_PREPAINT
Public Const CDDS_ITEMPOSTPAINT = CDDS_ITEM Or CDDS_POSTPAINT
Public Const CDDS_ITEMPREERASE = CDDS_ITEM Or CDDS_PREERASE
Public Const CDDS_ITEMPOSTERASE = CDDS_ITEM Or CDDS_POSTERASE
Public Const CDDS_SUBITEM = &H20000
Public Const CDDS_SUBITEMPREPAINT = CDDS_SUBITEM Or CDDS_ITEMPREPAINT
Public Const CDDS_SUBITEMPOSTPAINT = CDDS_SUBITEM Or CDDS_ITEMPOSTPAINT
Public Const CDDS_SUBITEMPREERASE = CDDS_SUBITEM Or CDDS_ITEMPREERASE
Public Const CDDS_SUBITEMPOSTERASE = CDDS_SUBITEM Or CDDS_ITEMPOSTERASE

Public Const CDIS_SELECTED = &H1
Public Const CDIS_GRAYED = &H2
Public Const CDIS_DISABLED = &H4
Public Const CDIS_CHECKED = &H8
Public Const CDIS_FOCUS = &H10
Public Const CDIS_DEFAULT = &H20
Public Const CDIS_HOT = &H40
Public Const CDIS_MARKED = &H80
Public Const CDIS_INDETERMINATE = &H100

Public Const CDRF_DODEFAULT = &H0
Public Const CDRF_NEWFONT = &H2
Public Const CDRF_SKIPDEFAULT = &H4
Public Const CDRF_NOTIFYPOSTPAINT = &H10
Public Const CDRF_NOTIFYITEMDRAW = &H20
Public Const CDRF_NOTIFYSUBITEMDRAW = &H20
Public Const CDRF_NOTIFYPOSTERASE = &H40


