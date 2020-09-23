Attribute VB_Name = "declare_mod"
Option Explicit

Public Enum vbColorsExtra
    vbLightGrey = &HC0C0C0
    vbGrey = &H808080
    vbLightYellow = &HC0FFFF
End Enum

Public Type POINT
    X       As Long
    Y       As Long
End Type

Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Const TA_CENTER = 6
Private Const TA_NOUPDATECP = 0
Private Const TA_TOP = 0
Public Const TA_MAP = (TA_CENTER + TA_TOP + TA_NOUPDATECP)

Public Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Public Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

