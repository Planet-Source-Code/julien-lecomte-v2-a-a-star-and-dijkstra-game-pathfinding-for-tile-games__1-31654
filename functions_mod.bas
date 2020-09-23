Attribute VB_Name = "functions_mod"
Option Explicit

'*******************************************************************************************
'Functions
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************
'

Public Enum PEN_MODE
    PM_SETTILE = 0
    PM_SETSTART = 1
    PM_SETEND = 2
End Enum

Public Enum TILE_HARDNESS
    TH_EASY = 1
    TH_NORMAL = 3
    TH_HARD = 6
    TH_VERYHARD = 9
    TH_UNWALKABLE = 10
End Enum

Public Enum PATH_MAP
    PATH_IMPOSSIBLE = -2
    PATH_EMPTY = -1
    PATH_HUGE = 2147483647
End Enum
    
Public Const SLOW_DOWN_VALUE = 10 '//Milliseconds

Public NUMBER_OF_TILES&
Public TILE_SIDE&

Private Const RDW_INVALIDATE = &H1
Private Const RDW_ERASE = &H4
Private Const RDW_UPDATENOW = &H100
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Sub RefreshRect(hwnd&, rectClient As RECT)
    '// Created 18/07/01
    RedrawWindow hwnd, rectClient, ByVal 0&, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE
End Sub

Public Function Random&(ByVal lMin&, ByVal lMax&)
    '//Assumes a randomize has already been called somewhere in the program
    Random = Int((lMax - lMin + 1) * Rnd + lMin)
End Function

Public Sub SetPoint(ptPoint As POINT, X&, Y&)
    '//Fills a point structure
    ptPoint.X = X
    ptPoint.Y = Y
End Sub

Public Function ComparePoints(ptPoint1 As POINT, ptPoint2 As POINT) As Boolean
    '//Return true if points are equal
    If ptPoint1.X = ptPoint2.X And ptPoint1.Y = ptPoint2.Y Then ComparePoints = True
End Function

Public Function FormatTime$(lTimeMilliseconds&)
    Dim lMilliseconds&
    Dim lSeconds&
    Dim lMinutes&
    
    lSeconds = lTimeMilliseconds \ 1000
    lMilliseconds = lTimeMilliseconds - lSeconds * 1000
    lMinutes = lSeconds \ 60
    lSeconds = lSeconds - lMinutes * 60
    
    FormatTime = Right$("00" & lMinutes, 2) & "m " & Right$("00" & lSeconds, 2) & "s " & Right$("000" & lMilliseconds, 3)
End Function
