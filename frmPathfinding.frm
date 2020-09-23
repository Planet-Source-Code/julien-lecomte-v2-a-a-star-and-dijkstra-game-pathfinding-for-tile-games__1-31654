VERSION 5.00
Begin VB.Form frmPathfinding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pathfinding"
   ClientHeight    =   7410
   ClientLeft      =   150
   ClientTop       =   690
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   494
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optRunMode 
      Caption         =   "A*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   7425
      TabIndex        =   19
      Top             =   2700
      Width           =   1815
   End
   Begin VB.OptionButton optRunMode 
      Caption         =   "Dijkstra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   7425
      TabIndex        =   18
      Top             =   2400
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox txtNodes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7425
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4650
      Width           =   1815
   End
   Begin VB.CheckBox chkRunSlow 
      Caption         =   "Slow things down"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   13
      Top             =   3555
      Width           =   1815
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   120
      ScaleHeight     =   478
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   478
      TabIndex        =   0
      Top             =   120
      Width           =   7200
   End
   Begin VB.CheckBox chkShowPathTaken 
      Caption         =   "Show path values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   12
      Top             =   3315
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.TextBox txtTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4125
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox txtSave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      MaxLength       =   20
      TabIndex        =   6
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ComboBox cmbOpenFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7425
      TabIndex        =   4
      Top             =   1950
      Width           =   1815
   End
   Begin VB.Frame fraTiles 
      Caption         =   "Tiles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   7440
      TabIndex        =   1
      Top             =   225
      Width           =   1815
      Begin VB.OptionButton optTile 
         Caption         =   "Set start tile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   225
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optTile 
         Caption         =   "Set end tile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   14
         Top             =   525
         Width           =   1455
      End
      Begin VB.ComboBox cmbTypeOfTile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPathfinding.frx":0000
         Left            =   150
         List            =   "frmPathfinding.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1125
         Width           =   1575
      End
      Begin VB.OptionButton optTile 
         Caption         =   "Set tile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   825
         Width           =   1455
      End
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Node loop count"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7425
      TabIndex        =   17
      Top             =   4425
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   496
      X2              =   616
      Y1              =   336
      Y2              =   336
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   9
      X1              =   496
      X2              =   616
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   8
      X1              =   496
      X2              =   616
      Y1              =   9
      Y2              =   9
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   7
      X1              =   496
      X2              =   616
      Y1              =   124
      Y2              =   124
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   6
      X1              =   496
      X2              =   616
      Y1              =   125
      Y2              =   125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   3
      X1              =   496
      X2              =   616
      Y1              =   337
      Y2              =   337
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   496
      X2              =   616
      Y1              =   457
      Y2              =   457
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   496
      X2              =   616
      Y1              =   456
      Y2              =   456
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Total time taken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   11
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Save file as"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   9
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Open file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   8
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileRandom 
         Caption         =   "Random map"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "&Clear map"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuFileRedraw 
         Caption         =   "&Redraw Map"
      End
      Begin VB.Menu mnuFileNone0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileTiny 
         Caption         =   "&Tiny mode"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuFileNone 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunMode 
         Caption         =   "Run Djikstra"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuRunMode 
         Caption         =   "Run A*"
         Index           =   1
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmPathfinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************
'Pathfinding
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************
'
Option Explicit

'// Used by class collection via array method
Private objTree   As clsTree

Private lNodeLooped   As Long '// Track number of Recursive calls
Private ptStart       As POINT
Private ptEnd         As POINT
Private aiMap()       As Long '//(0 based) to store type of tiles.
Private bDirty        As Boolean '//Do we have to clear the map ?

    
Private Sub chkShowPathTaken_Click()
    chkRunSlow.Enabled = CBool(chkShowPathTaken.Value = vbChecked)
    If Not chkRunSlow.Enabled Then chkRunSlow.Value = vbUnchecked
End Sub

Private Sub cmbOpenFile_Click()
    '//Does what it says, anyway this app isn't about opening files!
    Dim sFilePath As String
    Dim lFile As Long
    
    If cmbOpenFile.ListIndex = -1 Then Exit Sub
    If Not (objTree Is Nothing) Then Exit Sub '// Is running
    
    sFilePath = App.Path
    If Right$(sFilePath, 1) <> "\" Then sFilePath = sFilePath & "\"
    sFilePath = sFilePath & cmbOpenFile.List(cmbOpenFile.ListIndex)
    
On Error GoTo ErrHappened
    lFile = FreeFile
    Open sFilePath For Binary Access Read As lFile
    Get lFile, , ptStart
    Get lFile, , ptEnd
    Get lFile, , aiMap()
    Close lFile
    
    Populate_GridMap
    
    Exit Sub
ErrHappened:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Err n°" & Err.Number
On Error Resume Next
    Close lFile '//Close if was opened
End Sub

Private Sub cmbTypeOfTile_Change()
    optTile(0).Value = True
End Sub

Private Sub cmbTypeOfTile_Click()
    optTile(0).Value = True
End Sub

Private Sub cmdRun_Click()
    Dim vObj As OptionButton
    Dim I%
    
    For Each vObj In optRunMode
        If vObj.Value Then mnuRunMode_Click I
        I = I + 1
    Next
    Set vObj = Nothing
End Sub

Private Sub cmdSave_Click()
    Dim sFilePath As String
    Dim lAnswer As Long
    Dim lFile As Long
    
    '//Need of a filename
    If Trim$(txtSave) = "" Then
        MsgBox "Filename is blank. Can't save"
        Exit Sub
    End If
    
    '//Get filename
    sFilePath = App.Path
    If Right$(sFilePath, 1) <> "\" Then sFilePath = sFilePath & "\"
    sFilePath = sFilePath & Trim$(txtSave) & IIf(TILE_SIDE = 20, ".a", ".at")
    
    '//Overwrite ?
    If Dir$(sFilePath) <> "" Then
        lAnswer = MsgBox("File already exists, overwrite with this one ?", vbYesNo Or vbQuestion, "Overwrite ?")
        If lAnswer = vbNo Then Exit Sub
    End If
    
On Error Resume Next
    Kill sFilePath '//Just for the fun of it :-), or in case, to overwrite nicely.
On Error GoTo ErrHappened
    '//Write to file
    lFile = FreeFile
    Open sFilePath For Binary Access Write As lFile
    Put lFile, , ptStart
    Put lFile, , ptEnd
    Put lFile, , aiMap()
    Close lFile
    
    Populate_FileCombo
    
    Exit Sub
ErrHappened:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Err n°" & Err.Number
On Error Resume Next
    Close lFile '//Close if was opened
End Sub

Private Sub Form_Load()
    'SetTextAlign picMap.hdc, TA_MAP '// Doesn't seem to work
        
    '//Populate the combo box
    With cmbTypeOfTile
    .AddItem "Easy tile (1)"
    .ItemData(.NewIndex) = TH_EASY
    .AddItem "Normal tile (3)"
    .ItemData(.NewIndex) = TH_NORMAL
    .AddItem "Hard tile (6)"
    .ItemData(.NewIndex) = TH_HARD
    .AddItem "Very Hard tile (9)"
    .ItemData(.NewIndex) = TH_VERYHARD
    .AddItem "Unwalkable"
    .ItemData(.NewIndex) = TH_UNWALKABLE
    .ListIndex = 0
    End With
    
    mnuFileTiny_Click
End Sub

Private Sub Map_Setup()
    ReDim aiMap(NUMBER_OF_TILES - 1, NUMBER_OF_TILES - 1) As Long
    '//Populate file ComboBox
    Populate_FileCombo
    
    '//Create a blank file in case loading of the file fails
    mnuFileClear_Click
    
    '//open first file
On Error Resume Next
    cmbOpenFile.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (objTree Is Nothing) Then
        Cancel = True
        Exit Sub
    End If
    Erase aiMap()
End Sub

Private Function GetColorFromHardness(ByVal lTileHardness As Long) As Long
    Select Case lTileHardness
        Case TH_EASY:       GetColorFromHardness = vbLightYellow
        Case TH_NORMAL, 0:  GetColorFromHardness = vbWhite
        Case TH_HARD:       GetColorFromHardness = vbLightGrey
        Case TH_VERYHARD:   GetColorFromHardness = vbGrey
        Case TH_UNWALKABLE: GetColorFromHardness = vbBlack
    End Select
End Function

Private Sub Tile_Paint(ByRef ptWhich As POINT, ByVal lColor As Long, Optional bRefresh As Boolean = True)
    Dim rectPaint As RECT
    Dim hBrush As Long
    
    With ptWhich
    SetRect rectPaint, .X * TILE_SIDE, .Y * TILE_SIDE, .X * TILE_SIDE + TILE_SIDE, .Y * TILE_SIDE + TILE_SIDE
    End With
    hBrush = CreateSolidBrush(lColor)
    FillRect picMap.hdc, rectPaint, hBrush
    If bRefresh Then RefreshRect picMap.hwnd, rectPaint
    DeleteObject hBrush
End Sub

Private Sub Tile_Print(ByRef ptWhich As POINT, ByRef sString As String, Optional bRefresh As Boolean = True)
    Dim rectPaint As RECT
        
    With ptWhich
    SetRect rectPaint, .X * TILE_SIDE, .Y * TILE_SIDE, .X * TILE_SIDE + TILE_SIDE, .Y * TILE_SIDE + TILE_SIDE
    End With
    
    TextOut picMap.hdc, rectPaint.Left, rectPaint.Top, sString, Len(sString)
    If bRefresh Then RefreshRect picMap.hwnd, rectPaint
End Sub

Private Sub Tile_SetNewPoint(ByRef ptWhichOld As POINT, ByRef ptWhichNew As POINT)
    '//Erases the old start or end tile and makes new one
    '//Erase old
    Tile_Paint ptWhichOld, vbWhite
    aiMap(ptWhichOld.X, ptWhichOld.Y) = TH_NORMAL
    '//Prepare new
    aiMap(ptWhichNew.X, ptWhichNew.Y) = TH_NORMAL
    '//Values are passed by reference
    '//So ptWhichOld is either ptStart or ptNew
    ptWhichOld = ptWhichNew
End Sub

Private Function GetArrayPositionFromLocation(ByVal lX As Long, ByVal lY As Long, ByRef ptReturn As POINT) As Boolean
    '//Returns a point with array position depending on X,Y coord
    ptReturn.X = lX \ TILE_SIDE
    ptReturn.Y = lY \ TILE_SIDE
    
    '//Sometimes values can be out of bounds
    If ptReturn.X >= NUMBER_OF_TILES Or _
       ptReturn.Y >= NUMBER_OF_TILES Or _
       ptReturn.X < 0 Or _
       ptReturn.Y < 0 Then
        GetArrayPositionFromLocation = False
    Else
        GetArrayPositionFromLocation = True
    End If
End Function

Private Sub Populate_FileCombo()
    Dim sPath As String
    Dim sFileFound As String

    cmbOpenFile.Clear
    
    sPath = App.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    sPath = sPath & IIf(TILE_SIDE = 20, "*.a", "*.at")
    
    '//Find all files and add then to the combo
    sFileFound = Dir$(sPath)
    Do While Not sFileFound = ""
        cmbOpenFile.AddItem sFileFound
        sFileFound = Dir$
    Loop
End Sub

Private Sub Populate_GridMap()
    Dim I As Long, J As Long
    Dim lColor As Long
    Dim ptPassed As POINT
    
    picMap.Cls
    
    For I = LBound(aiMap, 1) To UBound(aiMap, 1)
        For J = LBound(aiMap, 2) To UBound(aiMap, 2)
            SetPoint ptPassed, I, J
            lColor = GetColorFromHardness(aiMap(I, J))
            Tile_Paint ptPassed, lColor, False
        Next
    Next
    
    '//Start & End
    Tile_Paint ptStart, vbGreen, False
    Tile_Paint ptEnd, vbRed, False
    picMap.Refresh
End Sub

Private Sub mnuFileClear_Click()
    Dim I As Long, J As Long
    
    If Not (objTree Is Nothing) Then Exit Sub '// Is running
    '//Fill map with normal tiles
    For I = LBound(aiMap, 1) To UBound(aiMap, 1)
        For J = LBound(aiMap, 2) To UBound(aiMap, 2)
            aiMap(I, J) = TH_NORMAL
        Next
    Next
    SetPoint ptStart, 0&, 0&
    SetPoint ptEnd, NUMBER_OF_TILES - 1, NUMBER_OF_TILES - 1
    
    Populate_GridMap '//Refresh
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuFileRandom_Click()
    Dim I As Long, J As Long
    Dim lCase&

    If Not (objTree Is Nothing) Then Exit Sub '// Is running
    
    '//Fill map with random tiles
    For I = LBound(aiMap, 1) To UBound(aiMap, 1)
        For J = LBound(aiMap, 2) To UBound(aiMap, 2)
            lCase = Random(0, 80)
            Select Case lCase
                Case Is < 10: aiMap(I, J) = TH_EASY
                Case Is < 60: aiMap(I, J) = TH_NORMAL
                Case Is < 70: aiMap(I, J) = TH_HARD
                Case Is < 75: aiMap(I, J) = TH_VERYHARD
                Case Else: aiMap(I, J) = TH_UNWALKABLE
            End Select
        Next
    Next
    SetPoint ptStart, Random(0&, NUMBER_OF_TILES - 1), Random(0&, NUMBER_OF_TILES - 1)
    SetPoint ptEnd, Random(0&, NUMBER_OF_TILES - 1), Random(0&, NUMBER_OF_TILES - 1)
    
    Populate_GridMap '//Refresh
End Sub

Private Sub mnuFileRedraw_Click()
    If Not (objTree Is Nothing) Then Exit Sub '// Is running
    Populate_GridMap '//Refresh
End Sub

Private Sub mnuFileTiny_Click()
    If Not (objTree Is Nothing) Then Exit Sub '// Is running
    mnuFileTiny.Checked = Not mnuFileTiny.Checked
    
    If mnuFileTiny.Checked Then
        NUMBER_OF_TILES = 48
        TILE_SIDE = 10
    Else
        NUMBER_OF_TILES = 24
        TILE_SIDE = 20
    End If
    
    Map_Setup
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ptArray As POINT
    Dim lPenMode As Long, lTileHardness As Long
    Dim lColor&

    '//Clear & refresh ?
    If bDirty Then
        Populate_GridMap
        bDirty = False
    End If

    '//Get array location
    GetArrayPositionFromLocation X, Y, ptArray
    
    '//Which pen mode is used ?
    If optTile(0).Value Then
        lPenMode = PM_SETTILE
    ElseIf optTile(1).Value Then
        lPenMode = PM_SETSTART
    ElseIf optTile(2).Value Then
        lPenMode = PM_SETEND
    End If
    
    '//Don't replace ptStart or ptEnd by another point.
    If ComparePoints(ptArray, ptStart) Then Exit Sub
    If ComparePoints(ptArray, ptEnd) Then Exit Sub
    
    '//Act depending on pen mode
    Select Case lPenMode
        Case PM_SETSTART
            Tile_SetNewPoint ptStart, ptArray
            Tile_Paint ptArray, vbGreen
        
        Case PM_SETEND
            Tile_SetNewPoint ptEnd, ptArray
            Tile_Paint ptArray, vbRed
        
        Case PM_SETTILE
            lTileHardness = cmbTypeOfTile.ItemData(cmbTypeOfTile.ListIndex)
            aiMap(ptArray.X, ptArray.Y) = lTileHardness
            lColor = GetColorFromHardness(lTileHardness)
            Tile_Paint ptArray, lColor
    End Select
End Sub

'*******************************************************************************************
'Pathfinding
'*******************************************************************************************

Private Sub mnuRunMode_Click(Index As Integer)
    Dim lStart As Long, lEnd As Long, lTime As Long
    Dim bChoosePath As Boolean
    
    If Not (objTree Is Nothing) Then Exit Sub '// Is running
    
    '//Clear all possible previous pathfinding searches
    picMap.Cls
    Populate_GridMap

    lNodeLooped = 0&
    txtTime = ""
    txtNodes = ""
    bDirty = True
    MousePointer = vbHourglass
    DoEvents
        
    Set objTree = New clsTree
    objTree.bAStar = CBool(Index)

    lStart = GetTickCount '// Start timer
    '// To time just the pathfinding, not the drawing
    '// You should turn off the "draw" function
    '// There is a "doevents" in the search to prevent blocking the system.
    bChoosePath = Path_Search
    '
    lEnd = GetTickCount   '// End timer
    lTime = lEnd - lStart
    '
    If bChoosePath Then Path_ChoosePath
    Set objTree = Nothing
    
    txtTime = FormatTime(lTime)
    txtNodes = lNodeLooped
    
    MousePointer = vbNormal
End Sub

Private Function Path_Search() As Boolean
    Dim lNodeX As Long, lNodeY As Long
    
    Dim bShowPath As Boolean, bSlowDown As Boolean
    Dim bAchieved As Boolean
    
    '//Prepare
    bShowPath = (chkShowPathTaken.Value)
    bSlowDown = (chkRunSlow.Value)
    
    '// Speed isn't a factor when showing the path that is going to be taken,
    '// since the drawing routine is what slows things down.
    If bShowPath Then
        objTree.StartSearch ptStart.X, ptStart.Y, ptEnd.X, ptEnd.Y, aiMap()
        
        Do Until objTree.NextNode Or bAchieved
        
            lNodeLooped = lNodeLooped + 1            'Measures
            If bSlowDown Then Sleep SLOW_DOWN_VALUE  'Measures
        
            bAchieved = objTree.UpdateCurrentNode
        
            If bShowPath Then
                objTree.GetNode lNodeX, lNodeY
                Path_PaintNode lNodeX, lNodeY
            End If
            DoEvents
        Loop
        objTree.BackTracePath
        Path_Search = bAchieved
    Else
        '// Not showing the path, find out the fastest path.
        Path_Search = objTree.RunSearch(ptStart.X, ptStart.Y, ptEnd.X, ptEnd.Y, aiMap())
    End If
End Function

Private Sub Path_PaintNode(ByVal lX As Long, ByVal lY As Long, Optional bColorIt As Boolean = False)
    Dim pt As POINT
    pt.X = lX
    pt.Y = lY
    If bColorIt Then Tile_Paint pt, vbYellow
    If (TILE_SIDE = 10) Then
        Tile_Print pt, "x" '// Don't write numbers in tiny mode
    Else
        Tile_Print pt, objTree.GetNodeValue(lX, lY)
    End If
End Sub

Private Sub Path_ChoosePath()                 'The path is ready to be traced by now
    Dim lX As Long, lY As Long, rc As Long
        
    lX = ptStart.X
    lY = ptStart.Y
    Do
      rc = objTree.PathStepNext(lX, lY)       'now step forward one at a time
      Path_PaintNode lX, lY, True
    Loop Until rc = 0                         'the last node has no link forward (like a string termination)
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ptNew As POINT
    Static ptOld As POINT

    '//Get array location
    If Button <> vbLeftButton Then Exit Sub
    '//Sometimes values can be out of bounds, GetArrayPosition thus returns false.
    If Not GetArrayPositionFromLocation(X, Y, ptNew) Then Exit Sub
    If ptNew.X = ptOld.X And ptNew.Y = ptOld.Y Then Exit Sub
    ptOld = ptNew
    picMap_MouseDown Button, Shift, X, Y
End Sub
