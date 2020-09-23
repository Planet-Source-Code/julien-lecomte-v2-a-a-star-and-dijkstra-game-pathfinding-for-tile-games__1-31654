VERSION 5.00
Begin VB.Form frmDraw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast draw tile map"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   430
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   422
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEmpty 
      Caption         =   "Empty map"
      Height          =   495
      Left            =   5040
      TabIndex        =   23
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate redundant random map"
      Height          =   735
      Index           =   1
      Left            =   5040
      TabIndex        =   22
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnload 
      Cancel          =   -1  'True
      Caption         =   "&Close this page"
      Height          =   495
      Left            =   5040
      TabIndex        =   20
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdErase 
      Caption         =   "&Erase"
      Height          =   495
      Left            =   5040
      TabIndex        =   19
      Top             =   5280
      Width           =   1215
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton cmdDrawTiles 
      Caption         =   "&Run a speed test"
      Height          =   495
      Index           =   2
      Left            =   5040
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtTime 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrawTiles 
      Caption         =   "&Double array draw tile map"
      Height          =   495
      Index           =   1
      Left            =   5040
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDrawTiles 
      Caption         =   "&Loop through draw tile map"
      Height          =   495
      Index           =   0
      Left            =   5040
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate random map"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraTilesUsed 
      Caption         =   "Tiles used"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   2175
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   7
         Left            =   1560
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":0000
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   6
         Left            =   1560
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":0C44
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   5
         Left            =   1080
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":1886
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   4
         Left            =   1080
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":24CA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   3
         Left            =   600
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":310E
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   2
         Left            =   600
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":3D50
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   1
         Left            =   120
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":4994
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   3
         Top             =   720
         Width           =   480
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   480
         Index           =   0
         Left            =   120
         LinkTimeout     =   0
         Picture         =   "frmDraw.frx":55D8
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   4800
      Left            =   120
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   336
      X2              =   416
      Y1              =   143
      Y2              =   143
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   0
      X1              =   336
      X2              =   416
      Y1              =   144
      Y2              =   144
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDraw.frx":621C
      Height          =   1335
      Left            =   2400
      TabIndex        =   21
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Time taken"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Time taken"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************************
'Tile map drawing
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************
'
'
'NOTE :
' All picture boxes have :
' .AutoRedraw = true
' .ClipControls = false '// Cause it's useless to have it to true when autoredraw = true

Private Const NUMBER_OF_TILES_PER_SIDE = 10 '//per side of map
Private Const TILE_SIDE_SIZE = 32   '//Pixels

Private aiMap%(NUMBER_OF_TILES_PER_SIDE - 1, NUMBER_OF_TILES_PER_SIDE - 1) '//Integer map of 10 x 10 tiles

Private Sub cmdDrawTiles_Click(Index As Integer)
    Select Case Index
        Case 0: DrawTiles_LoopMethod
        Case 1: DrawTiles_DoubleArrayMethod
        Case 2: Compare_Speeds
    End Select
End Sub

Private Sub cmdErase_Click()
    picMap.Cls
End Sub

Private Sub DrawTiles_LoopMethod()
    Dim I&, J&
    Dim lX&, lY&
    Dim lIndex&
    
    For I = 0 To NUMBER_OF_TILES_PER_SIDE - 1
        For J = 0 To NUMBER_OF_TILES_PER_SIDE - 1
            '//Which tile should be drawn
            lIndex = aiMap(I, J)
            '//Usually, the tiles are loaded from
            '//  - either a ressource file (.picture = LoadResPicture)
            '//  - from a file             (.picture = LoadPicture)
            '//  - from a self made ressource file ...
            '// Since this is not the case, because I have stored all the bitmaps on the
            '// form, I will simulate the loading of the picture by using LoadPicture.
            '// The picture to be bitblt'ed will be from picTemp, and the bitmap in picTemp
            '// will have been loaded from the picTile() objects.
            picTemp.Picture = picTile(lIndex).Picture
            '// Get the location
            lX = I * TILE_SIDE_SIZE
            lY = J * TILE_SIDE_SIZE
            '// Now do the bitblt
            BitBlt picMap.hdc, lX, lY, TILE_SIDE_SIZE, TILE_SIDE_SIZE, picTemp.hdc, 0&, 0&, vbSrcCopy
        Next
    Next
    picMap.Refresh
End Sub

Private Sub DrawTiles_DoubleArrayMethod()
    Dim I&, J&, I2&, J2&
    Dim lX&, lY&
    Dim lIndex&
    Dim aTileDrawn(NUMBER_OF_TILES_PER_SIDE - 1, NUMBER_OF_TILES_PER_SIDE - 1) As Boolean

    For I = 0 To NUMBER_OF_TILES_PER_SIDE - 1
        For J = 0 To NUMBER_OF_TILES_PER_SIDE - 1
            
            '//Did we alreay bitblt this tile ?
            If Not (aTileDrawn(I, J)) Then
                '//No
                '//Which tile should be drawn
                lIndex = aiMap(I, J)
                '//Usually, the tiles are loaded from
                '//  - either a ressource file (.picture = LoadResPicture)
                '//  - from a file             (.picture = LoadPicture)
                '//  - from a self made ressource file ...
                '// Since this is not the case, because I have stored all the bitmaps on the
                '// form, I will simulate the loading of the picture by using LoadPicture.
                '// The picture to be bitblt'ed will be from picTemp, and the bitmap in picTemp
                '// will have been loaded from the picTile() objects.
                '//
                '// We'll only load the picture once
                picTemp.Picture = picTile(lIndex).Picture
                For I2 = 0 To NUMBER_OF_TILES_PER_SIDE - 1
                    For J2 = 0 To NUMBER_OF_TILES_PER_SIDE - 1
                        If Not (aTileDrawn(I2, J2)) Then
                            '//Should we bitblt this picture or another ?
                            If lIndex = aiMap(I2, J2) Then
                                '//This current one is the right one.
                                '// Get the location
                                lX = I2 * TILE_SIDE_SIZE
                                lY = J2 * TILE_SIDE_SIZE
                                '// Now do the bitblt
                                BitBlt picMap.hdc, lX, lY, TILE_SIDE_SIZE, TILE_SIDE_SIZE, picTemp.hdc, 0&, 0&, vbSrcCopy
                                '//Set it to drawn
                                aTileDrawn(I2, J2) = True
                           End If
                        End If
                    Next
                Next
            End If
            
        Next
    Next
    picMap.Refresh
End Sub

Private Sub Compare_Speeds()
    Dim I&
    Dim lLoop_Start&, lLoop_End&, lLoop_Difference&
    Dim lArray_Start&, lArray_End&, lArray_Difference&
    '//The Number of loops is very high since the map is very small.
    '//Of course, when the size of the map gets bigger, the process of tiling gets longer.
    '//One the same basis, when the number of tiles used gets to be very high, it runs slower.
    '//Another factor : when the map uses a lot of the same tiles throughout the array
    '//the array method runs much faster then the loop method.
    '//
    '//Remember, this example is a 10 x 10 tile array for a 320x320 pixel resolution (32pixel tiles)
    '// Resolution   | Tiles   | Number of tiles
    '// 320 x 320    | 10 x 10 | 100
    '// 640 x 480    | 20 x 15 | 300
    '// 800 x 600    | 25 x 20 | 500
    '// etc... (actually the array for a horizontal scrolling map could be 20 x 400 tiles
    
    Const NUMBER_OF_LOOPS = 250
    
    MousePointer = vbHourglass
    
    '//Run each map drawing method NUMBER_OF_LOOPS times
    lLoop_Start = GetTickCount
    For I = 1 To NUMBER_OF_LOOPS
        DrawTiles_LoopMethod
    Next
    lLoop_End = GetTickCount
    
    lArray_Start = GetTickCount
    For I = 1 To NUMBER_OF_LOOPS
        DrawTiles_DoubleArrayMethod
    Next
    lArray_End = GetTickCount
    
    '//Show time taken
    lLoop_Difference = lLoop_End - lLoop_Start
    lArray_Difference = lArray_End - lArray_Start
    txtTime(0) = FormatTime(lLoop_Difference)
    txtTime(1) = FormatTime(lArray_Difference)
    
    MousePointer = vbNormal
End Sub

Private Sub cmdGenerate_Click(Index As Integer)
    '// 25/09/01 Fixed a *stupid* bug :
    '// A redundant map was generated every time
    Dim I&, J&
    Dim lbnd&
    
    Erase aiMap
    
    '//Use only first 3 tiles if redundant
    lbnd = IIf(Index = 1, 2, 7)
    
    '//Generate a random map
    For I = 0 To NUMBER_OF_TILES_PER_SIDE - 1
        For J = 0 To NUMBER_OF_TILES_PER_SIDE - 1
            aiMap(I, J) = Random(0, lbnd)
        Next
    Next
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub cmdEmpty_Click()
    Erase aiMap()
End Sub

Private Sub Form_Load()
    '//Automatic generation of a map
    cmdGenerate_Click 0
End Sub

