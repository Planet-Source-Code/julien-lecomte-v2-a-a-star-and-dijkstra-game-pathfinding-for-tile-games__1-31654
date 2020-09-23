VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tile Examples"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   2490
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   2490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Quit"
      Height          =   375
      Left            =   75
      TabIndex        =   2
      Top             =   975
      Width           =   2295
   End
   Begin VB.CommandButton cmdPathFinding 
      Caption         =   "Pathfinding"
      Height          =   375
      Left            =   75
      TabIndex        =   1
      Top             =   525
      Width           =   2295
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "Fast draw tile map"
      Height          =   375
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************************************
'Tile map examples
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************

Private Sub cmdDraw_Click()
    frmDraw.Show vbModal, Me
End Sub

Private Sub cmdPathFinding_Click()
    frmPathfinding.Show vbModal, Me
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Randomize
End Sub
