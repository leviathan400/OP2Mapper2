VERSION 5.00
Begin VB.Form frmTilesetModal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Tile"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmTilesetModal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboTileset 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   3195
   End
   Begin VB.VScrollBar vsbScroll 
      Height          =   2940
      LargeChange     =   6
      Left            =   2940
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2940
      Left            =   0
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   0
      Top             =   360
      Width           =   2940
   End
End
Attribute VB_Name = "frmTilesetModal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLACKNESS = &H42

Public intDc As Long, tmpBm As Long
Private palWidth As Long, palHeight As Long, nTiles As Long, dcExists As Boolean

Private realTileSets As Long

Private thisMgr As TileSetManager

Public enableGrid As Boolean

Public curTileset As Long
Public curTileID As Long

Private Sub Form_Load()
curTileset = -1
curTileID = -1
If enableGridDef Then enableGrid = True
End Sub

'Redraw the available tiles based on the tile set manager
Public Sub SetNewMgr(mgr As TileSetManager)
'Me.Hide
realTileSets = 0
palWidth = 192 '6 tiles across
Dim i As Integer, j As Integer
'Determine the number of tiles

Set thisMgr = mgr

cboTileset.Clear

For i = 0 To mgr.numTilesets - 1
    If mgr.tileset(i) Is Nothing Then Exit For
    nTiles = nTiles + mgr.tileset(i).NumTiles
    realTileSets = realTileSets + 1
    cboTileset.AddItem mgr.TileSetName(i)
Next

cboTileset.ListIndex = 0
ChangeSet 0
End Sub

Public Sub ChangeSet(ByVal idx As Long)
Dim i As Integer

If thisMgr Is Nothing Then Exit Sub

'Determine tile palette height (numrows*32)
If thisMgr.tileset(idx).NumTiles < 7 Then
    palHeight = 32
Else
    palHeight = (Int(thisMgr.tileset(idx).NumTiles / 6) + 1) * 32
End If

'Destroy any preexisting memory DC and bitmap
If dcExists Then
    DeleteDC intDc
    DeleteObject tmpBm
End If
dcExists = True
'Create a DC and bitmap compatible with the picture box
intDc = CreateCompatibleDC(picTiles.hDC)
tmpBm = CreateCompatibleBitmap(picTiles.hDC, palWidth, palHeight)
'Out with the old... in with the new
DeleteObject SelectObject(intDc, tmpBm)

'Color fill it to black
BitBlt intDc, 0, 0, palWidth, palHeight, 0, 0, 0, BLACKNESS

Dim curX As Long, curY As Long
'Draw the tiles into the DC
curX = 0: curY = 0
'For i = 0 To mgr.NumTileSets - 1
For i = 0 To thisMgr.tileset(idx).NumTiles - 1
    thisMgr.tileset(idx).PasteTile intDc, curX, curY, i
    curX = curX + 32
    If curX > 160 Then curX = 0: curY = curY + 32
Next

'Resize the scroll bars
If Int(thisMgr.tileset(idx).NumTiles / (palWidth / 32)) < Int(picTiles.ScaleHeight / 32) + 1 Then
    vsbScroll.Max = 0
Else
    vsbScroll.Max = (palHeight / 32) - Int(picTiles.ScaleHeight / 32)
End If
RedrawSelf
End Sub

Private Sub cboTileset_Click()
ChangeSet cboTileset.ListIndex
End Sub

Private Sub Form_Unload(cancel As Integer)
'Clean up bitmaps
If dcExists Then
    DeleteDC intDc
    DeleteObject tmpBm
End If
Set thisMgr = Nothing
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long, id As Long
'Figure out the tile id of whatever was clicked
i = Int(x / 32)
j = Int(y / 32) + vsbScroll.Value
id = j * 6 + i
curTileID = id
curTileset = cboTileset.ListIndex
Me.Hide
End Sub

Private Sub vsbScroll_Change()
RedrawSelf
End Sub

Private Sub vsbScroll_Scroll()
vsbScroll_Change
End Sub

Public Sub RedrawSelf()
'Redraw
picTiles.Cls
BitBlt picTiles.hDC, 0, 0, 192, 384, intDc, 0, vsbScroll.Value * 32, vbSrcCopy
'Add gridlines
Dim i As Long
If enableGrid Then
    'Add gridlines
    For i = 1 To Int(picTiles.Width / 32)
        picTiles.Line (i * 32, 0)-(i * 32, picTiles.Height), vbBlack
    Next
    For i = 1 To Int(picTiles.Height / 32)
        picTiles.Line (0, i * 32)-(picTiles.Width, i * 32), vbBlack
    Next
End If
picTiles.Refresh
End Sub
