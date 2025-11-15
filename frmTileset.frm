VERSION 5.00
Begin VB.Form frmTileset 
   Caption         =   "Tile Set"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmTileset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   3195
   Begin VB.ComboBox cboTileset 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   3195
   End
   Begin VB.VScrollBar vsbScroll 
      Height          =   5820
      LargeChange     =   6
      Left            =   2940
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5820
      Left            =   0
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   0
      Top             =   360
      Width           =   2940
   End
End
Attribute VB_Name = "frmTileset"
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

Public ownerForm As MapManager

Public curMappingIndex As Long

Public oneShot As Boolean 'If true, close the window afterwards

'ONLY used for consecutive fill
Public curTileset As Long
Public curTileID As Long

Public prepareClose As Boolean

Private Sub Form_Deactivate()
If ownerForm Is Nothing Then Exit Sub
ownerForm.ToolDeactivated
End Sub

Private Sub Form_Load()
'Set window settings
Me.Left = GetSettingIni("Window", "TilesetLeft", 0)
Me.TOp = GetSettingIni("Window", "TilesetTop", 0)
Me.Height = GetSettingIni("Window", "TilesetHeight", 6585)
Me.WindowState = GetSettingIni("Window", "TilesetWindowState", 0)

If enableGridDef Then enableGrid = True
oneShot = False
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If prepareClose = False Then cancel = True
End Sub

'Redraw the available tiles based on the tile set manager
Public Sub SetNewMgr(mgr As TileSetManager)
realTileSets = 0
palWidth = 192 '6 tiles across
Dim i As Integer, j As Integer
'Determine the number of tiles in all
'Debug.Print "NumMappings: " & mgr.NumMappings
'Debug.Print "NumTileSets: " & mgr.NumTileSets

Set thisMgr = mgr

cboTileset.Clear

For i = 0 To mgr.numTilesets - 1
    If mgr.TileSet(i) Is Nothing Then Exit For
    nTiles = nTiles + mgr.TileSet(i).NumTiles
    realTileSets = realTileSets + 1
    cboTileset.AddItem mgr.TileSetName(i)
Next

cboTileset.ListIndex = 0
ChangeSet 0

''Determine tile palette height (numrows*32)
'palHeight = Int(nTiles / 6) * 32
'
''Destroy any preexisting memory DC and bitmap
'If dcExists Then
'    DeleteDC intDc
'    DeleteObject tmpBm
'End If
'dcExists = True
''Create a DC and bitmap compatible with the picture box
'intDc = CreateCompatibleDC(picTiles.hdc)
'tmpBm = CreateCompatibleBitmap(picTiles.hdc, palWidth, palHeight)
''Out with the old... in with the new
'DeleteObject SelectObject(intDc, tmpBm)
'
''Color fill it to black
'BitBlt intDc, 0, 0, palWidth, palHeight, 0, 0, 0, BLACKNESS
'
'Dim curX As Long, curY As Long
''Draw the tiles into the DC
'curX = 0: curY = 0
''For i = 0 To mgr.NumTileSets - 1
'For i = 0 To realTileSets - 1
'    For j = 0 To mgr.TileSet(i).NumTiles - 1
'        mgr.TileSet(i).PasteTile intDc, curX, curY, j
'        curX = curX + 32
'        If curX > 192 Then curX = 0: curY = curY + 32
'    Next
'Next
'
''Resize the scroll bars
'vsbScroll.Max = (palHeight / 32) - 6
''Redraw
'BitBlt picTiles.hdc, 0, 0, 192, 192, intDc, 0, vsbScroll.value * 32, vbSrcCopy
'picTiles.Refresh
End Sub

Public Sub ChangeSet(ByVal idx As Long)
Dim i As Integer

If thisMgr Is Nothing Then Exit Sub

'Determine tile palette height (numrows*32)
If thisMgr.TileSet(idx).NumTiles < 7 Then
    palHeight = 32
Else
    palHeight = (Int(thisMgr.TileSet(idx).NumTiles / 6) + 1) * 32
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
For i = 0 To thisMgr.TileSet(idx).NumTiles - 1
    thisMgr.TileSet(idx).PasteTile intDc, curX, curY, i
    curX = curX + 32
    If curX > 160 Then curX = 0: curY = curY + 32
Next

'Resize the scroll bars
If Int(thisMgr.TileSet(idx).NumTiles / (palWidth / 32)) < Int(picTiles.ScaleHeight / 32) + 1 Then
    vsbScroll.Max = 0
Else
    vsbScroll.Max = (palHeight / 32) - Int(picTiles.ScaleHeight / 32)
End If
RedrawSelf
End Sub

Private Sub cboTileset_Click()
ChangeSet cboTileset.ListIndex
End Sub

Private Sub Form_Resize()
On Error Resume Next
'Allow the user to resize the window vertically, not horizontally
Me.Width = vsbScroll.Left + vsbScroll.Width + 115
picTiles.Height = Me.ScaleHeight - cboTileset.Height - 50
vsbScroll.Height = Me.ScaleHeight - cboTileset.Height - 50
ChangeSet cboTileset.ListIndex
End Sub

Private Sub Form_Unload(cancel As Integer)
'Save window settings
If Me.WindowState <> vbMinimized Then
    SaveSettingIni "Window", "TilesetLeft", Me.Left
    SaveSettingIni "Window", "TilesetTop", Me.TOp
    SaveSettingIni "Window", "TilesetHeight", Me.Height
    SaveSettingIni "Window", "TilesetWindowState", Me.WindowState
End If
'Clean up bitmaps
If dcExists Then
    DeleteDC intDc
    DeleteObject tmpBm
End If
Set thisMgr = Nothing
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long, id As Long, newIdx As Long

If Button = vbRightButton Then
    'Pop menu
    fMainForm.mnuTilesetEnableGrid.Checked = enableGrid
    Me.SetFocus
    curTileset = cboTileset.ListIndex
    i = Int(x / 32)
    j = Int(y / 32) + vsbScroll.Value
    curTileID = j * 6 + i
    Me.PopupMenu fMainForm.mnuTileset
Else

    'Set the current mapping index to whatever was clicked
    
    'Figure out the tile id of whatever was clicked
    i = Int(x / 32)
    j = Int(y / 32) + vsbScroll.Value
    id = j * 6 + i
    'Figure out the mapping index
    newIdx = thisMgr.GetMappingIndex(cboTileset.ListIndex, id, 0, 0)
    'Make sure that it's valid!
    If newIdx <> -1 Then
        curMappingIndex = newIdx
        SetStatusBar "Selected Tile: Tileset " & CStr(cboTileset.ListIndex) & ", TileID " & CStr(id) & ", Mapping " & CStr(newIdx)
        If oneShot Then Me.Hide
    Else
        'It doesn't have a mapping index, so tell the user that!
        MsgBox "The tile you have selected (TileID " & CStr(id) & ") has no mapping index defined." & vbNewLine & "You must create an index for this tile in order to use it on the map." & vbNewLine & vbNewLine & "Use the Tileset Editor toolbar button in the map to create mapping indexes.", vbInformation, "No Mapping Index"
    End If

End If

End Sub

Private Sub vsbScroll_Change()
RedrawSelf
End Sub

Private Sub vsbScroll_Scroll()
vsbScroll_Change
End Sub

Public Sub SetNewName(ByVal strName As String)
Me.Caption = "Tile Set - " & strName
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

Public Sub SetCurTile(ByVal mappingIndex As Long)
'Set the current mapping index manually
curMappingIndex = mappingIndex
SetStatusBar "Selected Tile: Mapping " & CStr(mappingIndex)
End Sub

Public Sub PasteConsecutive(ByVal startId As Long)
Dim newIdx As Long
'Begin consecutive paste operation
'If starting ID is entered, begin from that
If startId <> -1 Then curTileID = startId
ownerForm.massMode = False: ownerForm.massStep = 0
ownerForm.myToolbar.ButtonChecked("massedit") = False
ownerForm.consecStep = 1
SetStatusBar "Click upper left corner of area to fill with tiles"
End Sub
