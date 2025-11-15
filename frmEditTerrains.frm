VERSION 5.00
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmEditTerrains 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Terrain Editor"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
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
   Icon            =   "frmEditTerrains.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame grpCurrentSettings 
      Caption         =   "Current Selection Settings"
      Height          =   3375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.PictureBox picHolder 
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   3015
         Left            =   120
         ScaleHeight     =   3015
         ScaleWidth      =   3255
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
         Begin VB.Label lblNoSet 
            Alignment       =   2  'Center
            Caption         =   "No settings for this item."
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   1320
            Width           =   3255
         End
      End
      Begin VB.ComboBox cboTileset 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtTileId 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.PictureBox picPreview 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   540
         Left            =   2760
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   5
         Top             =   1200
         Width           =   540
      End
      Begin VB.CommandButton cmdPickTile 
         Caption         =   "Tile Picker..."
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblTilesetLbl 
         Caption         =   "Tileset:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblTileIdLbl 
         Caption         =   "Tile ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   750
         Width           =   855
      End
      Begin VB.Label lblPreviewLbl 
         Caption         =   "Preview:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
   End
   Begin vbalTreeViewLib6.vbalTreeView tvTerrains 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5530
      HotTracking     =   0   'False
      LineStyle       =   0
      SingleSel       =   -1  'True
      Style           =   2
      ScaleMode       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblWarn 
      Caption         =   "Warning: Be careful in this dialog. Changes made take effect immediately on the current (unsaved) map."
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   6975
   End
   Begin VB.Label lblMappings 
      Caption         =   "All Terrains in Current Map:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmEditTerrains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myManager As TileSetManager
Private curMapping As Long, curIteration As Long

Public Sub SetNewMgr(mgr As TileSetManager)
On Error Resume Next
'Assigns a new manager to this form
tvTerrains.Nodes.Clear
cboTileset.Clear

Set myManager = mgr

'Enumerate all the terrains
Dim i As Long, j As Long, tmpNode As cTreeViewNode, tmpNode2 As cTreeViewNode
For i = 0 To mgr.numTerrains - 1
    Set tmpNode = tvTerrains.Nodes.Add(, , CStr(i), "Terrain " & CStr(i + 1))
    'Add the different items for each terrain object
    tmpNode.AddChildNode CStr(i) & ",start", "Starting Tile"
    tmpNode.AddChildNode CStr(i) & ",end", "Ending Tile"
    tmpNode.AddChildNode CStr(i) & ",dozed", "Bulldozed"
    tmpNode.AddChildNode CStr(i) & ",rubble", "Rubble"
    
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",tubeunk", "Tubes (Unknown)")
    AddDirections tmpNode2, False, i, "tubeunk"
    
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",lavawall", "Lava Wall")
    AddDirections tmpNode2, True, i, "lavawall"
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",microbewall", "Microbe Wall")
    AddDirections tmpNode2, True, i, "microbewall"
    
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",normalwall", "Normal Wall")
    AddDirections tmpNode2, True, i, "normalwall"
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",damagedwall", "Damaged Wall")
    AddDirections tmpNode2, True, i, "damagedwall"
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",ruinedwall", "Ruined Wall")
    AddDirections tmpNode2, True, i, "ruinedwall"
    
    tmpNode.AddChildNode CStr(i) & ",lava", "Lava"
    tmpNode.AddChildNode CStr(i) & ",flat1", "Flat 1"
    tmpNode.AddChildNode CStr(i) & ",flat2", "Flat 2"
    tmpNode.AddChildNode CStr(i) & ",flat3", "Flat 3"
    
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",tube", "Tubes")
    AddDirections tmpNode2, True, i, "tube"
    
    tvTerrains.Nodes.Add tmpNode, etvwChild, CStr(i) & ",scorched", "Scorched"
    
    Set tmpNode2 = tmpNode.AddChildNode(CStr(i) & ",unk", "Unknown Tiles")
    For j = 0 To 19
        tmpNode2.AddChildNode CStr(i) & ",unk," & CStr(j), "Unknown " & CStr(j + 1)
    Next
Next
'Enumerate all the tilesets
For i = 0 To mgr.numTilesets - 1
    If mgr.TileSetName(i) <> "" Then cboTileset.AddItem mgr.TileSetName(i)
Next
End Sub

Private Sub cboTileset_Click()
If txtTileId.Text <> "" Then
    picPreview.Cls
    myManager.TileSet(cboTileset.ListIndex).PasteTile picPreview.hDC, 0, 0, CLng(txtTileId.Text)
    picPreview.Refresh
End If
End Sub

Private Sub cmdAdd_Click()
myManager.SetNumTerrains myManager.numTerrains + 1
SetNewMgr myManager
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
MsgBox "Sorry, terrain deletion is not implemented yet."
End Sub

Private Sub cmdPickTile_Click()
'Pick a tile from the tileset picker window
Dim tilesetWnd As New frmTilesetModal
tilesetWnd.SetNewMgr myManager
tilesetWnd.Show 1, Me
If tilesetWnd.curTileset <> -1 And tilesetWnd.curTileID <> -1 Then
    cboTileset.ListIndex = tilesetWnd.curTileset
    txtTileId.Text = CStr(tilesetWnd.curTileID)
    picPreview.Cls
    myManager.TileSet(tilesetWnd.curTileset).PasteTile picPreview.hDC, 0, 0, tilesetWnd.curTileID
    picPreview.Refresh
End If
Unload tilesetWnd
Set tilesetWnd = Nothing
End Sub

Private Sub cmdSave_Click()
On Error GoTo oops
Dim node As cTreeViewNode, tmpArr() As String
Dim terrainId As Long, itemId As String, dirId As Long
Set node = tvTerrains.SelectedItem

terrainId = -1
itemId = ""
dirId = -1

'Parse it and see what it contains
tmpArr = Split(node.Key, ",")
If UBound(tmpArr) >= 0 Then
    terrainId = CLng(tmpArr(0))
    If UBound(tmpArr) >= 1 Then
        itemId = tmpArr(1)
        If UBound(tmpArr) >= 2 Then dirId = CLng(tmpArr(2))
    End If
End If

'Setup the display on the right based on what we have
If itemId = "" Then picHolder.Visible = True: Exit Sub

Select Case itemId
    Case "start"
        picHolder.Visible = False
        myManager.TerrainStartTile(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainStartTile(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainStartTile(terrainId))
    Case "end"
        picHolder.Visible = False
        myManager.TerrainEndTile(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainEndTile(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainEndTile(terrainId))
    Case "dozed"
        picHolder.Visible = False
        myManager.TerrainDozed(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainDozed(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainDozed(terrainId))
    Case "rubble"
        picHolder.Visible = False
        myManager.TerrainRubble(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainRubble(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainRubble(terrainId))
    Case "tubeunk"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainTubeUnk(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainTubeUnk(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainTubeUnk(terrainId, dirId))
    Case "lavawall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainLavaWall(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainLavaWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainLavaWall(terrainId, dirId))
    Case "microbewall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainMicrobeWall(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainMicrobeWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainMicrobeWall(terrainId, dirId))
    Case "normalwall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainNormalWall(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainNormalWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainNormalWall(terrainId, dirId))
    Case "damagedwall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainDamagedWall(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainDamagedWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainDamagedWall(terrainId, dirId))
    Case "ruinedwall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainRuinedWall(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainRuinedWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainRuinedWall(terrainId, dirId))
    Case "lava"
        picHolder.Visible = False
        myManager.TerrainLava(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainLava(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainLava(terrainId))
    Case "flat1"
        picHolder.Visible = False
        myManager.TerrainFlat1(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainFlat1(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainFlat1(terrainId))
    Case "flat2"
        picHolder.Visible = False
        myManager.TerrainFlat2(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainFlat2(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainFlat2(terrainId))
    Case "flat3"
        picHolder.Visible = False
        myManager.TerrainFlat3(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainFlat3(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainFlat3(terrainId))
    Case "tube"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainTube(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainTube(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainTube(terrainId, dirId))
    Case "scorched"
        picHolder.Visible = False
        myManager.TerrainScorched(terrainId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainScorched(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainScorched(terrainId))
    Case "unk"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        myManager.TerrainUnknown(terrainId, dirId) = myManager.GetMappingIndex(cboTileset.ListIndex, CLng(txtTileId.Text), -1, -1)
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainUnknown(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainUnknown(terrainId, dirId))
    Case Else
        MsgBox "**TODO**"
End Select
picPreview.Refresh
Exit Sub
oops:
GenerateError "Save Terrain operation failed.", "frmEditTerrains"
End Sub

Private Sub Form_Unload(cancel As Integer)
Set myManager = Nothing
End Sub


Private Sub tvTerrains_SelectedNodeChanged()
Dim node As cTreeViewNode, tmpArr() As String
Dim terrainId As Long, itemId As String, dirId As Long
Set node = tvTerrains.SelectedItem

terrainId = -1
itemId = ""
dirId = -1

'Parse it and see what it contains
tmpArr = Split(node.Key, ",")
If UBound(tmpArr) >= 0 Then
    terrainId = CLng(tmpArr(0))
    If UBound(tmpArr) >= 1 Then
        itemId = tmpArr(1)
        If UBound(tmpArr) >= 2 Then dirId = CLng(tmpArr(2))
    End If
End If

'Setup the display on the right based on what we have
If itemId = "" Then picHolder.Visible = True: Exit Sub

Select Case itemId
    Case "start"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainStartTile(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainStartTile(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainStartTile(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainStartTile(terrainId))
    Case "end"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainEndTile(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainEndTile(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainEndTile(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainEndTile(terrainId))
    Case "dozed"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainDozed(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainDozed(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainDozed(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainDozed(terrainId))
    Case "rubble"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainRubble(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainRubble(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainRubble(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainRubble(terrainId))
    Case "tubeunk"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainTubeUnk(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainTubeUnk(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainTubeUnk(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainTubeUnk(terrainId, dirId))
    Case "lavawall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainLavaWall(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainLavaWall(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainLavaWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainLavaWall(terrainId, dirId))
    Case "microbewall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainMicrobeWall(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainMicrobeWall(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainMicrobeWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainMicrobeWall(terrainId, dirId))
    Case "normalwall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainNormalWall(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainNormalWall(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainNormalWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainNormalWall(terrainId, dirId))
    Case "damagedwall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainDamagedWall(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainDamagedWall(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainDamagedWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainDamagedWall(terrainId, dirId))
    Case "ruinedwall"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainRuinedWall(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainRuinedWall(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainRuinedWall(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainRuinedWall(terrainId, dirId))
    Case "lava"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainLava(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainLava(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainLava(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainLava(terrainId))
    Case "flat1"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainFlat1(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainFlat1(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainFlat1(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainFlat1(terrainId))
    Case "flat2"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainFlat2(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainFlat2(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainFlat2(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainFlat2(terrainId))
    Case "flat3"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainFlat3(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainFlat3(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainFlat3(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainFlat3(terrainId))
    Case "tube"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainTube(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainTube(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainTube(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainTube(terrainId, dirId))
    Case "scorched"
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainScorched(terrainId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainScorched(terrainId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainScorched(terrainId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainScorched(terrainId))
    Case "unk"
        If dirId = -1 Then picHolder.Visible = True: Exit Sub
        picHolder.Visible = False
        cboTileset.ListIndex = myManager.tileSetIndex(myManager.TerrainUnknown(terrainId, dirId))
        txtTileId.Text = CStr(myManager.TileSetTileIndex(myManager.TerrainUnknown(terrainId, dirId)))
        myManager.TileSet(myManager.tileSetIndex(myManager.TerrainUnknown(terrainId, dirId))).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(myManager.TerrainUnknown(terrainId, dirId))
    Case Else
        MsgBox "**TODO**"
End Select
picPreview.Refresh
End Sub

Public Sub SetNewName(ByVal strName As String)
Me.Caption = "Terrain Editor - " & strName
End Sub

Private Sub AddDirections(tmpNode2 As cTreeViewNode, ByVal fullSet As Boolean, ByVal id As Long, ByVal idStr As String)
    tmpNode2.AddChildNode CStr(id) & "," & idStr & ",0", "Left to Right"
    tmpNode2.AddChildNode CStr(id) & "," & idStr & ",1", "Top to Bottom"
    tmpNode2.AddChildNode CStr(id) & "," & idStr & ",2", "Left to Bottom"
    tmpNode2.AddChildNode CStr(id) & "," & idStr & ",3", "Right to Bottom"
    tmpNode2.AddChildNode CStr(id) & "," & idStr & ",4", "Left to Top"
    tmpNode2.AddChildNode CStr(id) & "," & idStr & ",5", "Right to Top"
    If fullSet = True Then
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",6", "Left-Right-Top"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",7", "Left-Right-Bottom"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",8", "Left-Right-Top-Bottom"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",9", "Left-Top-Bottom"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",10", "Right-Top-Bottom"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",11", "Bottom"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",12", "Top"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",13", "Right"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",14", "Left"
        tmpNode2.AddChildNode CStr(id) & "," & idStr & ",15", "Middle Section"
    End If
End Sub

Private Sub txtTileId_Change()
If txtTileId.Text <> "" Then
    picPreview.Cls
    myManager.TileSet(cboTileset.ListIndex).PasteTile picPreview.hDC, 0, 0, CLng(txtTileId.Text)
    picPreview.Refresh
End If
End Sub
