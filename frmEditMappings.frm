VERSION 5.00
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmEditMappings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tileset Editor"
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
   Icon            =   "frmEditMappings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Left            =   5160
      Top             =   4200
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
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
      Begin VB.CommandButton cmdPickTile 
         Caption         =   "Pick Tile..."
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   1455
      End
      Begin VB.ComboBox cboTileset 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtTileId 
         Height          =   315
         Left            =   1920
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtDelay 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtFrames 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   1200
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
         Top             =   2160
         Width           =   540
      End
      Begin VB.TextBox txtTilesetInfo 
         BackColor       =   &H8000000F&
         Height          =   3015
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label lblTilesetLbl 
         Caption         =   "Tileset:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblTileIdLbl 
         Caption         =   "Tile ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   750
         Width           =   855
      End
      Begin VB.Label lblDelayLbl 
         Caption         =   "Animation Delay:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1710
         Width           =   1695
      End
      Begin VB.Label lblFramesLbl 
         Caption         =   "Number of Frames:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1230
         Width           =   1815
      End
      Begin VB.Label lblPreviewLbl 
         Caption         =   "Preview:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2280
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
   Begin vbalTreeViewLib6.vbalTreeView tvMappings 
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
      Caption         =   "All Tilesets/Mappings in Current Map:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "Add"
      Visible         =   0   'False
      Begin VB.Menu mnuAddTileset 
         Caption         =   "Tile &Set"
      End
      Begin VB.Menu mnuAddTilesetReplace 
         Caption         =   "Tile Set (&Replace Existing)"
      End
      Begin VB.Menu mnuAddMapping 
         Caption         =   "Tile &Mapping"
      End
   End
End
Attribute VB_Name = "frmEditMappings"
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
tvMappings.Nodes.Clear
cboTileset.Clear

Set myManager = mgr

'Enumerate all the tilesets
Dim i As Long
For i = 0 To mgr.numTilesets - 1
    If mgr.TileSetName(i) <> "" Then
        tvMappings.Nodes.Add , , "tileset" & CStr(i), mgr.TileSetName(i)
        cboTileset.AddItem mgr.TileSetName(i)
    End If
Next
'Enumerate all the mappings
For i = 0 To mgr.NumMappings - 1
    Dim tilesetIdx As Long
    Dim tileIdx As Long
    tilesetIdx = mgr.tileSetIndex(i)
    tileIdx = mgr.TileSetTileIndex(i)
    tvMappings.Nodes.Add tvMappings.Nodes("tileset" & CStr(tilesetIdx)), etvwChild, "mapping" & CStr(i), "Mapping#" & CStr(i) & " Tile#" & CStr(tileIdx)
Next
cboTileset.ListIndex = 0
End Sub

Private Sub cboTileset_Click()
If txtTileId.Text <> "" Then
    picPreview.Cls
    myManager.TileSet(cboTileset.ListIndex).PasteTile picPreview.hDC, 0, 0, CLng(txtTileId.Text)
    picPreview.Refresh
End If
End Sub

Private Sub cmdAdd_Click()
Me.PopupMenu mnuAdd
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo oops
'Delete selected item
Dim node As cTreeViewNode
Set node = tvMappings.SelectedItem
Dim strKey As String, lngId As Long
'Parse the mappingid out of the node
strKey = Replace(node.Key, "mapping", "", , , vbTextCompare)
'Or maybe it's not a mapping id
If Left(strKey, 7) = "tileset" Then
    strKey = Replace(node.Key, "tileset", "", , , vbTextCompare)
    lngId = CLng(strKey)
    'Ask to be sure
    If MsgBox("Are you sure you want to delete tileset " & myManager.TileSetName(lngId) & "?" & vbNewLine & "You will need to repair the mapping indexes for your map to display properly.", vbYesNo Or vbQuestion, "Delete Tileset") = vbYes Then
        'Now delete the selected tileset
        myManager.RemoveTileSet myManager.TileSetName(lngId)
        SetNewMgr myManager
    End If
Else
    'Delete selected mapping
    lngId = CLng(strKey)
    MsgBox "Sorry, tile mappings cannot be deleted.", vbExclamation, "Can't Delete Tile Mapping"
End If
Exit Sub
oops:
GenerateError "Delete operation failed.", "frmEditMappings"
End Sub

Private Sub cmdSave_Click()
On Error GoTo oops
Dim node As cTreeViewNode
Set node = tvMappings.SelectedItem
'Fill in the data on the right
Dim strKey As String, lngId As Long
'Parse the mappingid out of the node
strKey = Replace(node.Key, "mapping", "", , , vbTextCompare)
lngId = CLng(strKey)
'Save the data into the manager
myManager.tileSetIndex(lngId) = CLng(cboTileset.ListIndex)
myManager.TileSetTileIndex(lngId) = CLng(txtTileId.Text)
myManager.numTileReplacements(lngId) = CLng(txtFrames.Text)
myManager.cycleDelay(lngId) = CLng(txtDelay.Text)
SetNewMgr myManager
Exit Sub
oops:
GenerateError "Save Mapping operation failed.", "frmEditMappings"
End Sub

Private Sub Form_Unload(cancel As Integer)
Set myManager = Nothing
End Sub

Private Sub mnuAddMapping_Click()
On Error GoTo oops
'Adds a new mapping to the tileset
Dim lngMapping As Long
'First check to be sure there's a tileset loaded first
If myManager.numTilesets = 0 Then
    MsgBox "You must add at least one tileset before creating a mapping.", vbExclamation, "No Tilesets Defined"
    Exit Sub
End If
lngMapping = myManager.MapInTiles(0, 0, 1)
MsgBox "Mapping #" & CStr(lngMapping) & " was created successfully."
SetNewMgr myManager
Exit Sub
oops:
GenerateError "Add Mapping operation failed.", "frmEditMappings"
End Sub

Private Sub mnuAddTileset_Click()
On Error GoTo oops
'Adds a tileset to the map
Dim sTilesetName As String, setId As Long, i As Long
'**TODO** This is lazy...
sTilesetName = InputBox("Enter the tileset name below. The tileset must be a .bmp file in the OP2 folder in either standard Windows bitmap or PBMP formats. Do not add .bmp to the end, paths are not allowed.", "Add Tileset", "well0000")
If Trim(sTilesetName) = "" Then Exit Sub
setId = myManager.AddTileSet(sTilesetName)
'Ask the user if they want to create mappings for all the tiles
If MsgBox("Do you want to create initial mappings for each tile in this tileset?", vbYesNo Or vbQuestion, "Create Initial Mappings") = vbYes Then
    'Create mappings
    For i = 0 To myManager.TileSet(setId).NumTiles - 1
        myManager.MapInTiles setId, i, 1
    Next
    MsgBox CStr(myManager.TileSet(setId).NumTiles) & " mappings have been created.", vbInformation, "Initial Mappings Created"
End If
SetNewMgr myManager
Exit Sub
oops:
GenerateError "Add Tileset operation failed.", "frmEditMappings"
End Sub

Private Sub mnuAddTilesetReplace_Click()
On Error GoTo oops
'First figure out the id of the tileset
Dim node As cTreeViewNode
Set node = tvMappings.SelectedItem
Dim strKey As String, lngId As Long
'Parse the mappingid out of the node
strKey = Replace(node.Key, "mapping", "", , , vbTextCompare)
'Or maybe it's not a mapping id
If Left(strKey, 7) = "tileset" Then
    strKey = Replace(node.Key, "tileset", "", , , vbTextCompare)
    lngId = CLng(strKey)
    'Replace the tileset
    Dim sTilesetName As String
    '**TODO** This is lazy...
    sTilesetName = InputBox("Enter the tileset name below. The tileset must be a .bmp file in the OP2 folder in either standard Windows bitmap or PBMP formats. Do not add .bmp to the end, paths are not allowed.", "Replace Tileset", "well0000")
    If Trim(sTilesetName) = "" Then Exit Sub
    myManager.ReplaceTileSet lngId, sTilesetName
    SetNewMgr myManager
Else
    'Can't do it
    lngId = CLng(strKey)
    MsgBox "Sorry, this operation only works on tilesets.", vbExclamation, "Can't Perform Operation"
End If
Exit Sub
oops:
GenerateError "Replace Tileset operation failed.", "frmEditMappings"
End Sub

Private Sub tmrAnim_Timer()
'Update the preview with the next frame
curIteration = curIteration + 1
If curIteration > myManager.numTileReplacements(curMapping) Then curIteration = 0
myManager.TileSet(myManager.tileSetIndex(curMapping)).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(curMapping) + curIteration
picPreview.Refresh
End Sub

Private Sub tvMappings_SelectedNodeChanged()
Dim node As cTreeViewNode
Set node = tvMappings.SelectedItem
'Fill in the data on the right
Dim strKey As String, lngId As Long
'Parse the mappingid out of the node
strKey = Replace(node.Key, "mapping", "", , , vbTextCompare)
'Or maybe it's not a mapping id
If Left(strKey, 7) = "tileset" Then
    strKey = Replace(node.Key, "tileset", "", , , vbTextCompare)
    lngId = CLng(strKey)
    'Now go into "tile set mode"
    tmrAnim.Enabled = False
    lblPreviewLbl.Visible = False
    picPreview.Visible = False
    picPreview.Cls
    lblTilesetLbl.Visible = False
    cboTileset.Visible = False
    lblTileIdLbl.Visible = False
    txtTileId.Visible = False
    txtTileId.Text = ""
    lblFramesLbl.Visible = False
    txtFrames.Visible = False
    txtFrames.Text = ""
    lblDelayLbl.Visible = False
    txtDelay.Visible = False
    txtDelay.Text = ""
    cmdSave.Visible = False
    cmdPickTile.Visible = False
    
    txtTilesetInfo.Visible = True
    txtTilesetInfo.Text = "Tileset Info" & vbNewLine & _
        "------------" & vbNewLine & _
        "Index: " & CStr(lngId) & vbNewLine & _
        "Name: " & myManager.TileSetName(lngId) & vbNewLine & _
        "Num Tiles: " & CStr(myManager.TileSet(lngId).NumTiles) & vbNewLine & _
        "Num Palette Entries: " & CStr(myManager.TileSet(lngId).NumPaletteEntries) & vbNewLine & _
        "Bit Depth: " & CStr(myManager.TileSet(lngId).BitDepth) & vbNewLine & _
        "Tile Size: " & CStr(myManager.TileSet(lngId).TileSize)

Else
    'Mapping mode
    lngId = CLng(strKey)
    lblTilesetLbl.Visible = True
    cboTileset.Visible = True
    cboTileset.ListIndex = myManager.tileSetIndex(lngId)
    lblTileIdLbl.Visible = True
    txtTileId.Visible = True
    txtTileId.Text = CStr(myManager.TileSetTileIndex(lngId))
    lblFramesLbl.Visible = True
    txtFrames.Visible = True
    txtFrames.Text = CStr(myManager.numTileReplacements(lngId))
    lblDelayLbl.Visible = True
    txtDelay.Visible = True
    txtDelay.Text = CStr(myManager.cycleDelay(lngId))
    cmdSave.Visible = True
    cmdPickTile.Visible = True
    
    txtTilesetInfo.Visible = False
    'Figure out whether to use the animation timer
    If myManager.numTileReplacements(lngId) <> 0 Then
        curMapping = lngId
        curIteration = 0
        'Scale it to a reasonable value - so that the user can see the animation
        tmrAnim.Interval = myManager.cycleDelay(lngId) * 25
        tmrAnim.Enabled = True
    Else
        tmrAnim.Enabled = False
    End If
    'Draw the preview in the box
    lblPreviewLbl.Visible = True
    picPreview.Visible = True
    picPreview.Cls
    myManager.TileSet(myManager.tileSetIndex(lngId)).PasteTile picPreview.hDC, 0, 0, myManager.TileSetTileIndex(lngId)
    picPreview.Refresh
End If
End Sub

Public Sub SetNewName(ByVal strName As String)
Me.Caption = "Tileset Editor - " & strName
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

Private Sub txtTileId_Change()
If txtTileId.Text <> "" Then
    picPreview.Cls
    myManager.TileSet(cboTileset.ListIndex).PasteTile picPreview.hDC, 0, 0, CLng(txtTileId.Text)
    picPreview.Refresh
End If
End Sub
