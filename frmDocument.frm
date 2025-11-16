VERSION 5.00
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Begin VB.Form MapManager 
   Caption         =   "Map"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
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
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   Begin VB.Timer tmrScroll 
      Interval        =   100
      Left            =   6960
      Top             =   0
   End
   Begin vbalTBar6.cReBar myRebar 
      Left            =   2280
      Top             =   0
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin vbalTBar6.cToolbar myToolbar 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      ScaleHeight     =   325
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   2
      Top             =   480
      Width           =   6975
      Begin VB.Shape shpInnerUnit 
         BorderColor     =   &H0000FFFF&
         Height          =   480
         Left            =   720
         Top             =   0
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Line shpHTube 
         BorderColor     =   &H00FFFFFF&
         Visible         =   0   'False
         X1              =   0
         X2              =   32
         Y1              =   40
         Y2              =   40
      End
      Begin VB.Line shpVTube 
         BorderColor     =   &H00FFFFFF&
         Visible         =   0   'False
         X1              =   40
         X2              =   40
         Y1              =   0
         Y2              =   32
      End
      Begin VB.Shape shpCursor 
         BorderColor     =   &H00FFFFFF&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.HScrollBar hsbScroll 
      Height          =   255
      LargeChange     =   6
      Left            =   0
      TabIndex        =   1
      Top             =   5520
      Width           =   6975
   End
   Begin VB.VScrollBar vsbScroll 
      Height          =   4935
      LargeChange     =   6
      Left            =   7080
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "MapManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Core stuff
Public myMap As MapFile
Public myName As String, myTitle As String, isFile As Boolean
Public toolMode As ToolModeConstants
Public myTool As Form
Public myMinimap As frmMinimap

Public curX As Single, curY As Single 'mouse cords on map

Private myIml As New cVBALImageList
Private myCurUnit As UnitEntry

Public enableGrid As Boolean

'Massediting
Public massMode As Boolean, massStep As Integer
Private massx1 As Long, massy1 As Long, massx2 As Long, massy2 As Long

'Consecutive paste (from the tileset window)
Public consecStep As Integer
Private consecx1 As Long, consecy1 As Long, consecx2 As Long, consecy2 As Long

'debounce edits
Private oldX As Long, oldY As Long

'Copy stuff
Private copyStep As Integer, pasteActive As Boolean
Private copyx1 As Long, copyy1 As Long, copyx2 As Long, copyy2 As Long

Private unitRecs() As UnitRec
Public numUnitRecs As Long

Private myTabId As Long

'Undo stuff
Private undoHist() As UndoRec 'array of undo histories
Private numUndos As Long

Implements ISubclass

'Creation / saving of maps

Public Sub CreateMap(ByVal w As Long, ByVal h As Long, Optional ByVal makeDefaultTilesets As Long = 1)
On Error GoTo oops

Dim i As Long, j As Long

Set myMap = ResMan.CreateNewMap(w, h)
myMap.TileSetManager.AddTileSet defTilesets(0)

'Do they want the default tilesets and terrains?
If (makeDefaultTilesets = 1) And (numTilesets > 1) Then
    For i = 1 To numTilesets - 1
        myMap.TileSetManager.AddTileSet defTilesets(i)
        myMap.TileSetManager.MapInTiles i, 0, myMap.TileSetManager.TileSet(i).NumTiles
    Next
    If numTerrains > 0 Then 'Create default terrains
        myMap.TileSetManager.SetNumTerrains numTerrains
        For i = 0 To numTerrains - 1
            myMap.TileSetManager.TerrainStartTile(i) = defTerrains(i).startTile
            myMap.TileSetManager.TerrainEndTile(i) = defTerrains(i).endTile
            myMap.TileSetManager.TerrainDozed(i) = defTerrains(i).dozed
            myMap.TileSetManager.TerrainRubble(i) = defTerrains(i).rubble
            For j = 0 To 5
                myMap.TileSetManager.TerrainTubeUnk(i, j) = defTerrains(i).tubeUnk(j)
            Next
            For j = 0 To 15
                myMap.TileSetManager.TerrainLavaWall(i, j) = defTerrains(i).lavaWall(j)
            Next
            For j = 0 To 15
                myMap.TileSetManager.TerrainMicrobeWall(i, j) = defTerrains(i).microbeWall(j)
            Next
            For j = 0 To 15
                myMap.TileSetManager.TerrainNormalWall(i, j) = defTerrains(i).normalWall(j)
            Next
            For j = 0 To 15
                myMap.TileSetManager.TerrainDamagedWall(i, j) = defTerrains(i).damagedWall(j)
            Next
            For j = 0 To 15
                myMap.TileSetManager.TerrainRuinedWall(i, j) = defTerrains(i).ruinedWall(j)
            Next
            myMap.TileSetManager.TerrainLava(i) = defTerrains(i).lava
            myMap.TileSetManager.TerrainFlat1(i) = defTerrains(i).flat1
            myMap.TileSetManager.TerrainFlat2(i) = defTerrains(i).flat2
            myMap.TileSetManager.TerrainFlat3(i) = defTerrains(i).flat3
            For j = 0 To 15
                myMap.TileSetManager.TerrainTube(i, j) = defTerrains(i).tube(j)
            Next
            myMap.TileSetManager.TerrainScorched(i) = defTerrains(i).scorched
            For j = 0 To 20
                myMap.TileSetManager.TerrainUnknown(i, j) = defTerrains(i).unkTile(j)
            Next
        Next
    End If
End If
'Init the map with blank values
Dim id As Long
id = myMap.TileSetManager.MapInTiles(0, 0, 1)
For i = 0 To w - 1
    For j = 0 To h - 1
        myMap.TileData(i, j) = 0
        myMap.mappingIndex(i, j) = id
        'myMap.CellType(i, j) = FastPassable1
        'myMap.Expand(i, j) = 0
        'myMap.lava(i, j) = 0
        'myMap.LavaPossible(i, j) = 0
        'myMap.Microbe(i, j) = 0
        'myMap.UnitIndex(i, j) = 0
        'myMap.WallOrBuilding(i, j) = 0
    Next
Next

'Set the bottom row to all impassable 1
For i = 0 To myMap.TileWidth - 1
    myMap.CellType(i, myMap.TileHeight - 1) = Impassable1
Next

'Set the current tool to the tile picker
ChangeTool PlaceTile
'ChangeTool CellTypeEdit

'Load the minimap
Set myMinimap = New frmMinimap
'Set myMinimap.ownerForm = Me
myMinimap.Show
myMinimap.SetNewMap Me

'Me.ZOrder 0
Exit Sub
oops:
GenerateError "Map creation failed. Perhaps the default tilesets do not exist?", "MapManager::CreateMap"
End Sub

Public Function LoadMap(ByVal Filename As String) As Boolean
On Error GoTo oops

Set myMap = ResMan.LoadMapFile(Filename, TileGroups)

'Load the unit definitions
If Dir$(Filename & ".dat") <> "" Then
    Dim f As Integer, i As Long, j As Long
    f = FreeFile
    Open Filename & ".dat" For Binary As #f
    Get #f, , numUnitRecs
    For i = 0 To numUnitRecs - 1
        ReDim Preserve unitRecs(numUnitRecs - 1) As UnitRec
        Get #f, , unitRecs(i).locX
        Get #f, , unitRecs(i).locY
        Get #f, , unitRecs(i).playerNum
        Get #f, , unitRecs(i).uType.mapID
        Get #f, , unitRecs(i).wType.mapID
        Get #f, , unitRecs(i).uType.isGaia
        If unitRecs(i).uType.isGaia = True Then
            'Load the gaia extra data
            Get #f, , unitRecs(i).uType.extra1
            Get #f, , unitRecs(i).uType.extra2
            Get #f, , unitRecs(i).uType.extra3
        End If
        'Scan for these values
        If allUnitDefs(j).isGaia = True Then
            'Must match it based on the extra data values as well
            For j = 0 To numUnitDefs
                If allUnitDefs(j).mapID = unitRecs(i).uType.mapID And allUnitDefs(j).extra1 = unitRecs(i).uType.extra1 And allUnitDefs(j).extra2 = unitRecs(i).uType.extra2 And allUnitDefs(j).extra3 = unitRecs(i).uType.extra3 Then
                    unitRecs(i).uType = allUnitDefs(j)
                    GoTo scanWeapId
                End If
            Next
        Else
            For j = 0 To numUnitDefs
                If allUnitDefs(j).mapID = unitRecs(i).uType.mapID Then
                    unitRecs(i).uType = allUnitDefs(j)
                    GoTo scanWeapId
                End If
            Next
        End If
        GenerateError "Map ID " & CStr(unitRecs(i).uType.mapID) & " not found", "MapManager::LoadMap"
scanWeapId:
        'Scan for weapon id
        If unitRecs(i).wType.mapID = 0 Then GoTo endScan
        For j = 0 To numWeaponDefs
            If allWeaponDefs(j).mapID = unitRecs(i).wType.mapID Then
                unitRecs(i).wType = allWeaponDefs(j)
                GoTo endScan
            End If
        Next
        GenerateError "Map ID " & CStr(unitRecs(i).uType.mapID) & " not found", "MapManager::LoadMap"
endScan:
    Next
    Close #f
End If

'Set the current tool to the tile picker
ChangeTool PlaceTile
'ChangeTool CellTypeEdit

'Load the minimap
Set myMinimap = New frmMinimap
'Set myMinimap.ownerForm = Me
myMinimap.Show
myMinimap.SetExtents Int(picMap.ScaleWidth / 32), Int(picMap.ScaleHeight / 32)
myMinimap.SetNewMap Me

Me.ZOrder 0

isFile = True
LoadMap = True
Exit Function
oops:
GenerateError "Map file failed to load. This is not a mapper bug, however the map file could be corrupt.", "MapManager::LoadMap"
LoadMap = False
End Function

Public Function SaveSelf() As Boolean
If isFile = False Then
    'Save as for this file
    If SaveSelfAs = True Then isFile = True Else Exit Function
    SetNewName myName, myTitle
    SaveSelf = True
Else
    'Save the active file
    Dim stream As StreamWriter
    Set stream = ResMan.OpenStreamWrite(myName)
    myMap.SaveMap stream, TileGroups
    Set stream = Nothing
    'Save the unit definitions
    If numUnitRecs > 0 Then
        Dim f As Integer, i As Long
        f = FreeFile
        If Dir$(myName & ".dat") <> "" Then Kill myName & ".dat"
        Open myName & ".dat" For Binary As #f
        Put #f, , numUnitRecs
        For i = 0 To numUnitRecs - 1
            Put #f, , unitRecs(i).locX
            Put #f, , unitRecs(i).locY
            Put #f, , unitRecs(i).playerNum
            Put #f, , unitRecs(i).uType.mapID
            Put #f, , unitRecs(i).wType.mapID
            Put #f, , unitRecs(i).uType.isGaia
            If unitRecs(i).uType.isGaia = True Then
                'Save the gaia extra data
                Put #f, , unitRecs(i).uType.extra1
                Put #f, , unitRecs(i).uType.extra2
                Put #f, , unitRecs(i).uType.extra3
            End If
        Next
        Close #f
    End If
    SaveSelf = True
End If
End Function

Public Function SaveSelfAs() As Boolean
'Save the map file
Dim sFile As String
sFile = myName
If fMainForm.cCMDlg.VBGetSaveFileName(sFile, myTitle, True, "OP2 Map Files (*.map)|*.map", , mapsDir, "Save Map File", "map") = False Then Exit Function
myName = sFile
Dim stream As StreamWriter
Set stream = ResMan.OpenStreamWrite(sFile)
myMap.SaveMap stream, TileGroups
Set stream = Nothing

'Save the unit definitions
If numUnitRecs > 0 Then
    Dim f As Integer, i As Long
    f = FreeFile
    If Dir$(sFile & ".dat") <> "" Then Kill sFile & ".dat"
    Open sFile & ".dat" For Binary As #f
    Put #f, , numUnitRecs
    For i = 0 To numUnitRecs - 1
        Put #f, , unitRecs(i).locX
        Put #f, , unitRecs(i).locY
        Put #f, , unitRecs(i).playerNum
        Put #f, , unitRecs(i).uType.mapID
        Put #f, , unitRecs(i).wType.mapID
        Put #f, , unitRecs(i).uType.isGaia
        If unitRecs(i).uType.isGaia = True Then
            'Save the gaia extra data
            Put #f, , unitRecs(i).uType.extra1
            Put #f, , unitRecs(i).uType.extra2
            Put #f, , unitRecs(i).uType.extra3
        End If
    Next
    Close #f
End If

SaveSelfAs = True
SetNewName myName, myTitle
End Function

Private Sub Form_Activate()
If enableKbPoll Then tmrScroll.Enabled = True
'Select this tab
fMainForm.DTabControl.Tabs.Item("MAP" & CStr(myTabId)).Selected = True
End Sub

Private Sub Form_Deactivate()
tmrScroll.Enabled = False
End Sub

Private Sub Form_Load()
'Set up the imagelist
myIml.OwnerHDC = fMainForm.picStatus.hDC
myIml.ColourDepth = myIml.SystemColourDepth
myIml.IconSizeX = 16
myIml.IconSizeY = 16
myIml.Create
Dim hImg As Picture
Set hImg = LoadResPicture(102, vbResBitmap)
myIml.AddFromHandle hImg.handle, IMAGE_BITMAP, , &HFF00FF
Set hImg = Nothing

'Set up the toolbar
myToolbar.ImageSource = CTBExternalImageList
myToolbar.SetImageList myIml.hIml
myToolbar.CreateToolbar 16, , , True
myToolbar.AddButton "Place Tile", 0, , , , CTBCheckGroup, "placetile"
myToolbar.AddButton "Place Tile Group", 1, , , , CTBCheckGroup, "placegroup"
myToolbar.AddButton "Edit Cell Types", 2, , , , CTBCheckGroup, "editcelltype"
myToolbar.AddButton "Place Object", 3, , , , CTBCheckGroup, "placeunit"
'myToolbar.AddButton "Tile Info", 4, , , , CTBCheckGroup, "tileinfo"
myToolbar.AddButton , , , , , CTBSeparator
myToolbar.AddButton "Tileset Editor", 5, , , , CTBNormal, "editmapping"
myToolbar.AddButton "Terrain Editor", 6, , , , CTBNormal, "editterrain"
myToolbar.AddButton , , , , , CTBSeparator
myToolbar.AddButton "Mass Edit Mode", 7, , , , CTBCheck, "massmode"
myToolbar.AddButton "Toggle Grid", 8, , , , CTBCheck, "grid"
myToolbar.ButtonChecked("placetile") = True

'Link the toolbar to the rebar
myRebar.CreateRebar Me.hWnd
myRebar.AddBandByHwnd myToolbar.hWnd, , False, , "ToolsBar"

'Add this form to the tab bar
fMainForm.DTabControl.Tabs.Add "MAP" & CStr(fMainForm.DTabControl.Tabs.Count), , Me.Caption
myTabId = fMainForm.DTabControl.Tabs.Count - 1
Me.tag = "MAP" & CStr(myTabId)

'Stop the scrollbars from taking keydown notifications
AttachMessage Me, hsbScroll.hWnd, WM_KEYDOWN
AttachMessage Me, vsbScroll.hWnd, WM_KEYDOWN
'Also subclass the MDI activate message so we can know what window is activating
'AttachMessage Me, Me.hwnd, WM_MDIACTIVATE

'Start with default grid mode
enableGrid = enableGridDef
myToolbar.ButtonChecked("grid") = enableGrid

'Form_Resize
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
Select Case MsgBox("Do you want to save changes to " & myName & "?", vbQuestion Or vbYesNoCancel, "Save Changes?")
    Case vbYes
        If SaveSelf = False Then cancel = True: Exit Sub
    Case vbCancel
        cancel = True: Exit Sub
End Select
myRebar.RemoveAllRebarBands

'Unsubclass the scrollbars and MDI message
DetachMessage Me, hsbScroll.hWnd, WM_KEYDOWN
DetachMessage Me, vsbScroll.hWnd, WM_KEYDOWN
'DetachMessage Me, Me.hwnd, WM_MDIACTIVATE
End Sub

Private Sub Form_Resize()
On Error Resume Next
picMap.Move 0, myRebar.RebarHeight, Me.ScaleWidth - vsbScroll.Width, Me.ScaleHeight - hsbScroll.Height - myRebar.RebarHeight
vsbScroll.TOp = myRebar.RebarHeight
vsbScroll.Left = picMap.Width
vsbScroll.Height = picMap.Height
hsbScroll.TOp = picMap.Height + myRebar.RebarHeight
hsbScroll.Width = picMap.Width
'Resize the scrollers accordingly
hsbScroll.Max = myMap.TileWidth - Int(picMap.ScaleWidth / 32) ' - 1
vsbScroll.Max = myMap.TileHeight - Int(picMap.ScaleHeight / 32) ' - 1
'Resize the rebar
myRebar.RebarSize
RedrawSelf
'Resize the minimap box
myMinimap.SetExtents Int(picMap.ScaleWidth / 32), Int(picMap.ScaleHeight / 32)
myMinimap.SetLoc hsbScroll.Value, vsbScroll.Value
End Sub

Private Sub Form_Unload(cancel As Integer)
Set myMap = Nothing
If myTool Is Nothing Then Exit Sub
myTool.prepareClose = True
Unload myTool
Set myTool = Nothing

If myMinimap Is Nothing Then Exit Sub
Unload myMinimap
Set myMinimap = Nothing

'Remove this from the tab bar
fMainForm.DTabControl.Tabs.Remove "MAP" & CStr(myTabId)
End Sub

Private Sub hsbScroll_Change()
myMinimap.SetLoc hsbScroll.Value, vsbScroll.Value
RedrawSelf
End Sub

Private Sub hsbScroll_Scroll()
hsbScroll_Change
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
'Nothing here, it refuses to compile without it though
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
'If CurrentMessage = WM_MDIACTIVATE Then
'    ISubclass_MsgResponse = emrPreprocess 'we process it before VB does
'Else 'yum yum, for the scroll bar stuff
    ISubclass_MsgResponse = emrConsume 'Yep, eat it up! =:-)
'End If
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
''handle tools activation / deactivation if this is an MDI activate message
'If iMsg = WM_MDIACTIVATE Then
'    If wParam = Me.hwnd Then
'        'If we're deactivating because our tool windows are activating, then stop
'        If lParam = myTool.hwnd Or lParam = myMinimap.hwnd Then Exit Function
'        'Safe to hide tool windows
'        myTool.Hide
'        myMinimap.Hide
'    End If
'End If
End Function

Private Sub myToolbar_ButtonClick(ByVal lButton As Long)
'handle button clicks and activate the proper tool
Select Case myToolbar.ButtonKey(lButton)
    Case "placetile"
        ChangeTool PlaceTile
    Case "placegroup"
        ChangeTool PlaceTileGroup
    Case "editcelltype"
        ChangeTool CellTypeEdit
    Case "placeunit"
        ChangeTool PlaceUnit
    Case "tileinfo"
        ChangeTool InfoView
    Case "editmapping"
        'Pop up the Edit Mappings dialog
        frmEditMappings.Hide
        SetStatusBar "Loading..."
        frmEditMappings.SetNewMgr myMap.TileSetManager
        frmEditMappings.SetNewName myName
        SetStatusBar "Ready"
        frmEditMappings.Show 1, fMainForm
        'Refresh the tools & map
        Select Case toolMode
            Case PlaceTile
                myTool.SetNewMgr myMap.TileSetManager
            Case PlaceTileGroup
                myTool.SetNewMap myMap
            Case PlaceUnit
            Case CellTypeEdit
            Case InfoView
                '**TODO** this needs a form too
        End Select
        RedrawSelf
    Case "editterrain"
        'Pop up the Edit Terrains dialog
        frmEditTerrains.Hide
        SetStatusBar "Loading..."
        frmEditTerrains.SetNewMgr myMap.TileSetManager
        frmEditTerrains.SetNewName myName
        SetStatusBar "Ready"
        frmEditTerrains.Show 1, fMainForm
        'Refresh the tools & map
        Select Case toolMode
            Case PlaceTile
                myTool.SetNewMgr myMap.TileSetManager
            Case PlaceTileGroup
                myTool.SetNewMap myMap
            Case PlaceUnit
            Case CellTypeEdit
            Case InfoView
                '**TODO** this needs a form too
        End Select
        RedrawSelf
    Case "massmode"
        'Turn on or off the Mass Edit Mode
        massMode = Not massMode
        If massMode = False Then
            SetStatusBar "Ready"
            massStep = 0
        Else
            'Check to see if the mode is valid
            If toolMode = PlaceTile Or toolMode = CellTypeEdit Then
                massStep = 1
                SetStatusBar "Click upper left corner of area to mass edit"
            Else
                massMode = False
                massStep = 0
                myToolbar.ButtonChecked("massmode") = False
                Beep
                SetStatusBar "Mass mode is not supported for this tool."
            End If
        End If
    Case "grid"
        enableGrid = Not enableGrid
        RedrawSelf
End Select
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If ((Int(x / 32) + hsbScroll.Value) < 0) Or ((Int(x / 32) + hsbScroll.Value) > myMap.TileWidth - 1) Then Exit Sub
If ((Int(y / 32) + vsbScroll.Value) < 0) Or ((Int(y / 32) + vsbScroll.Value) > myMap.TileHeight - 1) Then Exit Sub

curX = x: curY = y

'Debug.Print "MouseDown Event. Button = " & CStr(Button = vbLeftButton)

HandleMouseCmds (Button = vbLeftButton), x, y

If Button = vbRightButton Then
    Me.SetFocus
    fMainForm.mnuMapEnableGrid.Checked = enableGrid
    Me.PopupMenu fMainForm.mnuMap
End If

'Show X-Y cords on statusbar
SetStatusXY Int(x / 32) + hsbScroll.Value + 1, Int(y / 32) + vsbScroll.Value + 1

'curProject.editMap.SetTile Int(X / 32) + hsbScroll.Value, Int(Y / 32) + vsbScroll.Value, mainTileset.GetCurTile
RedrawSelf
Exit Sub
oops:
GenerateError "Map window event handling failed", "MapManager::Map.MouseDown"
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo oops
If ((Int(x / 32) + hsbScroll.Value) < 0) Or ((Int(x / 32) + hsbScroll.Value) > myMap.TileWidth - 1) Then Exit Sub
If ((Int(y / 32) + vsbScroll.Value) < 0) Or ((Int(y / 32) + vsbScroll.Value) > myMap.TileHeight - 1) Then Exit Sub

curX = x: curY = y

'Debug.Print "MouseMove Event. Button = " & CStr(Button = vbLeftButton)

HandleMouseCmds (Button = vbLeftButton), x, y

'Show X-Y cords on statusbar
SetStatusXY Int(x / 32) + hsbScroll.Value + 1, Int(y / 32) + vsbScroll.Value + 1

'curProject.editMap.SetTile Int(X / 32) + hsbScroll.Value, Int(Y / 32) + vsbScroll.Value, mainTileset.GetCurTile
RedrawSelf
Exit Sub
oops:
GenerateError "Map window event handling failed", "MapManager::Map.MouseMove"
End Sub

Private Sub picMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
oldX = -1
oldY = -1
End Sub

Private Sub tmrScroll_Timer()
On Error Resume Next
If GetKeyState(VK_RIGHT) And KEY_PRESSED Then hsbScroll.Value = hsbScroll.Value + 1
If GetKeyState(VK_LEFT) And KEY_PRESSED Then hsbScroll.Value = hsbScroll.Value - 1
If GetKeyState(VK_DOWN) And KEY_PRESSED Then vsbScroll.Value = vsbScroll.Value + 1
If GetKeyState(VK_UP) And KEY_PRESSED Then vsbScroll.Value = vsbScroll.Value - 1
End Sub

Private Sub vsbScroll_Change()
myMinimap.SetLoc hsbScroll.Value, vsbScroll.Value
RedrawSelf
End Sub

Private Sub vsbScroll_Scroll()
vsbScroll_Change
End Sub

Public Sub RedrawSelf()
On Error GoTo oops
'Redraw the map
Dim i As Long
picMap.Cls
myMap.Draw Me.picMap.hDC, hsbScroll.Value * 32, vsbScroll.Value * 32, picMap.ScaleWidth, picMap.ScaleHeight
'Handle other commands
If toolMode = CellTypeEdit Then
    'Go thru each visible tile, then AND the bitmaps for each celltype
    Dim j As Long
    For i = hsbScroll.Value To hsbScroll.Value + Int(picMap.ScaleWidth / 32)
        For j = vsbScroll.Value To vsbScroll.Value + Int(picMap.ScaleHeight / 32)
            'Prevent the user from going outside a valid region and crashing the editor
            If (i < myMap.TileWidth) And (j < myMap.TileHeight) Then
                Dim sName As String, lColor As Long, R As RECT
                SelectOverlay i, j, sName, lColor
                R.Left = (i - hsbScroll.Value) * 32: R.TOp = (j - vsbScroll.Value) * 32
                R.Right = (i - hsbScroll.Value + 1) * 32: R.Bottom = (j - vsbScroll.Value + 1) * 32
                picMap.ForeColor = lColor
                'BitBlt picMap.hDC, (i - hsbScroll.Value) * 32, (j - vsbScroll.Value) * 32, 32, 32, fMainForm.picCellType.hDC, SrcX, SrcY, vbSrcAnd
                DrawText picMap.hDC, sName, Len(sName), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            End If
        Next
    Next
'Draw units if unit placer is activated
ElseIf toolMode = PlaceUnit Then
    'Scan thru all units on the map
    For i = 0 To numUnitRecs - 1
        'Check that each unit is inside the viewable area and that it has an ART ID
        If unitRecs(i).locX >= hsbScroll.Value And unitRecs(i).locX <= hsbScroll.Value + picMap.ScaleWidth \ 32 And unitRecs(i).uType.artID <> -1 Then
            If unitRecs(i).locY >= vsbScroll.Value And unitRecs(i).locY <= vsbScroll.Value + picMap.ScaleHeight \ 32 Then
                'It's inside, so draw it
                If unitRecs(i).uType.hasTubes = True Then
                    '**NOTE** to self: remember that when using tubes to calculate the building rectangle
                    'the ver tube is relative to the X position NOT the Y, and vice versa
                    PRTFile.PasteAnimTrans picMap.hDC, _
                        (unitRecs(i).locX - unitRecs(i).uType.verTubeLoc + 1 - hsbScroll.Value) * 32, _
                        (unitRecs(i).locY - unitRecs(i).uType.horTubeLoc + 1 - vsbScroll.Value) * 32, _
                        (unitRecs(i).locX - unitRecs(i).uType.verTubeLoc + 1 + unitRecs(i).uType.sizeX - hsbScroll.Value) * 32, _
                        (unitRecs(i).locY - unitRecs(i).uType.horTubeLoc + 1 + unitRecs(i).uType.sizeY - vsbScroll.Value) * 32, _
                        unitRecs(i).uType.artID, unitRecs(i).playerNum
                    If unitRecs(i).uType.canHaveTurret = True And unitRecs(i).wType.artID <> -1 Then
                        'Have to draw weapon art
                        PRTFile.PasteAnimTrans picMap.hDC, _
                            (unitRecs(i).locX - unitRecs(i).uType.verTubeLoc + 1 - hsbScroll.Value) * 32, _
                            (unitRecs(i).locY - unitRecs(i).uType.horTubeLoc + 1 - vsbScroll.Value) * 32, _
                            (unitRecs(i).locX - unitRecs(i).uType.verTubeLoc + 1 + unitRecs(i).uType.sizeX - hsbScroll.Value) * 32, _
                            (unitRecs(i).locY - unitRecs(i).uType.horTubeLoc + 1 + unitRecs(i).uType.sizeY - vsbScroll.Value) * 32, _
                            unitRecs(i).wType.artID, unitRecs(i).playerNum
                    End If
                    'Draw the tubes as white lines to make lining things up a bit easier
                    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
                    
                    Dim h1 As POINTAPI, h2 As POINTAPI, v1 As POINTAPI, v2 As POINTAPI, upper As POINTAPI
                    upper.x = (unitRecs(i).locX - hsbScroll.Value - unitRecs(i).uType.verTubeLoc)
                    upper.y = (unitRecs(i).locY - vsbScroll.Value - unitRecs(i).uType.horTubeLoc)
                    
                    h1.x = (upper.x + unitRecs(i).uType.sizeX + 1) * 32
                    h2.x = h1.x + 32
                    h1.y = ((upper.y + unitRecs(i).uType.horTubeLoc) * 32) + 16
                    h2.y = h1.y
                    v1.x = ((upper.x + unitRecs(i).uType.verTubeLoc) * 32) + 16
                    v2.x = v1.x
                    v1.y = (upper.y + unitRecs(i).uType.sizeY + 1) * 32
                    v2.y = v1.y + 32
                    
                    picMap.Line (h1.x, h1.y)-(h2.x, h2.y), vbWhite
                    picMap.Line (v1.x, v1.y)-(v2.x, v2.y), vbWhite
                    
                ElseIf unitRecs(i).uType.isStructure = False Then
                    'It's a vehicle (or some other 1 cell object), very simple blit at the current position
                    
                    'See if it uses one of the special Art IDs
                    If unitRecs(i).uType.artID < 0 Then
                        Dim nam As String, col As Long, rct As RECT
                        Select Case unitRecs(i).uType.artID
                            Case -2 'White T
                                nam = "T": col = vbWhite
                            Case -3 'Cyan NW
                                nam = "NW": col = vbCyan
                            Case -4 'Cyan LW
                                nam = "LW": col = vbCyan
                            Case -5 'Cyan MW
                                nam = "MW": col = vbCyan
                        End Select
                        rct.Left = (unitRecs(i).locX - hsbScroll.Value) * 32: rct.TOp = (unitRecs(i).locY - vsbScroll.Value) * 32
                        rct.Right = (unitRecs(i).locX - hsbScroll.Value + 1) * 32: rct.Bottom = (unitRecs(i).locY - vsbScroll.Value + 1) * 32
                        picMap.ForeColor = col
                        DrawText picMap.hDC, nam, Len(nam), rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
                    Else
                        PRTFile.PasteAnimTrans picMap.hDC, _
                            (unitRecs(i).locX - hsbScroll.Value) * 32, _
                            (unitRecs(i).locY - vsbScroll.Value) * 32, _
                            (unitRecs(i).locX + 1 - hsbScroll.Value) * 32, _
                            (unitRecs(i).locY + 1 - vsbScroll.Value) * 32, _
                            unitRecs(i).uType.artID, unitRecs(i).playerNum
                        If unitRecs(i).uType.canHaveTurret = True And unitRecs(i).wType.artID <> -1 Then
                            'Have to draw weapon art
                            PRTFile.PasteAnimTrans picMap.hDC, _
                                (unitRecs(i).locX - hsbScroll.Value) * 32, _
                                (unitRecs(i).locY - vsbScroll.Value) * 32, _
                                (unitRecs(i).locX + 1 - hsbScroll.Value) * 32, _
                                (unitRecs(i).locY + 1 - vsbScroll.Value) * 32, _
                                unitRecs(i).wType.artID, unitRecs(i).playerNum
                        End If
                    End If
                Else
                    'It doesn't have tubes, so calc the rect without the tube positions
                    PRTFile.PasteAnimTrans picMap.hDC, _
                        (unitRecs(i).locX - hsbScroll.Value) * 32, _
                        (unitRecs(i).locY - vsbScroll.Value) * 32, _
                        (unitRecs(i).locX + unitRecs(i).uType.sizeX - hsbScroll.Value) * 32, _
                        (unitRecs(i).locY + unitRecs(i).uType.sizeY - vsbScroll.Value) * 32, _
                        unitRecs(i).uType.artID, unitRecs(i).playerNum
                    If unitRecs(i).uType.canHaveTurret = True And unitRecs(i).wType.artID <> -1 Then
                        'Have to draw weapon art
                        PRTFile.PasteAnimTrans picMap.hDC, _
                            (unitRecs(i).locX - hsbScroll.Value) * 32, _
                            (unitRecs(i).locY - vsbScroll.Value) * 32, _
                            (unitRecs(i).locX + unitRecs(i).uType.sizeX - hsbScroll.Value) * 32, _
                            (unitRecs(i).locY + unitRecs(i).uType.sizeY - vsbScroll.Value) * 32, _
                            unitRecs(i).wType.artID, unitRecs(i).playerNum
                    End If
                End If
            End If
        End If
    Next
End If
'Add gridlines
If enableGrid Then
    For i = 1 To Int(picMap.Width / 32)
        picMap.Line (i * 32, 0)-(i * 32, picMap.Height), vbBlack
    Next
    For i = 1 To Int(picMap.Height / 32)
        picMap.Line (0, i * 32)-(picMap.Width, i * 32), vbBlack
    Next
End If
picMap.Refresh
Exit Sub
oops:
GenerateError "Map redrawing failed", "MapManager::RedrawSelf"
End Sub

Private Sub SelectOverlay(ByVal SrcX As Long, ByVal SrcY As Long, strName As String, lngColor As Long)
Select Case myMap.CellType(SrcX, SrcY)
    Case FastPassable1
        strName = "F1": lngColor = vbGreen
        'SrcX = 0: SrcY = 0
    Case FastPassable2
        strName = "F2": lngColor = vbGreen
        'SrcX = 0: SrcY = 32
    Case MediumPassable1
        strName = "M1": lngColor = vbGreen
        'SrcX = 0: SrcY = 64
    Case MediumPassable2
        strName = "M2": lngColor = vbGreen
        'SrcX = 0: SrcY = 96
    Case SlowPassable1
        strName = "S1": lngColor = vbGreen
        'SrcX = 0: SrcY = 128
    Case SlowPassable2
        strName = "S2": lngColor = vbGreen
        'SrcX = 0: SrcY = 160
    Case Impassable1
        strName = "I1": lngColor = vbRed
        'SrcX = 32: SrcY = 0
    Case Impassable2
        strName = "I2": lngColor = vbRed
        'SrcX = 32: SrcY = 32
    Case NorthCliffs
        strName = "NC": lngColor = vbRed
        'SrcX = 32: SrcY = 64
    Case CliffsHighSide
        strName = "CHS": lngColor = vbRed
        'SrcX = 32: SrcY = 96
    Case CliffsLowSide
        strName = "CLS": lngColor = vbRed
        'SrcX = 32: SrcY = 128
    Case zPad12
        strName = "Z12": lngColor = vbBlue
        'SrcX = 64: SrcY = 0
    Case zPad13
        strName = "Z13": lngColor = vbBlue
        'SrcX = 64: SrcY = 32
    Case zPad14
        strName = "Z14": lngColor = vbBlue
        'SrcX = 64: SrcY = 64
    Case zPad15
        strName = "Z15": lngColor = vbBlue
        'SrcX = 64: SrcY = 96
    Case zPad16
        strName = "Z16": lngColor = vbBlue
        'SrcX = 64: SrcY = 128
    Case zPad17
        strName = "Z17": lngColor = vbBlue
        'SrcX = 64: SrcY = 160
    Case zPad18
        strName = "Z18": lngColor = vbBlue
        'SrcX = 64: SrcY = 192
    Case zPad19
        strName = "Z19": lngColor = vbBlue
        'SrcX = 64: SrcY = 224
    Case zPad20
        strName = "Z20": lngColor = vbBlue
        'SrcX = 64: SrcY = 256
    Case DozedArea
        strName = "D": lngColor = vbYellow
        'SrcX = 96: SrcY = 0
    Case rubble
        strName = "R": lngColor = vbYellow
        'SrcX = 96: SrcY = 32
    Case VentsAndFumaroles
        strName = "F": lngColor = vbMagenta
        'SrcX = 128: SrcY = 0
    Case normalWall
        strName = "NW": lngColor = vbCyan
        'SrcX = 160: SrcY = 0
    Case microbeWall
        strName = "MW": lngColor = vbCyan
        'SrcX = 160: SrcY = 32
    Case lavaWall
        strName = "LW": lngColor = vbCyan
        'SrcX = 160: SrcY = 64
    Case Tube0
        strName = "T0": lngColor = vbWhite
        'SrcX = 192: SrcY = 0
    Case Tube1
        strName = "T1": lngColor = vbWhite
        'SrcX = 192: SrcY = 32
    Case Tube2
        strName = "T2": lngColor = vbWhite
        'SrcX = 192: SrcY = 64
    Case Tube3
        strName = "T3": lngColor = vbWhite
        'SrcX = 192: SrcY = 96
    Case Tube4
        strName = "T4": lngColor = vbWhite
        'SrcX = 192: SrcY = 128
    Case Tube5
        strName = "T5": lngColor = vbWhite
        'SrcX = 192: SrcY = 160
End Select
End Sub

Private Sub HandleMouseCmds(ByVal buttonDown As Boolean, ByVal x As Single, ByVal y As Single) 'handle mouse input onto the map
On Error GoTo oops
Dim i As Long, j As Long
If copyStep = 1 Then
1
shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
    If buttonDown Then
2
        'handle upper left copy click
        copyx1 = Int(x / 32) + hsbScroll.Value
        copyy1 = Int(y / 32) + vsbScroll.Value
        copyStep = 2
        SetStatusBar "Click lower right corner of selection area"
    End If
ElseIf copyStep = 2 Then
3
    'On Error Resume Next
    shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
    'shpCursor.Move (copyx1 - hsbScroll.value) * 32, (copyy1 - vsbScroll.value) * 32, (Int(X / 32) - copyx1 - hsbScroll.value + 1) * 32, (Int(Y / 32) - copyy1 - vsbScroll.value + 1) * 32
    If buttonDown Then
4
        'handle lower right copy click
        copyx2 = Int(x / 32) + hsbScroll.Value
        copyy2 = Int(y / 32) + vsbScroll.Value
        If copyx2 >= copyx1 And copyy2 >= copyy1 Then
5
            Set copyBuffer = myMap.Copy(copyx1, copyy1, copyx2, copyy2)
            'Copy cell types if wanted
            If copyPasteCellTypes Then
                ReDim cellTypeBuffer(copyBuffer.TileWidth * copyBuffer.TileHeight - 1) As CellTypes
                For i = 0 To copyBuffer.TileWidth - 1
                    For j = 0 To copyBuffer.TileHeight - 1
                        cellTypeBuffer(j * copyBuffer.TileWidth + i) = myMap.CellType(i + copyx1, j + copyy1)
                    Next j
                Next i
            End If
            
            copyStep = 0
            SetStatusBar "Selection copied to clipboard."
        Else
6
            copyStep = 0
            Beep
            SetStatusBar "Invalid selection area"
        End If
    End If
ElseIf pasteActive Then
7
    'Security. :P
    If copyBuffer Is Nothing Then Exit Sub
    shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, copyBuffer.TileWidth * 32, copyBuffer.TileHeight * 32
    If buttonDown Then
8
        'Save undo step
        ReDim Preserve undoHist(numUndos) As UndoRec
        undoHist(numUndos).typeChg = UndoGroup
        Set undoHist(numUndos).savedGroup = myMap.Copy(Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value, Int(x / 32) + hsbScroll.Value + copyBuffer.TileWidth - 1, Int(y / 32) + vsbScroll.Value + copyBuffer.TileHeight - 1)
        undoHist(numUndos).xPos = Int(x / 32) + hsbScroll.Value
        undoHist(numUndos).yPos = Int(y / 32) + vsbScroll.Value
        numUndos = numUndos + 1
        'Do operation
        myMap.Paste Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value, copyBuffer
        'Paste celltypes if selected
        If copyPasteCellTypes Then
            'Add cell types to undo history as well
            ReDim undoHist(numUndos - 1).savedMassCellType(copyBuffer.TileWidth * copyBuffer.TileHeight)
            For i = Int(x / 32) + hsbScroll.Value To Int(x / 32) + hsbScroll.Value + copyBuffer.TileWidth - 1
                For j = Int(y / 32) + vsbScroll.Value To Int(y / 32) + vsbScroll.Value + copyBuffer.TileHeight - 1
                    'Debug.Print "Attempting to paste celltype at " & i & "," & j & " => " & (i - (Int(x / 32) + hsbScroll.Value)) & "," & (j - (Int(y / 32) + vsbScroll.Value))
                    undoHist(numUndos - 1).savedMassCellType((j - (Int(y / 32) + vsbScroll.Value)) * copyBuffer.TileWidth + (i - (Int(x / 32) + hsbScroll.Value))) = myMap.CellType(i, j)
                    myMap.CellType(i, j) = cellTypeBuffer((j - (Int(y / 32) + vsbScroll.Value)) * copyBuffer.TileWidth + (i - (Int(x / 32) + hsbScroll.Value)))
                Next
            Next
        End If
        'Update minimap
        For i = Int(x / 32) + hsbScroll.Value To Int(x / 32) + hsbScroll.Value + copyBuffer.TileWidth - 1
            For j = Int(y / 32) + vsbScroll.Value To Int(y / 32) + vsbScroll.Value + copyBuffer.TileHeight - 1
9
                myMinimap.UpdatePixel i, j
            Next
        Next
        pasteActive = False
        SetStatusBar "Paste successful."
    End If
ElseIf massMode Then 'Mass mode handler
    Select Case massStep
        Case 1 'Select upper left
10
            shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
            If buttonDown Then
                'handle upper left mass click
                massx1 = Int(x / 32) + hsbScroll.Value
                massy1 = Int(y / 32) + vsbScroll.Value
                massStep = 2
                SetStatusBar "Click lower right corner of area to mass edit"
            End If
        Case 2 'Select lower right
11
            'On Error Resume Next
            shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
            'shpCursor.Move (copyx1 - hsbScroll.value) * 32, (copyy1 - vsbScroll.value) * 32, (Int(X / 32) - copyx1 - hsbScroll.value + 1) * 32, (Int(Y / 32) - copyy1 - vsbScroll.value + 1) * 32
            If buttonDown Then
                'handle lower right mass click
                massx2 = Int(x / 32) + hsbScroll.Value
                massy2 = Int(y / 32) + vsbScroll.Value
                If massx2 >= massx1 And massy2 >= massy1 Then
12
                    'Now, use supported tools to massedit
                    Select Case toolMode
                        Case PlaceTile
13
                            'Save undo step
                            ReDim Preserve undoHist(numUndos) As UndoRec
                            undoHist(numUndos).typeChg = UndoMassTile
                            'Add old values to array
                            ReDim undoHist(numUndos).savedMassMapping(massx2 - massx1, massy2 - massy1) As Long
                            For i = massx1 To massx2
                                For j = massy1 To massy2
                                    undoHist(numUndos).savedMassMapping(i - massx1, j - massy1) = myMap.mappingIndex(i, j)
                                Next
                            Next
                            undoHist(numUndos).xPos = massx1
                            undoHist(numUndos).yPos = massy1
                            undoHist(numUndos).x2Pos = massx2
                            undoHist(numUndos).y2Pos = massy2
                            numUndos = numUndos + 1
                            'Do operation & update minimap
                            For i = massx1 To massx2
                                For j = massy1 To massy2
                                    myMap.mappingIndex(i, j) = myTool.curMappingIndex
                                    myMinimap.UpdatePixel i, j
                                Next
                            Next
                        Case CellTypeEdit
14
                            'Make sure that the user isn't trying to edit the bottommost row of the cell types
                            If massy2 = myMap.TileHeight - 1 Then
                                MsgBox "You cannot modify the cell type for the bottom row of the map. Due to an internal limitation in Outpost 2, the bottom row must be all Impassable 1 (I1) cell type.", vbExclamation, "Cannot Edit"
                                massStep = 1
                                SetStatusBar "Click upper left corner of area to mass edit"
                                GoTo JustDebounce
                            End If
                            
                            'Save undo step
                            ReDim Preserve undoHist(numUndos) As UndoRec
                            undoHist(numUndos).typeChg = UndoMassCellType
                            'Add old values to array
                            ReDim undoHist(numUndos).savedMassCellType(massx2 - massx1, massy2 - massy1) As CellTypes
                            For i = massx1 To massx2
                                For j = massy1 To massy2
                                    undoHist(numUndos).savedMassCellType(i - massx1, j - massy1) = myMap.CellType(i, j)
                                Next
                            Next
                            undoHist(numUndos).xPos = massx1
                            undoHist(numUndos).yPos = massy1
                            undoHist(numUndos).x2Pos = massx2
                            undoHist(numUndos).y2Pos = massy2
                            numUndos = numUndos + 1
                            'Do operation & update minimap
                            For i = massx1 To massx2
                                For j = massy1 To massy2
                                    myMap.CellType(i, j) = myTool.currentSel
                                    myMinimap.UpdatePixel i, j
                                Next
                            Next
                        Case Else
15
                            Beep
                            SetStatusBar "This tool does not support mass editing"
                            GoTo JustDebounce
                    End Select
                    'Set copyBuffer = myMap.Copy(copyx1, copyy1, copyx2, copyy2)
                    massStep = 1
                    SetStatusBar "Mass edit ok. Click upper left corner of area to mass edit"
                Else
16
                    massStep = 1
                    Beep
                    SetStatusBar "Invalid selection area. Click upper left corner of area to mass edit"
                End If
            End If
        Case Else 'Checking, some logic error somewhere
            GenerateError "Mass mode is active but invalid step count", "MapManager::HandleMouseCmds"
    End Select
ElseIf consecStep > 0 Then 'Consecutive paste mode handler
17
    If toolMode <> PlaceTile Then
        'check for logic error
        GenerateError "Consecutive paste is active but PlaceTile mode isn't", "MapManager::HandleMouseCmds"
        consecStep = 0
        Exit Sub
    End If
    Select Case consecStep
        Case 1 'Pick upper left area
18
            shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
            If buttonDown Then
                'handle upper left area click
                consecx1 = Int(x / 32) + hsbScroll.Value
                consecy1 = Int(y / 32) + vsbScroll.Value
                consecStep = 2
                SetStatusBar "Click lower right corner of area to fill with tiles"
            End If
        Case 2
19
            'On Error Resume Next
            shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
            'shpCursor.Move (copyx1 - hsbScroll.value) * 32, (copyy1 - vsbScroll.value) * 32, (Int(X / 32) - copyx1 - hsbScroll.value + 1) * 32, (Int(Y / 32) - copyy1 - vsbScroll.value + 1) * 32
            'Debounce this operation, since it's stupid to copy 1x1 group this way anyway :-|
            If oldX = Int(x / 32) + hsbScroll.Value And oldY = Int(y / 32) + vsbScroll.Value And buttonDown Then Exit Sub
            If buttonDown Then
                'handle lower right area click
                consecx2 = Int(x / 32) + hsbScroll.Value
                consecy2 = Int(y / 32) + vsbScroll.Value
                If massx2 >= massx1 And massy2 >= massy1 Then
                    'Save undo step
                    ReDim Preserve undoHist(numUndos) As UndoRec
                    undoHist(numUndos).typeChg = UndoMassTile
                    'Add old values to array
                    ReDim undoHist(numUndos).savedMassMapping(consecx2 - consecx1, consecy2 - consecy1) As Long
                    For i = consecx1 To consecx2
                        For j = consecy1 To consecy2
                            undoHist(numUndos).savedMassMapping(i - consecx1, j - consecy1) = myMap.mappingIndex(i, j)
                        Next
                    Next
                    undoHist(numUndos).xPos = consecx1
                    undoHist(numUndos).yPos = consecy1
                    undoHist(numUndos).x2Pos = consecx2
                    undoHist(numUndos).y2Pos = consecy2
                    numUndos = numUndos + 1
                    'Do operation & update minimap
                    Dim whatMapping As Long
                    whatMapping = myMap.TileSetManager.GetMappingIndex(myTool.curTileset, myTool.curTileID, 0, 0)
                    If whatMapping = -1 Then 'Not a valid mapping index
                        Beep
                        SetStatusBar "Tileset #" & CStr(myTool.curTileset) & ", TileID 0x" & Hex(myTool.curTileID) & " does not have a mapping defined. Operation halted."
                        consecStep = 0
                        GoTo JustDebounce
                    End If
20
                    For j = consecy1 To consecy2
                        For i = consecx1 To consecx2 'Backwards ordering is important since we fill left to right first
                            'Apply tile to map
                            myMap.mappingIndex(i, j) = whatMapping
                            myMinimap.UpdatePixel i, j
                            If j = consecy2 And i = consecx2 Then 'If this is the last iteration, suppress the warning about no mapping
                                Exit For
                            End If
                            'Try to get the next mapping, if any
                            myTool.curTileID = myTool.curTileID + 1
                            whatMapping = myMap.TileSetManager.GetMappingIndex(myTool.curTileset, myTool.curTileID, 0, 0)
                            If whatMapping = -1 Then 'Not a valid mapping index
                                Beep
                                SetStatusBar "Tile 0x" & Hex(myTool.curTileID) & " does not have a mapping defined. Operation halted."
                                consecStep = 0
                                GoTo JustDebounce
                            End If
                        Next
                    Next
                    consecStep = 0
                    SetStatusBar "Tile fill completed"
                Else
21
                    consecStep = 0
                    Beep
                    SetStatusBar "Invalid selection area"
                End If
            End If
        Case Else 'Checking, some logic error somewhere
            GenerateError "Consecutive mode is active but invalid step count", "MapManager::HandleMouseCmds"
    End Select
Else 'Normal mode handler
22
    'Debounce the editing. Dragging across a tile creates hundreds of undo actions
    'so this prevents that from happening
    If oldX = Int(x / 32) + hsbScroll.Value And oldY = Int(y / 32) + vsbScroll.Value And buttonDown Then Exit Sub
    Select Case toolMode
        Case PlaceTile
23
            shpInnerUnit.Visible = False
            shpHTube.Visible = False: shpVTube.Visible = False
            shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
            If buttonDown Then
                'Save undo step
                ReDim Preserve undoHist(numUndos) As UndoRec
                undoHist(numUndos).typeChg = UndoTile
                undoHist(numUndos).savedMapping = myMap.mappingIndex(Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value)
                undoHist(numUndos).xPos = Int(x / 32) + hsbScroll.Value
                undoHist(numUndos).yPos = Int(y / 32) + vsbScroll.Value
                numUndos = numUndos + 1
                'Do operation
                myMap.mappingIndex(Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value) = myTool.curMappingIndex
                'Update minimap
                myMinimap.UpdatePixel Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value
            End If
        Case PlaceTileGroup
24
            shpInnerUnit.Visible = False
            shpHTube.Visible = False: shpVTube.Visible = False
            shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, myTool.curGroupW * 32, myTool.curGroupH * 32
            If buttonDown Then
                'Save undo step
                ReDim Preserve undoHist(numUndos) As UndoRec
                undoHist(numUndos).typeChg = UndoGroup
                Set undoHist(numUndos).savedGroup = myMap.Copy(Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value, Int(x / 32) + hsbScroll.Value + myTool.curGroupW - 1, Int(y / 32) + vsbScroll.Value + myTool.curGroupH - 1)
                undoHist(numUndos).xPos = Int(x / 32) + hsbScroll.Value
                undoHist(numUndos).yPos = Int(y / 32) + vsbScroll.Value
                numUndos = numUndos + 1
                'Do operation
                myMap.Paste Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value, myTool.curTileGroup
                'Update minimap
                For i = Int(x / 32) + hsbScroll.Value To Int(x / 32) + hsbScroll.Value + myTool.curGroupW - 1
                    For j = Int(y / 32) + vsbScroll.Value To Int(y / 32) + vsbScroll.Value + myTool.curGroupH - 1
                        myMinimap.UpdatePixel i, j
                    Next
                Next
            End If
        Case CellTypeEdit
25
            shpInnerUnit.Visible = False
            shpHTube.Visible = False: shpVTube.Visible = False
            shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
            If buttonDown Then
                'Make sure that the user isn't trying to edit the bottommost row of the cell types
                If Int(y / 32) + vsbScroll.Value = myMap.TileHeight - 1 Then
                    MsgBox "You cannot modify the cell type for the bottom row of the map. Due to an internal limitation in Outpost 2, the bottom row must be all Impassable 1 (I1) cell type.", vbExclamation, "Cannot Edit"
                    GoTo JustDebounce
                End If
                'Save undo step
                ReDim Preserve undoHist(numUndos) As UndoRec
                undoHist(numUndos).typeChg = UndoCellType
                undoHist(numUndos).savedCellType = myMap.CellType(Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value)
                undoHist(numUndos).xPos = Int(x / 32) + hsbScroll.Value
                undoHist(numUndos).yPos = Int(y / 32) + vsbScroll.Value
                numUndos = numUndos + 1
                'Do operation
                myMap.CellType(Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value) = myTool.currentSel
                'Update minimap
                myMinimap.UpdatePixel Int(x / 32) + hsbScroll.Value, Int(y / 32) + vsbScroll.Value
            End If
        Case PlaceUnit
26
            'tell tool to update data
            myTool.UpdateRec
            If curUnitSel.unitDef.isStructure = True Then
                'Is a structure
                If curUnitSel.unitDef.hasTubes = True Then
                    'Has tubes
                    shpCursor.Move (Int(x / 32) - curUnitSel.unitDef.verTubeLoc) * 32, (Int(y / 32) - curUnitSel.unitDef.horTubeLoc) * 32, (curUnitSel.unitDef.sizeX + 2) * 32, (curUnitSel.unitDef.sizeY + 2) * 32
                    shpInnerUnit.Move (Int(x / 32) - curUnitSel.unitDef.verTubeLoc + 1) * 32, (Int(y / 32) - curUnitSel.unitDef.horTubeLoc + 1) * 32, curUnitSel.unitDef.sizeX * 32, curUnitSel.unitDef.sizeY * 32
                    'shpHTube.X1 = (Int(X / 32) + curUnitSel.unitDef.sizeX + 1) * 32
                    shpHTube.x1 = (Int(shpCursor.Left / 32) + curUnitSel.unitDef.sizeX + 1) * 32
                    shpHTube.x2 = shpHTube.x1 + 32
                    'shpHTube.Y1 = ((Int(Y / 32) + curUnitSel.unitDef.horTubeLoc) * 32) + 16
                    shpHTube.y1 = ((Int(shpCursor.TOp / 32) + curUnitSel.unitDef.horTubeLoc) * 32) + 16
                    shpHTube.y2 = shpHTube.y1
                    'shpVTube.X1 = ((Int(X / 32) + curUnitSel.unitDef.verTubeLoc) * 32) + 16
                    shpVTube.x1 = ((Int(shpCursor.Left / 32) + curUnitSel.unitDef.verTubeLoc) * 32) + 16
                    shpVTube.x2 = shpVTube.x1
                    'shpVTube.Y1 = (Int(Y / 32) + curUnitSel.unitDef.sizeY + 1) * 32
                    shpVTube.y1 = (Int(shpCursor.TOp / 32) + curUnitSel.unitDef.sizeY + 1) * 32
                    shpVTube.y2 = shpVTube.y1 + 32
                    shpInnerUnit.Visible = True
                    shpHTube.Visible = True: shpVTube.Visible = True
                Else
                    'No tubes
                    shpHTube.Visible = False: shpVTube.Visible = False
                    shpCursor.Move (Int(x / 32) - 1) * 32, (Int(y / 32) - 1) * 32, (curUnitSel.unitDef.sizeX + 2) * 32, (curUnitSel.unitDef.sizeY + 2) * 32
                    shpInnerUnit.Move Int(x / 32) * 32, Int(y / 32) * 32, curUnitSel.unitDef.sizeX * 32, curUnitSel.unitDef.sizeY * 32
                    shpInnerUnit.Visible = True
                End If
            Else
                'Is a vehicle/other object
                shpInnerUnit.Visible = False
                shpHTube.Visible = False: shpVTube.Visible = False
                shpCursor.Move Int(x / 32) * 32, Int(y / 32) * 32, 32, 32
            End If
            'Handle mouse downs
            If buttonDown Then
27
                'Save undo step
                ReDim Preserve undoHist(numUndos) As UndoRec
                undoHist(numUndos).typeChg = UndoUnit
                undoHist(numUndos).savedUnitId = numUnitRecs
                numUndos = numUndos + 1
                'Do operation
                ReDim Preserve unitRecs(numUnitRecs) As UnitRec
                unitRecs(numUnitRecs).locX = (x \ 32) + hsbScroll.Value
                unitRecs(numUnitRecs).locY = (y \ 32) + vsbScroll.Value
                unitRecs(numUnitRecs).playerNum = myTool.CurrentPlayer
                unitRecs(numUnitRecs).uType = curUnitSel.unitDef
                unitRecs(numUnitRecs).wType = curUnitSel.weaponDef
                numUnitRecs = numUnitRecs + 1
                myMinimap.UpdatePixel (x \ 32) + hsbScroll.Value, (y \ 32) + vsbScroll.Value
            End If
    End Select
End If
JustDebounce:

If buttonDown Then
    'Update debouncer data if a button is down (i.e. they are drawing)
    oldX = Int(x / 32) + hsbScroll.Value
    oldY = Int(y / 32) + vsbScroll.Value
End If
Exit Sub
oops:
GenerateError "Map mouse event handling error", "MapManager::HandleMouseCmds"
End Sub

'Tool Architecture:
'Tool forms must implement the following method
'SetNewName(mgrName As String) - this is called to set part of the title to the filename or other name associated with the manager
'In addition it must implement any other methods or variables that the underlying implementation needs to work.
'(Eg. stuff to indicate the current selected item, to force a redraw, etc).
'** If these methods don't exist a runtime or shell error can be expected. **

Public Sub ChangeTool(toolType As ToolModeConstants)
'Stop the copy operation
If copyStep > 0 Then
    copyStep = 0
    SetStatusBar "Copy operation aborted."
End If
'Stop the paste operation
If pasteActive Then
    pasteActive = False
    SetStatusBar "Paste operation aborted."
End If
'Close any active tool
If myTool Is Nothing Then GoTo SelectNewTool
myTool.prepareClose = True
Unload myTool
Set myTool = Nothing
SelectNewTool:
Select Case toolType
    Case PlaceTile
        Set myTool = New frmTileset
        Set myTool.ownerForm = Me
        myTool.SetNewName myName
        myTool.SetNewMgr myMap.TileSetManager
    Case PlaceTileGroup
        Set myTool = New frmTileGroups
        'Set myTool.ownerForm = Me
        myTool.SetNewName myName
        myTool.SetNewMap myMap
        massMode = False
        massStep = 0
        myToolbar.ButtonChecked("massmode") = False
        SetStatusBar "Ready"

    Case PlaceUnit
        Set myTool = New frmUnits
        'Set myTool.ownerForm = Me
        myTool.SetNewName myName
        massMode = False
        massStep = 0
        myToolbar.ButtonChecked("massmode") = False
        SetStatusBar "Ready"

    Case CellTypeEdit
        Set myTool = New frmCellType
        'Set myTool.ownerForm = Me
        myTool.SetNewName myName
    Case InfoView
        '**TODO** this needs a form too
        massMode = False
        massStep = 0
        myToolbar.ButtonChecked("massmode") = False
        SetStatusBar "Ready"

End Select
toolMode = toolType
RedrawSelf

'Redraw minimap
If myMinimap Is Nothing Then Exit Sub
myMinimap.SetNewMap Me
End Sub

Public Sub SetNewName(ByVal strName As String, ByVal sTitle As String)
On Error GoTo baha
'myName = long name as in full path to the file
'myTitle = short name, just the filetitle or other
myName = strName
myTitle = sTitle
Me.Caption = myTitle
'Set this in the tool as well
myTool.SetNewName myTitle

'Set it on the tab bar
fMainForm.DTabControl.Tabs.Item("MAP" & CStr(myTabId)).Caption = Me.Caption

'And minimap
myMinimap.SetNewName myTitle
Exit Sub
baha:
GenerateError "SetNewName failed", "MapManager::SetNewName"
End Sub

Public Sub BeginCopy()
'start the copy operation
copyStep = 1
SetStatusBar "Click upper left corner of selection area"
End Sub

Public Sub BeginPaste()
'start the paste operation
If copyBuffer Is Nothing Then Beep: Exit Sub
pasteActive = True
SetStatusBar "Click location to paste tiles at"
End Sub

Public Sub SetScroll(ByVal xPos As Long, ByVal yPos As Long)
On Error Resume Next
'Move the scrollbars
hsbScroll.Value = xPos
vsbScroll.Value = yPos
RedrawSelf
End Sub

Public Sub UndoLast()
Dim i As Long, j As Long, x As Long, y As Long
'Undo the last operation, if possible
If numUndos = 0 Then Beep: SetStatusBar "Nothing left to undo": Exit Sub
Select Case undoHist(numUndos - 1).typeChg
    Case UndoTile
        'Undo the last tile
        myMap.mappingIndex(undoHist(numUndos - 1).xPos, undoHist(numUndos - 1).yPos) = undoHist(numUndos - 1).savedMapping
        myMinimap.UpdatePixel undoHist(numUndos - 1).xPos, undoHist(numUndos - 1).yPos
    Case UndoGroup
        'Undo the last group
        myMap.Paste undoHist(numUndos - 1).xPos, undoHist(numUndos - 1).yPos, undoHist(numUndos - 1).savedGroup
        'See if they pasted cell types and fix those as well
        For i = undoHist(numUndos - 1).xPos To undoHist(numUndos - 1).xPos + undoHist(numUndos - 1).savedGroup.TileWidth - 1
            For j = undoHist(numUndos - 1).yPos To undoHist(numUndos - 1).yPos + undoHist(numUndos - 1).savedGroup.TileHeight - 1
                If copyPasteCellTypes Then
                    myMap.CellType(i, j) = undoHist(numUndos - 1).savedMassCellType((j - undoHist(numUndos - 1).yPos) * undoHist(numUndos - 1).savedGroup.TileWidth + (i - undoHist(numUndos - 1).xPos))
                End If
                myMinimap.UpdatePixel i, j
            Next
            DoEvents
        Next
    Case UndoCellType
        'Undo the last cell type change
        myMap.CellType(undoHist(numUndos - 1).xPos, undoHist(numUndos - 1).yPos) = undoHist(numUndos - 1).savedCellType
        myMinimap.UpdatePixel undoHist(numUndos - 1).xPos, undoHist(numUndos - 1).yPos
    Case UndoUnit
        x = unitRecs(undoHist(numUndos - 1).savedUnitId).locX
        y = unitRecs(undoHist(numUndos - 1).savedUnitId).locY
        'Remove the existing unit
        For i = undoHist(numUndos - 1).savedUnitId + 1 To numUnitRecs - 1
            unitRecs(i - 1) = unitRecs(i)
        Next
        
        numUnitRecs = numUnitRecs - 1
        If numUnitRecs = 0 Then
            Erase unitRecs
        Else
            ReDim Preserve unitRecs(numUnitRecs) As UnitRec
        End If
        
        myMinimap.UpdatePixel x, y
    Case UndoMassTile
        'Undo mass tile
        For i = undoHist(numUndos - 1).xPos To undoHist(numUndos - 1).x2Pos
            For j = undoHist(numUndos - 1).yPos To undoHist(numUndos - 1).y2Pos
                myMap.mappingIndex(i, j) = undoHist(numUndos - 1).savedMassMapping(i - undoHist(numUndos - 1).xPos, j - undoHist(numUndos - 1).yPos)
                myMinimap.UpdatePixel i, j
            Next
            DoEvents
        Next
    Case UndoMassCellType
        'Undo mass celltype
        For i = undoHist(numUndos - 1).xPos To undoHist(numUndos - 1).x2Pos
            For j = undoHist(numUndos - 1).yPos To undoHist(numUndos - 1).y2Pos
                myMap.CellType(i, j) = undoHist(numUndos - 1).savedMassCellType(i - undoHist(numUndos - 1).xPos, j - undoHist(numUndos - 1).yPos)
                myMinimap.UpdatePixel i, j
            Next
            DoEvents
        Next
End Select
'Remove last record
numUndos = numUndos - 1
If numUndos = 0 Then
    Erase undoHist
Else
    ReDim Preserve undoHist(numUndos - 1) As UndoRec
End If
RedrawSelf
End Sub

Public Sub ToolDeactivated()
''If we have gotten focus, do nothing
'If fMainForm.ActiveForm Is Me Then Exit Sub
'Debug.Print "ToolDeactivated"
''Otherwise, hide the tools since another form is in focus
'myTool.Hide
'myMinimap.Hide
End Sub

Public Sub PickupCurTile()
'Message from menu to select tile under the cursor
If toolMode <> PlaceTile Then Beep: SetStatusBar "Not in Place Tile mode": Exit Sub
myTool.SetCurTile myMap.mappingIndex(Int(curX / 32) + hsbScroll.Value, Int(curY / 32) + vsbScroll.Value)
End Sub

Public Sub DelCurUnit()
Dim i As Long, j As Long
'Message from menu to delete unit under the cursor
If toolMode <> PlaceUnit Then Beep: SetStatusBar "Not in Place Object mode": Exit Sub

'Loop backwards through the units
If numUnitRecs > 0 Then
    For i = numUnitRecs - 1 To 0 Step -1
        If unitRecs(i).locX = Int(curX / 32) + hsbScroll.Value And unitRecs(i).locY = Int(curY / 32) + vsbScroll.Value Then
            'Unit found; delete it
            For j = i To numUnitRecs - 2
                unitRecs(j) = unitRecs(j + 1)
            Next
            numUnitRecs = numUnitRecs - 1
            ReDim Preserve unitRecs(numUnitRecs) As UnitRec
            Exit Sub
        End If
    Next
    Beep
    SetStatusBar "No unit underneath the cursor"
    Exit Sub
End If
Beep
SetStatusBar "No units on the map"
Exit Sub
End Sub

Public Sub CelltypeReplace()
'Message from menu to do a cell type replace operation
frmCelltypeReplace.Show 1, fMainForm
If frmCelltypeReplace.goAhead = True Then
    SetStatusBar "Scanning map..."
    Dim i As Long, j As Long
    For i = 0 To myMap.TileWidth - 1
        For j = 0 To myMap.TileHeight - 2 'Bottom row is untouchable
            If myMap.CellType(i, j) = frmCelltypeReplace.currentSel1 Then
                myMap.CellType(i, j) = frmCelltypeReplace.currentSel2
                myMinimap.UpdatePixel i, j
            End If
        Next
        DoEvents
    Next
    RedrawSelf
    SetStatusBar "Celltype Replace successful"
End If
Unload frmCelltypeReplace
End Sub

Public Sub GenerateCode()
On Error GoTo oops
Dim generatedCode As String, f As Integer

'Message from File > Compile menu
Load frmGenUnitCode
frmGenUnitCode.txtFilename = myName & ".cpp"
'show the form, wait for user response
frmGenUnitCode.Show 1, fMainForm
If frmGenUnitCode.templateName = "" Or frmGenUnitCode.outputName = "" Then Exit Sub

'Parse the code file
generatedCode = ParseAndGenerateUnits(App.Path & "\" & frmGenUnitCode.templateName, unitRecs, numUnitRecs, myTitle)
If generatedCode = "" Then Exit Sub

'Write it
On Error Resume Next
If FileLen(frmGenUnitCode.outputName) > 0 Then Kill frmGenUnitCode.outputName
On Error GoTo oops
f = FreeFile
Open frmGenUnitCode.outputName For Binary As #f
Put #f, , generatedCode
Close #f

Unload frmGenUnitCode

SetStatusBar "Code generated successfully."
MsgBox "Code generated successfully."

Exit Sub

oops:
GenerateError "Code generation failed.", "MapManager::GenerateCode"
End Sub

Public Sub UnitRecord(ByVal whichOne As Long, ByVal retUnitPtr As Long)
    'Stupid accessor method, VB doesn't like public arrays in forms
    CopyMemory ByVal retUnitPtr, unitRecs(whichOne), Len(unitRecs(whichOne))
End Sub

Public Sub SaveAutosave()
Dim stream As StreamWriter
Set stream = ResMan.OpenStreamWrite(App.Path & "\autosave.map")
myMap.SaveMap stream, TileGroups
Set stream = Nothing
'Save the unit definitions
If numUnitRecs > 0 Then
    Dim f As Integer, i As Long
    f = FreeFile
    If Dir$(App.Path & "\autosave.map.dat") <> "" Then Kill App.Path & "\autosave.map.dat"
    Open App.Path & "\autosave.map.dat" For Binary As #f
    Put #f, , numUnitRecs
    For i = 0 To numUnitRecs - 1
        Put #f, , unitRecs(i).locX
        Put #f, , unitRecs(i).locY
        Put #f, , unitRecs(i).playerNum
        Put #f, , unitRecs(i).uType.mapID
        Put #f, , unitRecs(i).wType.mapID
        Put #f, , unitRecs(i).uType.isGaia
        If unitRecs(i).uType.isGaia = True Then
            'Save the gaia extra data
            Put #f, , unitRecs(i).uType.extra1
            Put #f, , unitRecs(i).uType.extra2
            Put #f, , unitRecs(i).uType.extra3
        End If
    Next
    Close #f
End If
End Sub
