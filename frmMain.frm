VERSION 5.00
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "OP2Mapper"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAutosave 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   840
   End
   Begin vbalDTab6.vbalDTabControl DTabControl 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5295
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCloseButton =   0   'False
   End
   Begin cPopMenu6.PopMenu cMenu 
      Left            =   0
      Top             =   840
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      HighlightStyle  =   2
   End
   Begin vbalTBar6.cReBar cReBar 
      Left            =   0
      Top             =   360
      _ExtentX        =   3836
      _ExtentY        =   661
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   492
      TabIndex        =   1
      Top             =   0
      Width           =   7380
      Begin vbalTBar6.cToolbar cTbProj 
         Height          =   375
         Left            =   2520
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
      End
      Begin vbalTBar6.cToolbar cTbMain 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   7380
      TabIndex        =   0
      Top             =   5670
      Width           =   7380
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "&Generate Code..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCellTypes 
         Caption         =   "C&opy/Paste Cell Types"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewStartOP2 
         Caption         =   "Start &Outpost 2"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuProj 
      Caption         =   "&Project"
      Enabled         =   0   'False
      Begin VB.Menu mnuProjMap 
         Caption         =   "Add New &Map"
      End
      Begin VB.Menu mnuProjScript 
         Caption         =   "Add New &Script"
      End
      Begin VB.Menu mnuProjTech 
         Caption         =   "Add New &Techtree"
      End
      Begin VB.Menu mnuProjBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjFile 
         Caption         =   "&Add File..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuProjBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProjRem 
         Caption         =   "&Remove Current File"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuPlugins 
      Caption         =   "&Plugins"
      Enabled         =   0   'False
      Begin VB.Menu mnuPluginsConfig 
         Caption         =   "&Configure Plugins..."
      End
      Begin VB.Menu mnuPluginsBar0 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPluginsList 
         Caption         =   "Plugin List"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
   Begin VB.Menu mnuMinimap 
      Caption         =   "Minimap"
      Visible         =   0   'False
      Begin VB.Menu mnuMinimapSave 
         Caption         =   "&Save Minimap..."
      End
   End
   Begin VB.Menu mnuGroup 
      Caption         =   "TileGroup"
      Visible         =   0   'False
      Begin VB.Menu mnuGroupCreate 
         Caption         =   "&Create Group from Clipboard..."
      End
      Begin VB.Menu mnuGroupDelete 
         Caption         =   "&Delete Selected Group"
      End
   End
   Begin VB.Menu mnuTileset 
      Caption         =   "Tileset"
      Visible         =   0   'False
      Begin VB.Menu mnuTilesetPaste 
         Caption         =   "&Paste Consecutive Tiles by Click..."
      End
      Begin VB.Menu mnuTilesetPaste2 
         Caption         =   "&Paste Consecutive Tiles by ID..."
      End
      Begin VB.Menu mnuTilesetBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTilesetEnableGrid 
         Caption         =   "&Enable Grid"
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "Map"
      Visible         =   0   'False
      Begin VB.Menu mnuMapSelTile 
         Caption         =   "&Select Tile Under Cursor"
      End
      Begin VB.Menu mnuMapUnitDel 
         Caption         =   "&Delete Topmost Unit Under Cursor"
      End
      Begin VB.Menu mnuMapCelltypes 
         Caption         =   "&Replace Celltypes..."
      End
      Begin VB.Menu mnuMapBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMapEnableGrid 
         Caption         =   "&Enable Grid"
      End
   End
   Begin VB.Menu mnuTilesetEditAdd 
      Caption         =   "TilesetEditAdd"
      Visible         =   0   'False
      Begin VB.Menu mnuTilesetEditAddOne 
         Caption         =   "Add &Single Tile"
      End
      Begin VB.Menu mnuTilesetEditAddDir 
         Caption         =   "Add &Directory"
      End
   End
   Begin VB.Menu mnuTilesetEditExp 
      Caption         =   "TilesetEditExp"
      Visible         =   0   'False
      Begin VB.Menu mnuTilesetEditExpOne 
         Caption         =   "Export &Selected Tile"
      End
      Begin VB.Menu mnuTilesetEditExpAll 
         Caption         =   "Export &All Tiles"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Public DebugMapFile As MapManager

Public sbStatusBar As New cNoStatusBar

Public cCMDlg As New cCommonDialog

Private Sub cReBar_HeightChanged(lNewHeight As Long)
picToolbar.Height = picToolbar.ScaleY(lNewHeight, vbPixels, vbTwips)
End Sub

Private Sub cTbMain_ButtonClick(ByVal lButton As Long)
Select Case cTbMain.ButtonKey(lButton)
    Case "new"
        '**May want to change this to create a new map/script/techtree
        mnuFileNew_Click
    Case "open"
        mnuFileOpen_Click
    Case "save"
        mnuFileSave_Click
    Case "print"
        mnuFilePrint_Click
    Case "props"
        mnuFileProperties_Click
    Case "undo"
        mnuEditUndo_Click
    Case "cut"
        mnuEditCut_Click
    Case "copy"
        mnuEditCopy_Click
    Case "paste"
        mnuEditPaste_Click
    Case "runscript"
        mnuPluginsConfig_Click
    Case "op2"
        mnuViewStartOP2_Click
End Select
End Sub

Private Sub cTbProj_ButtonClick(ByVal lButton As Long)
Select Case cTbProj.ButtonKey(lButton)
'**TODO**
    Case "addnew"
    Case "addexist"
    Case "delsel"
End Select
End Sub

Private Sub DTabControl_TabSelected(theTab As cTab)
'Find the form that has the current caption
'On Error Resume Next
Dim f As Form
For Each f In Forms
    If f.tag = theTab.Key Then
        If f.MDIChild = True Then
            'Activate it
            f.Show
            f.ZOrder 0
            Exit Sub
        End If
    End If
Next
End Sub

Private Sub MDIForm_Load()
    Me.Left = GetSettingIni("Window", "MainLeft", 1000)
    Me.TOp = GetSettingIni("Window", "MainTop", 1000)
    Me.Width = GetSettingIni("Window", "MainWidth", 6500)
    Me.Height = GetSettingIni("Window", "MainHeight", 6500)
    Me.WindowState = GetSettingIni("Window", "MainWindowState", 0)
End Sub

Private Sub MDIForm_QueryUnload(cancel As Integer, UnloadMode As Integer)
cReBar.RemoveAllRebarBands

'Close all map windows
Dim f As Form
For Each f In Forms
    If f.Name = "MapManager" Then Unload f
Next
Set f = Nothing
End Sub

Private Sub MDIForm_Resize()
cReBar.RebarSize
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSettingIni "Window", "MainLeft", Me.Left
        SaveSettingIni "Window", "MainTop", Me.TOp
        SaveSettingIni "Window", "MainWidth", Me.Width
        SaveSettingIni "Window", "MainHeight", Me.Height
        SaveSettingIni "Window", "MainWindowState", Me.WindowState
    End If
#If NoInitResMan = 0 Then
    'Unload the resource manager
    Set mapsVolStream = Nothing
    Set ResMan = Nothing
#End If
End Sub

Private Sub mnuEditCellTypes_Click()
copyPasteCellTypes = Not copyPasteCellTypes
mnuEditCellTypes.Checked = copyPasteCellTypes

SaveSettingIni "Options", "CopyPasteCellTypes", IIf(copyPasteCellTypes, "1", "0")
End Sub

Private Sub mnuFileNew_Click()
'Invoke the new project form
frmNew.Show 1, Me
'MsgBox "Map creation isn't implemented yet. If you need a new map, open an existing one of the desired size, then use the Save As command."
End Sub

Private Sub mnuGroupCreate_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "frmTileGroups" Then
    Dim sName As String
    If copyBuffer Is Nothing Then
        MsgBox "You must have copied a section of tiles which will be used to create the new tile group.", vbInformation, "Must Copy Data"
        Exit Sub
    End If
    '**TODO** make a better input prompt
    sName = InputBox("Enter the desired name for your new tile group. This name must not clash with an existing name in the map.", "Tile Group Name")
    If Trim(sName) = "" Then Exit Sub
    ActiveForm.AddGroup Trim(sName)
End If
End Sub

Private Sub mnuGroupDelete_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "frmTileGroups" Then ActiveForm.DelCurGroup
End Sub

Private Sub mnuMapCelltypes_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "MapManager" Then ActiveForm.CelltypeReplace
End Sub

Private Sub mnuMapEnableGrid_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "MapManager" Then
    'Toggle the grid
    ActiveForm.enableGrid = Not ActiveForm.enableGrid
    ActiveForm.myToolbar.ButtonChecked("grid") = ActiveForm.enableGrid
    ActiveForm.RedrawSelf
End If
End Sub

Private Sub mnuMapSelTile_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "MapManager" Then ActiveForm.PickupCurTile
End Sub

Private Sub mnuMapUnitDel_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "MapManager" Then ActiveForm.DelCurUnit
End Sub

Private Sub mnuPluginsConfig_Click()
'**TODO**
End Sub

Private Sub mnuProjFile_Click()
'**TODO**
MsgBox "Add 'mnuProjFile_Click' code."
End Sub

Private Sub mnuProjMap_Click()
If openWorkspace Is Nothing Then Exit Sub
openWorkspace.AddMap "", ""
End Sub

Private Sub mnuProjScript_Click()
If openWorkspace Is Nothing Then Exit Sub
openWorkspace.AddScript ""
End Sub

Private Sub mnuProjTech_Click()
If openWorkspace Is Nothing Then Exit Sub
openWorkspace.AddTechtree ""
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuTilesetEditAddDir_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "TilesetEditMgr" Then ActiveForm.AddDir
End Sub

Private Sub mnuTilesetEditAddOne_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "TilesetEditMgr" Then ActiveForm.AddOne
End Sub

Private Sub mnuTilesetEditExpAll_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "TilesetEditMgr" Then ActiveForm.ExportAll
End Sub

Private Sub mnuTilesetEditExpOne_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "TilesetEditMgr" Then ActiveForm.ExportOne
End Sub

Private Sub mnuTilesetEnableGrid_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "frmTileset" Then
    'Toggle the grid
    ActiveForm.enableGrid = Not ActiveForm.enableGrid
    ActiveForm.RedrawSelf
End If
End Sub

Private Sub mnuTilesetPaste_Click()
'Paste tiles from the current tileset consecutively into a rect on the map,
'useful for recreating objects stored in different tilesets
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "frmTileset" Then ActiveForm.PasteConsecutive -1
End Sub

Private Sub mnuTilesetPaste2_Click()
Dim strVal As String
'Paste consecutively based on index the user enters
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "frmTileset" Then
    strVal = InputBox("Please enter the starting tile ID number to copy from. (Decimal number system)")
    If Not IsNumeric(strVal) Then Exit Sub
    ActiveForm.PasteConsecutive CLng(strVal)
End If
End Sub

Private Sub mnuViewStartOP2_Click()
On Error GoTo oops
'Run OP2 exe
SetStatusBar "Starting Outpost 2..."
Shell ResMan.RootPath & "\Outpost2.exe", vbNormalFocus
SetStatusBar "Ready"
Exit Sub
oops:
GenerateError "Could not start Outpost2.exe", "fMainForm::mnuViewStartOP2_Click"
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuViewOptions_Click()
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Then ActiveForm.RedrawSelf
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Then ActiveForm.BeginPaste
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Then ActiveForm.BeginCopy
End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next

End Sub

Private Sub mnuEditUndo_Click()
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Then ActiveForm.UndoLast
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Then ActiveForm.GenerateCode
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    'If ActiveForm Is Nothing Then Exit Sub
    MsgBox "The print functions aren't implemented."
End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "The print functions aren't implemented."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    MsgBox "The print functions aren't implemented."
End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Properties aren't implemented yet."
End Sub

Private Sub mnuFileSaveAll_Click()
    On Error Resume Next
    MsgBox "Buggy, doesn't work yet."
    'If ActiveForm Is Nothing Then Exit Sub
    'Dim f As Form
    'For Each f In Forms
    '    If f.Name = "MapManager" Then ActiveForm.SaveSelf
    'Next
End Sub

Private Sub mnuFileSaveAs_Click()
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Or _
        ActiveForm.Name = "TilesetEditMgr" Or _
        ActiveForm.Name = "VolManager" Then ActiveForm.SaveSelfAs
End Sub

Private Sub mnuFileSave_Click()
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Or _
        ActiveForm.Name = "TilesetEditMgr" Or _
        ActiveForm.Name = "VolManager" Then ActiveForm.SaveSelf
End Sub

Private Sub mnuFileClose_Click()
    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.Name = "MapManager" Or _
        ActiveForm.Name = "TilesetEditMgr" Or _
        ActiveForm.Name = "VolManager" Then Unload ActiveForm
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String, sTitle As String, mapForm As MapManager, bmpForm As TilesetEditMgr, volForm As VolManager
    If cCMDlg.VBGetOpenFileName(sFile, sTitle, , , , True, "All Supported Files|*.map;*.bmp;*.vol|OP2 Map Files (*.map)|*.map|Tileset Bitmaps (*.bmp)|*.bmp|VOL Archives (*.vol)|*.vol", , mapsDir, "Load File", "map") = False Then Exit Sub
    
    Select Case LCase$(Right$(sFile, 3)) 'Decide how to open the file
        Case "map" 'Map file
            Set mapForm = New MapManager
            If mapForm.LoadMap(sFile) Then
                mapForm.SetNewName sFile, sTitle
                mapForm.Show
            Else
                Set mapForm = Nothing
            End If
        Case "bmp" 'Tileset file
            Set bmpForm = New TilesetEditMgr
            If bmpForm.LoadTileset(sFile) Then
                bmpForm.SetNewName sFile, sTitle
                bmpForm.Show
            Else
                Set bmpForm = Nothing
            End If
        Case "vol" 'VOL file
            Set volForm = New VolManager
            If volForm.LoadVol(sFile) Then
                volForm.SetNewName sFile, sTitle
                volForm.Show
            Else
                Set volForm = Nothing
            End If
        Case Else
            'Extension not recognized
            SetStatusBar "Extension ." & LCase$(Right$(sFile, 3)) & " was not recognized, and the file was not loaded."
    End Select
End Sub

Private Sub picStatus_Paint()
sbStatusBar.Draw
End Sub

Private Sub picToolbar_Paint()
cReBar.RebarSize
End Sub

Private Sub mnuMinimapSave_Click()
If ActiveForm Is Nothing Then Exit Sub
If ActiveForm.Name = "frmMinimap" Then
    'Save the minimap image
    Dim sFile As String
    If fMainForm.cCMDlg.VBGetSaveFileName(sFile, , , "Windows Bitmap (*.bmp)|*.bmp", , mapsDir, "Save Minimap Image", "bmp") = False Then Exit Sub
    SavePicture ActiveForm.picMinimap.Image, sFile
End If
End Sub

Public Sub OpenMap(ByVal sFile As String)
'Open a map file (from the command line)
Dim mapForm As MapManager, sTitle As String, pos As Long
Set mapForm = New MapManager
If mapForm.LoadMap(sFile) Then
    'Split last part from the path name (File Title)
    'Convert / to \
    sFile = Replace(sFile, "/", "\", , , vbTextCompare)
    'Search for path delimiters
    pos = InStrRev(sFile, "\", , vbTextCompare)
    If pos <> 0 Then
        sTitle = Mid$(sFile, pos + 1)
    Else
        sTitle = sFile
    End If
    mapForm.SetNewName sFile, sTitle
    mapForm.Show
Else
    Set mapForm = Nothing
End If
End Sub

Public Sub OpenVol(ByVal sFile As String)
'Open a Vol file (from the command line)
Dim volForm As VolManager, sTitle As String, pos As Long
Set volForm = New VolManager
If volForm.LoadVol(sFile) Then
    'Split last part from the path name (File Title)
    'Convert / to \
    sFile = Replace(sFile, "/", "\", , , vbTextCompare)
    'Search for path delimiters
    pos = InStrRev(sFile, "\", , vbTextCompare)
    If pos <> 0 Then
        sTitle = Mid$(sFile, pos + 1)
    Else
        sTitle = sFile
    End If
    volForm.SetNewName sFile, sTitle
    volForm.Show
Else
    Set volForm = Nothing
End If
End Sub

Private Sub tmrAutosave_Timer()
'First increment the seconds
curAutosaveSec = curAutosaveSec + 1
'If 60 sec then carry over
If curAutosaveSec = 60 Then
    curAutosaveSec = 0
    curAutosaveMin = curAutosaveMin + 1
    If curAutosaveMin = Minute(autoSaveAfter) Then
        'The clock has struck, let's save the file
        On Error Resume Next
        If ActiveForm.Name = "MapManager" Or _
            ActiveForm.Name = "TilesetEditMgr" Or _
            ActiveForm.Name = "VolManager" Then ActiveForm.SaveAutosave
        curAutosaveMin = 0
        curAutosaveSec = 0
    End If
End If
'Change the status
SetStatusAutosave tmrAutosave.Enabled, TimeSerial(0, curAutosaveMin, curAutosaveSec)
End Sub
