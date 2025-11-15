Attribute VB_Name = "modMain"
Option Explicit

'Public Declare Sub InitCommonControls Lib "COMCTL32" ()

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function CoGetClassObjectFromFile Lib "ComLoad.dll" (clsid As IClassFactory_TLB.GUID, ByVal Filename As String, lpObj As Any) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public Type RECT
    Left As Long
    TOp As Long
    Right As Long
    Bottom As Long
End Type

Public Type SHITEMID
    cb As Long
    abID() As Byte
End Type
Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Const VER_PLATFORM_WIN32_NT As Long = 2
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1

Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const WM_MDIACTIVATE = &H222
Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const WM_KEYDOWN = &H100
Public Const KEY_PRESSED As Integer = &H1000

Public Const DT_CENTER As Long = &H1
Public Const DT_VCENTER As Long = &H4
Public Const DT_SINGLELINE As Long = &H20

Public Const INIFILENAME = "\Mapper2.ini"

Public fMainForm As frmMain

Public ResMan As ResourceManager

Public openWorkspace As New Workspace
Public isWorkspaceLoaded As Boolean 'used at program start to know if the object has been inited
Public mapsVolStream As ArchiveReader

Public PRTFile As New CPrtFile

Public docNumber As Long

Public mainIml As New cVBALImageList

Public Enum ToolModeConstants 'map editor mode constants
    PlaceTile = 0
    PlaceTileGroup = 1
    PlaceUnit = 2
    CellTypeEdit = 3
    InfoView = 4
End Enum

Public Enum CellTypes 'taken from game
    FastPassable1 = &H0
    Impassable2 = &H1
    SlowPassable1 = &H2
    SlowPassable2 = &H3
    MediumPassable1 = &H4
    MediumPassable2 = &H5
    Impassable1 = &H6
    FastPassable2 = &H7
    NorthCliffs = &H8
    CliffsHighSide = &H9
    CliffsLowSide = &HA
    VentsAndFumaroles = &HB
    zPad12 = &HC
    zPad13 = &HD
    zPad14 = &HE
    zPad15 = &HF
    zPad16 = &H10
    zPad17 = &H11
    zPad18 = &H12
    zPad19 = &H13
    zPad20 = &H14
    DozedArea = &H15
    rubble = &H16
    normalWall = &H17
    microbeWall = &H18
    lavaWall = &H19
    Tube0 = &H1A
    Tube1 = &H1B
    Tube2 = &H1C
    Tube3 = &H1D
    Tube4 = &H1E
    Tube5 = &H1F
End Enum

Public copyPasteCellTypes As Boolean

Public copyBuffer As TileGroup 'map copy buffer
Public cellTypeBuffer() As CellTypes

Public mapsDir As String 'Default save dir

Public defTilesets() As String 'Default tilesets. Entry 0 is always used
Public numTilesets As Long 'Num default tilesets

Public Type TerrainData
    startTile As Long
    endTile As Long
    dozed As Long
    rubble As Long
    tubeUnk(5) As Long
    lavaWall(15) As Long
    microbeWall(15) As Long
    normalWall(15) As Long
    damagedWall(15) As Long
    ruinedWall(15) As Long
    lava As Long
    flat1 As Long
    flat2 As Long
    flat3 As Long
    tube(15) As Long
    scorched As Long
    unkTile(20) As Long
End Type
Public defTerrains() As TerrainData 'Default terrains
Public numTerrains As Long 'Num default terrains

'Unit support stuff.
Public Type UnitType
    mapID As Long
    artID As Long
    strName As String
    isStructure As Boolean
    sizeX As Integer
    sizeY As Integer
    hasTubes As Boolean
    verTubeLoc As Integer
    horTubeLoc As Integer
    canHaveTurret As Boolean
    'Gaia object stuff
    isGaia As Boolean
    gaiaType As Integer
    extra1 As Integer
    extra2 As Integer
    extra3 As Integer
End Type
Public allUnitDefs() As UnitType
Public numUnitDefs As Long

Public Type WeaponType
    mapID As Long
    strName As String
    artID As Long
End Type
Public allWeaponDefs() As WeaponType
Public numWeaponDefs As Long

'Unit save/load/list stuff.
Public Type UnitRec
    playerNum As Long
    locX As Long
    locY As Long
    uType As UnitType
    wType As WeaponType
End Type

'Unit selector
Public Type UnitEntry
    strName As String
    unitDef As UnitType
    weaponDef As WeaponType
End Type

'this var must be updated frequently
Public curUnitSel As UnitEntry

Public doNotActivateMDI As Boolean

'Undo record
Public Enum UndoChgType
    UndoTile
    UndoGroup
    UndoCellType
    UndoUnit
    UndoMassTile
    UndoMassCellType
End Enum
Public Type UndoRec
    typeChg As UndoChgType
    savedMapping As Long
    savedGroup As TileGroup
    savedCellType As CellTypes
    savedMassMapping() As Long
    savedMassCellType() As CellTypes
    savedUnitId As Long
    xPos As Long
    yPos As Long
    x2Pos As Long
    y2Pos As Long
End Type

'INI Settings
Public enableGridDef As Boolean
Public enableKbPoll As Boolean
Public autoSaveAfter As Date
Public curAutosaveMin As Long
Public curAutosaveSec As Long

'Plugin stuff
'**TODO**: Add TypeLib Information DLL (Tlbinf32.dll) to references
Public Type PluginMenuItem
    strName As String
    lngCode As Long
End Type
Public Type PluginEventHook
    lngHookType As Long
    lngCode As Long
End Type
Public Type PluginCtx
    strFileName As String
    dllHandle As Long
    numMenuItems As Long
    menuItems() As PluginMenuItem
    numEventHooks As Long
    eventHooks() As PluginEventHook
    plugObject As IPluginBase
    appObject As IAppObject
End Type
Public privMode As Boolean
Public numLoadedPlugins As Long
Public loadedPlugins() As PluginCtx
Public testGuid As IClassFactory_TLB.GUID


Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767


'Public curProject As cProject
'Public mainTileset As New cTileset
'Public mainMinimap As New cMinimap
'
'Public newProjType As PROJECT_TYPE
'Public newMapX As Long
'Public newMapY As Long
'Public newInitTile As Integer
'Public canceledCreate As Boolean
'
'Public Enum PROJECT_TYPE
'    mapAndDll = 0
'    MapOnly = 1
'    dllOnly = 2
'    techtreeOnly = 3
'End Enum

'Public tileMan As New TileSetManager


Sub Main()
    On Error GoTo Augha
    numUnitDefs = -1: numWeaponDefs = -1
    'InitCommonControls
    
    frmSplash.Show
    frmSplash.Refresh
    frmSplash.lblStatus.Caption = "Loading shell..."
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    
    SetupUI
    
    frmSplash.lblStatus.Caption = "Reading units.ctl..."
    frmSplash.Refresh
    ParseUnitTable App.Path & "\units.ctl"
    
    frmSplash.lblStatus.Caption = "Reading weapons.ctl..."
    frmSplash.Refresh
    ParseWeaponTable App.Path & "\weapons.ctl"
    
    frmSplash.lblStatus.Caption = "Reading objects.ctl..."
    frmSplash.Refresh
    ParseObjectTable App.Path & "\objects.ctl"
    
    frmSplash.lblStatus.Caption = "Loading tilesets..."
    frmSplash.Refresh
#If NoInitResMan = 0 Then
    Set ResMan = New ResourceManager
    
    Dim rootStr As String
    rootStr = GetSettingIni("Paths", "Outpost2")
    'Ask the user for the folder if it's not cached in the INI
    If rootStr = "" Then
        rootStr = BrowseForFolder(frmSplash.hWnd, "Please specify the folder where your OP2 mapsXX.vol files are located.")
        SaveSettingIni "Paths", "Outpost2", rootStr
    End If
    ResMan.RootPath = rootStr
    
    'Check for and read additional VOLs
    Dim volStr As String, volArr() As String, i As Long
    volStr = GetSettingIni("Paths", "CustomVOL")
    If volStr <> "" Then
        volArr = Split(volStr, "|")
        On Error Resume Next
        For i = 0 To UBound(volArr)
            Set mapsVolStream = ResMan.LoadVolFile(volArr(i), True)
            If Err Or mapsVolStream Is Nothing Then
                MsgBox "Addon VOL file " & volArr(i) & " could not be loaded. It may be corrupt or nonexistant.", vbCritical, "Load Error"
            Else
                ResMan.AddArchiveToSearch mapsVolStream
            End If
        Next
        On Error GoTo Augha
    End If
    
    Set mapsVolStream = ResMan.LoadVolFile("maps.vol", True)
    ResMan.AddArchiveToSearch mapsVolStream
#End If

    'Ask user for their default save folder
    mapsDir = GetSettingIni("Paths", "Save")
    If mapsDir = "" Then
        mapsDir = BrowseForFolder(frmSplash.hWnd, "Please specify the default folder where you will be saving files.")
        SaveSettingIni "Paths", "Save", mapsDir
    End If
    
#If NoInitArt = 0 Then
    frmSplash.lblStatus.Caption = "Loading game art..."
    frmSplash.Refresh
    ReadPrt
#End If
        
    'Read tilesets.ctl for a list of tilesets to use
    frmSplash.lblStatus.Caption = "Reading tilesets.ctl..."
    frmSplash.Refresh
    ReadTilesetsCtl
    
    'INI settings
    enableKbPoll = IIf(GetSettingIni("Options", "DisableKBPoll", "0") = "1", False, True)
    enableGridDef = IIf(GetSettingIni("Options", "DisableGrid", "0") = "1", False, True)
    autoSaveAfter = TimeSerial(0, CLng(GetSettingIni("Options", "AutosaveMin", "5")), 0)
    copyPasteCellTypes = IIf(GetSettingIni("Options", "CopyPasteCellTypes", "1") = "1", True, False)
    fMainForm.mnuEditCellTypes.Checked = copyPasteCellTypes
    If CLng(GetSettingIni("Options", "AutosaveMin", "5")) > 0 Then fMainForm.tmrAutosave.Enabled = True
    curAutosaveMin = 0
    curAutosaveSec = 0
    SetStatusAutosave fMainForm.tmrAutosave.Enabled, TimeSerial(0, curAutosaveMin, curAutosaveSec)
    
    'Read terrains.ctl for a list of terrains to use
    frmSplash.lblStatus.Caption = "Reading terrains.ctl..."
    frmSplash.Refresh
    ReadTerrainsCtl
    
    Unload frmSplash

    fMainForm.Show
    On Error Resume Next
    frmTip.Show 1, fMainForm
    
    'Handle the commandline
    If Command <> "" Then ParseCommandLine
    
    
    Exit Sub
Augha:
    GenerateError "Loading of essential editor files failed.", "main()", True
End Sub

Sub LoadNewDoc()
'    If curProject Is Nothing Then
'        frmNewDialog.Show vbModal, fMainForm
'        If canceledCreate Then Exit Sub
'        'Create the project
'        Set curProject = New cProject
'        curProject.CreateProject newProjType, newMapX, newMapY, newInitTile
'    Else
'        frmNewDialog.Show vbModal, fMainForm
'        If canceledCreate Then Exit Sub
'        'Make sure the other project is closed OK
'        If curProject.CloseProject() = True Then
'            'Create the project
'            Set curProject = Nothing
'            Set curProject = New cProject
'            curProject.CreateProject newProjType, newMapX, newMapY, newInitTile
'        End If
'    End If
End Sub

Sub GenerateError(ByVal desc As String, ByVal calledFrom As String, Optional ByVal isFatal As Boolean = False)
    'Produce an error
    Dim fError As Form, osVerInfo As OSVERSIONINFO
    If isFatal Then
        Set fError = New frmFatalError
        fError.txtError = "Fatal "
    Else
        Set fError = New frmError
        fError.txtError = "Non-fatal "
    End If
    fError.txtError = fError.txtError & "error in " & calledFrom & ": " & desc & vbNewLine
    'Add the intrinsic error data, if any
    If Err Then
        fError.txtError = fError.txtError & vbNewLine & "Err = (" & CStr(Err.Number) & ") " & Err.Description & vbNewLine & _
            "GetLastError = " & GetLastError & ", Erl = " & Erl & vbNewLine
    End If
    'Add OS version information
    osVerInfo.dwOSVersionInfoSize = Len(osVerInfo)
    GetVersionEx osVerInfo
    Select Case osVerInfo.dwPlatformId
        Case VER_PLATFORM_WIN32_NT
            fError.txtError = fError.txtError & vbNewLine & "OS Version: NT "
        Case VER_PLATFORM_WIN32_WINDOWS
            fError.txtError = fError.txtError & vbNewLine & "OS Version: 9x "
        Case Else
            fError.txtError = fError.txtError & vbNewLine & "OS Version: Unknown "
    End Select
    fError.txtError = fError.txtError & CStr(osVerInfo.dwMajorVersion) & "." & CStr(osVerInfo.dwMinorVersion) & " " & osVerInfo.szCSDVersion
    
    'Show error dialog
    fError.Show 1, fMainForm
    'If fatal, end the program
    If isFatal Then End
End Sub

Sub SetStatusBar(ByVal newText As String)
    fMainForm.sbStatusBar.PanelText("main") = newText
    'fMainForm.sbStatusBar.Draw
End Sub
Sub SetStatusAutosave(ByVal isOn As Boolean, ByVal time As Date)
    If isOn = False Then
        fMainForm.sbStatusBar.PanelText("autosave") = "Off"
    Else
        fMainForm.sbStatusBar.PanelText("autosave") = Format(time, "nn:ss")
    End If
    'fMainForm.sbStatusBar.Draw
End Sub
Sub SetStatusXY(ByVal xPos As Long, ByVal yPos As Long)
    fMainForm.sbStatusBar.PanelText("cords") = "(" & CStr(xPos) & ", " & CStr(yPos) & ")"
    'fMainForm.sbStatusBar.Draw
End Sub

Sub SetupUI()
'Sets up the user interface
'Set up the imagelist
mainIml.OwnerHDC = fMainForm.picStatus.hDC
mainIml.ColourDepth = mainIml.SystemColourDepth
mainIml.IconSizeX = 16
mainIml.IconSizeY = 16
mainIml.Create
Dim hImg As Picture
Set hImg = LoadResPicture(101, vbResBitmap)
mainIml.AddFromHandle hImg.handle, IMAGE_BITMAP, , &HFF00FF
Set hImg = LoadResPicture(90, vbResIcon)
mainIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = Nothing

With fMainForm
    'Set up the statusbar
    .sbStatusBar.Create .picStatus
    .sbStatusBar.ImageList = mainIml.hIml
    .sbStatusBar.SizeGrip = True
    .sbStatusBar.AddPanel , , , , True, , , "main"
    .sbStatusBar.AddPanel , , , , , True, , "autosave"
    .sbStatusBar.AddPanel , , , , , True, , "cords"
    .picStatus.Height = .sbStatusBar.Height
    SetStatusBar "Ready"
    
    'Set up the menu
    .cMenu.ImageList = mainIml.hIml
    .cMenu.SubClassMenu fMainForm
    .cMenu.ItemIcon("mnuFileNew") = 11
    .cMenu.ItemIcon("mnuFileOpen") = 12
    .cMenu.ItemIcon("mnuFileSave") = 13
    .cMenu.ItemIcon("mnuFileProperties") = 15
    .cMenu.ItemIcon("mnuFilePrint") = 19
    .cMenu.ItemIcon("mnuEditUndo") = 8
    '.cMenu.ItemIcon("mnuEditCut") = 5
    .cMenu.ItemIcon("mnuEditCopy") = 6
    .cMenu.ItemIcon("mnuEditPaste") = 7
    .cMenu.ItemIcon("mnuViewRefresh") = 18
    .cMenu.ItemIcon("mnuViewStartOP2") = 47
    .cMenu.ItemIcon("mnuProjFile") = 44
    .cMenu.ItemIcon("mnuProjRem") = 10
    .cMenu.ItemIcon("mnuPluginsConfig") = 25
    
    
    'Set up the main toolbar
    .cTbMain.ImageSource = CTBExternalImageList
    .cTbMain.SetImageList mainIml.hIml
    .cTbMain.CreateToolbar 16, , , True
    .cTbMain.AddButton "New", 11, , , , CTBNormal, "new"
    .cTbMain.AddButton "Open", 12, , , , CTBNormal, "open"
    .cTbMain.AddButton "Save", 13, , , , CTBNormal, "save"
    .cTbMain.AddButton , , , , , CTBSeparator
    '.cTbMain.AddButton "Print", 19, , , , CTBNormal, "print"
    '.cTbMain.AddButton , , , , , CTBSeparator
    '.cTbMain.AddButton "Properties", 15, , , , CTBNormal, "props"
    '.cTbMain.AddButton , , , , , CTBSeparator
    .cTbMain.AddButton "Undo", 8, , , , CTBNormal, "undo"
    .cTbMain.AddButton , , , , , CTBSeparator
    '.cTbMain.AddButton "Cut", 5, , , , CTBNormal, "cut"
    .cTbMain.AddButton "Copy", 6, , , , CTBNormal, "copy"
    .cTbMain.AddButton "Paste", 7, , , , CTBNormal, "paste"
    .cTbMain.AddButton , , , , , CTBSeparator
    .cTbMain.AddButton "Start Outpost 2", 47, , , , CTBNormal, "op2"
    '.cTbMain.AddButton "Configure Plugins", 25, , , , CTBNormal, "plugconfig"
    
    'Set up the project toolbar
    '.cTbProj.ImageSource = CTBExternalImageList
    '.cTbProj.SetImageList mainIml.hIml
    '.cTbProj.CreateToolbar 16, , , True
    '.cTbProj.AddButton "Add New File", 41, , , , CTBDropDownArrow, "addnew"
    '.cTbProj.AddButton "Add Existing File", 44, , , , CTBDropDownArrow, "addexist"
    '.cTbProj.AddButton , , , , , CTBSeparator
    '.cTbProj.AddButton "Delete Selected File", 10, , , , CTBNormal, "delsel"
    
    .cReBar.CreateRebar .picToolbar.hWnd

    .cReBar.AddBandByHwnd .cTbMain.hWnd, , False, , "MainBar"
    '.cReBar.AddBandByHwnd .cTbProj.hwnd, , False, , "ProjBar"
    '.cReBar.BandMinimise 0
    .picToolbar.Height = .picToolbar.ScaleY(.cReBar.RebarHeight, vbPixels, vbTwips)
End With
End Sub

Public Sub SaveSettingIni(ByVal Section As String, ByVal Key As String, ByVal Setting As String)
WritePrivateProfileString Section, Key, Setting, App.Path & INIFILENAME
End Sub

Public Function GetSettingIni(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
Dim tStr As String
Dim rc As Long
tStr = String$(260, 0)
rc = GetPrivateProfileString(Section, Key, Default, tStr, 260, App.Path & INIFILENAME)
GetSettingIni = Left(tStr, rc)
End Function

Public Function GetSettingIniFromFile(ByVal Section As String, ByVal Key As String, ByVal iniName As String, Optional ByVal Default As String = "") As String
Dim tStr As String
Dim rc As Long
tStr = String$(260, 0)
rc = GetPrivateProfileString(Section, Key, Default, tStr, 260, iniName)
GetSettingIniFromFile = Left(tStr, rc)
End Function

Public Sub DeleteSettingIni(ByVal Section As String, Optional ByVal Key As String = vbNullString)
WritePrivateProfileString Section, Key, ByVal 0&, App.Path & INIFILENAME
End Sub

Public Function BrowseForFolder(ByVal hWnd As Long, ByVal dlgText As String)
BrowseForFolder = ""
Dim BI As BROWSEINFO
Dim nFolder As Long
Dim IDL As ITEMIDLIST
Dim pIdl As Long
Dim sPath As String
Dim SHFI As SHFILEINFO
With BI
    'The dialog's owner window...
    .hOwner = hWnd
    'Initialize the buffer that rtns the display name of the selected folder
    .pszDisplayName = String$(260, 0)
    'Set the dialog's banner text
    .lpszTitle = dlgText
    'Set the type of folders to display & return
    .ulFlags = BIF_RETURNONLYFSDIRS
End With
'Show the Browse dialog
pIdl = SHBrowseForFolder(BI)
'If the dialog was cancelled...
If pIdl = 0 Then Exit Function
'Fill sPath w/ the selected path from the id list
sPath = String$(260, 0)
SHGetPathFromIDList ByVal pIdl, ByVal sPath
'Display the path and the name of the selected folder
BrowseForFolder = Left(sPath, InStr(sPath, vbNullChar) - 1)
'Frees the memory
CoTaskMemFree pIdl
End Function

Public Sub ReadTilesetsCtl()
On Error GoTo useDefaults
'Read tilesets.ctl from the vol
Dim strmIn As SeekableStreamReader, strmOut As StreamWriter, f As Integer
Dim tmpStr As String, tmpStr2 As String

Set strmIn = ResMan.OpenStreamRead("tilesets.ctl")

'If it doesn't exist or can't be read, exit
If strmIn Is Nothing Then GoTo useDefaults

On Error GoTo oops

'Get a temp filename
tmpStr = String$(261, 0)
GetTempFileName App.Path, "opu", 0, tmpStr
tmpStr = Left$(tmpStr, lstrlen(tmpStr))

'Read it
tmpStr2 = String$(strmIn.StreamSize, 0)
strmIn.Read strmIn.StreamSize, StrPtr(tmpStr2)

'Write the string
Set strmOut = ResMan.OpenStreamWrite(tmpStr)
strmOut.Write strmIn.StreamSize, StrPtr(tmpStr2)
Set strmOut = Nothing
Set strmIn = Nothing
tmpStr2 = Empty

'Open the file with VB's native input mode
f = FreeFile
Open tmpStr For Input As #f
'Read them all in
Do Until EOF(f)
    ReDim Preserve defTilesets(numTilesets) As String
    Line Input #f, tmpStr2
    tmpStr2 = Trim$(tmpStr2)
    If tmpStr2 <> "" Then
        defTilesets(numTilesets) = tmpStr2
        numTilesets = numTilesets + 1
    End If
Loop

Close #f
Kill tmpStr
Exit Sub

oops:
GenerateError "Couldn't read tilesets.ctl, using default values", "ReadTilesetsCtl"
useDefaults:
'set defaults for tilesets
ReDim defTilesets(12) As String
defTilesets(0) = "well0000": defTilesets(1) = "well0001": defTilesets(2) = "well0002"
defTilesets(3) = "well0003": defTilesets(4) = "well0004": defTilesets(5) = "well0005"
defTilesets(6) = "well0006": defTilesets(7) = "well0007": defTilesets(8) = "well0008"
defTilesets(9) = "well0009": defTilesets(10) = "well0010": defTilesets(11) = "well0011"
defTilesets(12) = "well0012": numTilesets = 13
End Sub

Public Sub ReadTerrainsCtl()
On Error Resume Next
'Read terrains.ctl from the vol
Dim strmIn As SeekableStreamReader, strmOut As StreamWriter
Dim tmpStr As String, tmpStr2 As String, i As Long, j As Long, canKill As Boolean

canKill = True
Set strmIn = ResMan.OpenStreamRead("terrains.ctl")

'If it doesn't exist or can't be read, see if it exists in the mapper dir
If strmIn Is Nothing Then
    If FileLen(App.Path & "\terrains.ctl") > 0 Then
        canKill = False
        tmpStr = App.Path & "\terrains.ctl"
        On Error GoTo oops
    Else
        GoTo oops
    End If
Else
    On Error GoTo oops

    'Get a temp filename
    tmpStr = String$(261, 0)
    GetTempFileName App.Path, "opu", 0, tmpStr
    tmpStr = Left$(tmpStr, lstrlen(tmpStr))
    
    'Read it
    tmpStr2 = String$(strmIn.StreamSize, 0)
    strmIn.Read strmIn.StreamSize, StrPtr(tmpStr2)
    
    'Write the string
    Set strmOut = ResMan.OpenStreamWrite(tmpStr)
    strmOut.Write strmIn.StreamSize, StrPtr(tmpStr2)
    Set strmOut = Nothing
    Set strmIn = Nothing
End If

'Read the INI entries
tmpStr2 = GetSettingIniFromFile("terrains.ctl", "NumTerrains", tmpStr, "0")
numTerrains = CLng(tmpStr2)
If numTerrains < 1 Then GoTo oops
ReDim defTerrains(numTerrains - 1) As TerrainData

For i = 0 To numTerrains - 1
    defTerrains(i).startTile = CLng(GetSettingIniFromFile(CStr(i), "StartTile", tmpStr, "0"))
    defTerrains(i).endTile = CLng(GetSettingIniFromFile(CStr(i), "EndTile", tmpStr, "0"))
    defTerrains(i).dozed = CLng(GetSettingIniFromFile(CStr(i), "Dozed", tmpStr, "0"))
    defTerrains(i).rubble = CLng(GetSettingIniFromFile(CStr(i), "Rubble", tmpStr, "0"))
    For j = 0 To 5
        defTerrains(i).tubeUnk(j) = CLng(GetSettingIniFromFile(CStr(i), "TubeUnk" & CStr(j), tmpStr, "0"))
    Next
    For j = 0 To 15
        defTerrains(i).lavaWall(j) = CLng(GetSettingIniFromFile(CStr(i), "LavaWall" & CStr(j), tmpStr, "0"))
    Next
    For j = 0 To 15
        defTerrains(i).microbeWall(j) = CLng(GetSettingIniFromFile(CStr(i), "MicrobeWall" & CStr(j), tmpStr, "0"))
    Next
    For j = 0 To 15
        defTerrains(i).normalWall(j) = CLng(GetSettingIniFromFile(CStr(i), "NormalWall" & CStr(j), tmpStr, "0"))
    Next
    For j = 0 To 15
        defTerrains(i).damagedWall(j) = CLng(GetSettingIniFromFile(CStr(i), "DamagedWall" & CStr(j), tmpStr, "0"))
    Next
    For j = 0 To 15
        defTerrains(i).ruinedWall(j) = CLng(GetSettingIniFromFile(CStr(i), "RuinedWall" & CStr(j), tmpStr, "0"))
    Next
    defTerrains(i).lava = CLng(GetSettingIniFromFile(CStr(i), "Lava", tmpStr, "0"))
    defTerrains(i).flat1 = CLng(GetSettingIniFromFile(CStr(i), "Flat1", tmpStr, "0"))
    defTerrains(i).flat2 = CLng(GetSettingIniFromFile(CStr(i), "Flat2", tmpStr, "0"))
    defTerrains(i).flat3 = CLng(GetSettingIniFromFile(CStr(i), "Flat3", tmpStr, "0"))
    For j = 0 To 15
        defTerrains(i).tube(j) = CLng(GetSettingIniFromFile(CStr(i), "Tube" & CStr(j), tmpStr, "0"))
    Next
    defTerrains(i).scorched = CLng(GetSettingIniFromFile(CStr(i), "Scorched", tmpStr, "0"))
    For j = 0 To 20
        defTerrains(i).unkTile(j) = CLng(GetSettingIniFromFile(CStr(i), "Unknown" & CStr(j), tmpStr, "0"))
    Next
Next

If canKill Then Kill tmpStr
Exit Sub

oops:
GenerateError "Couldn't read terrains.ctl", "ReadTerrainsCtl"
End Sub


Public Sub ReadPrt()
On Error GoTo noGfx
1

'Read op2_art.bmp
If PRTFile.OpenBitmapFile(ResMan.RootPath & "\op2_art.bmp") = -1 Then GoTo noGfx
2

'Read op2_art.prt from the vol
'Dim strmIn As SeekableStreamReader
Dim strmIn As StreamReader
'Dim strmOut As StreamWriter, f As Integer
'Dim tmpStr As String, tmpStr2 As String

3
Set strmIn = ResMan.OpenStreamRead("op2_art.prt")

4
'If it doesn't exist or can't be read, exit
If strmIn Is Nothing Then GoTo noGfx

5
PRTFile.OpenPrtFile strmIn
Set strmIn = Nothing

'5
''Get a temp filename
'tmpStr = String$(261, 0)
'GetTempFileName App.Path, "opu", 0, tmpStr
'tmpStr = Left$(tmpStr, lstrlen(tmpStr))
'
'6
''Read it
'tmpStr2 = String$(strmIn.StreamSize, 0)
'strmIn.Read strmIn.StreamSize, StrPtr(tmpStr2)
'
'7
''Write the string
'Set strmOut = ResMan.OpenStreamWrite(tmpStr)
'strmOut.Write strmIn.StreamSize, StrPtr(tmpStr2)
'Set strmOut = Nothing
'Set strmIn = Nothing
'
'8
''Read PRT now
'PRTFile.OpenPrtFile tmpStr
''Delete the tempfile
'Kill tmpStr

9
'Now read color.bmp from the vol
Set strmIn = ResMan.OpenStreamRead("color.bmp")

10
'If it doesn't exist or can't be read, exit
If strmIn Is Nothing Then GoTo noGfx

11
PRTFile.OpenColorFile strmIn
Set strmIn = Nothing

'11
''Get a temp filename
'tmpStr = String$(261, 0)
'GetTempFileName App.Path, "opu", 0, tmpStr
'tmpStr = Left$(tmpStr, lstrlen(tmpStr))
'
'12
''Read it
'tmpStr2 = String$(strmIn.StreamSize, 0)
'strmIn.Read strmIn.StreamSize, StrPtr(tmpStr2)
'
'13
''Write the string
'Set strmOut = ResMan.OpenStreamWrite(tmpStr)
'strmOut.Write strmIn.StreamSize, StrPtr(tmpStr2)
'Set strmOut = Nothing
'Set strmIn = Nothing
'tmpStr2 = Empty
'
'14
''Read color.bmp now
'PRTFile.OpenColorFile tmpStr
''Delete the tempfile
'Kill tmpStr

15
Exit Sub

noGfx:
GenerateError "Couldn't read OP2_ART files, no unit graphics will be loaded.", "ReadPrt"
End Sub

Function UnsignedToLong(Value As Double) As Long
If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
If Value <= MAXINT_4 Then
    UnsignedToLong = Value
Else
    UnsignedToLong = Value - OFFSET_4
End If
End Function

Function LongToUnsigned(Value As Long) As Double
If Value < 0 Then
    LongToUnsigned = Value + OFFSET_4
Else
    LongToUnsigned = Value
End If
End Function

Function UnsignedToInteger(Value As Long) As Integer
If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
If Value <= MAXINT_2 Then
    UnsignedToInteger = Value
Else
    UnsignedToInteger = Value - OFFSET_2
End If
End Function

Function IntegerToUnsigned(Value As Integer) As Long
If Value < 0 Then
    IntegerToUnsigned = Value + OFFSET_2
Else
    IntegerToUnsigned = Value
End If
End Function

'Public Function LoadPlugin(ByVal strName As String, ByVal autoLoaded As Boolean) As Long
'On Error GoTo Grr
''Get a ClassFactory and obtain a pointer to an IPluginBase interface
'' first, if it fails then there's no use in allocating space
'Dim plugBase As IPluginBase
'Dim classFac As IClassFactory
'Dim clsid As IClassFactory_TLB.Guid
'Dim hDll As Long
'Dim appObj As MapperApp2
''try to obtain the clsid
'If GetClsidFromFile(strName, clsid) = False Then LoadPlugin = -1: Exit Function
''get a class factory
'hDll = CoGetClassObjectFromFile(clsid, strName, classFac)
'If hDll = 0 Then LoadPlugin = -1: Exit Function
''get a pointer to the actual IPluginBase
'
'classFac.CreateInstance ByVal 0&, IID_IPluginBase, plugBase
''class factory no longer needed, destroy it
'classFac.Release
''create an app object and call the onLoad procedure
'Set appObj = New MapperApp2
'If plugBase.onLoad(IIf(autoLoaded, PLR_AUTOSTART, PLR_REQUESTED), appObj) = PLE_OK Then
'    'Success - Set up the vars
'    ReDim Preserve loadedPlugins(numLoadedPlugins)
'    loadedPlugins(numLoadedPlugins).strFilename = strName
'    loadedPlugins(numLoadedPlugins).dllHandle = hDll
'    Set loadedPlugins(numLoadedPlugins).plugObject = plugBase
'    Set loadedPlugins(numLoadedPlugins).appObject = appObj
'    loadedPlugins(numLoadedPlugins).appObject.pluginId = numLoadedPlugins
'    loadedPlugins(numLoadedPlugins).numMenuItems = 0
'    loadedPlugins(numLoadedPlugins).numEventHooks = 0
'    'Increment loadedPlugins counter
'    LoadPlugin = numLoadedPlugins
'    numLoadedPlugins = numLoadedPlugins + 1
'    Exit Function
'End If
'Grr:
'LoadPlugin = -1
'FreeLibrary hLib
'GenerateError "Error while loading plugin", "<PluginLoader>::LoadPlugin"
'End Function
'Public Sub UnloadPlugin(ByVal id As Long, ByVal appExit As Boolean)
'On Error GoTo Grr
''First, make sure the plugin's unload event gets called
'loadedPlugins(id).plugObject.onUnload IIf(appExit, PUR_APPEXIT, PUR_REQUESTED)
''Destroy the plugin object and app object
'Set loadedPlugins(id).plugObject = Nothing
'Set loadedPlugins(id).appObject = Nothing
''Unload the plugin dll
'FreeLibrary loadedPlugins(id).dllHandle
''Remove it from the array
'PluginCtxRemoveItem loadedPlugins, id
'numLoadedPlugins = numLoadedPlugins - 1
'Exit Sub
'Grr:
'GenerateError "Error while unloading plugin", "<PluginLoader>::UnloadPlugin"
'End Sub
'Public Sub ConvGuid(ByVal guidStr As String, outGuid As IClassFactory_TLB.Guid)
'On Error GoTo zeroOut
''Converts a GUID from form {XXXXXX ..  to a real guid struct
''First parse off the { } - tokens
'Dim buildStr As String
'buildStr = Trim$(UCase$(guidStr))
'buildStr = Replace$(buildStr, "{", "", , , vbTextCompare)
'buildStr = Replace$(buildStr, "}", "", , , vbTextCompare)
'buildStr = Replace$(buildStr, "-", "", , , vbTextCompare)
''Now just a hex string. Convert to GUID format
'outGuid.Data1 = CLng("&H" & Left$(buildStr, 8))
'outGuid.Data2 = CInt("&H" & Mid$(buildStr, 9, 4))
'outGuid.Data3 = CInt("&H" & Mid$(buildStr, 13, 4))
'outGuid.Data4(0) = CByte("&H" & Mid$(buildStr, 17, 2))
'outGuid.Data4(1) = CByte("&H" & Mid$(buildStr, 19, 2))
'outGuid.Data4(2) = CByte("&H" & Mid$(buildStr, 21, 2))
'outGuid.Data4(3) = CByte("&H" & Mid$(buildStr, 23, 2))
'outGuid.Data4(4) = CByte("&H" & Mid$(buildStr, 25, 2))
'outGuid.Data4(5) = CByte("&H" & Mid$(buildStr, 27, 2))
'outGuid.Data4(6) = CByte("&H" & Mid$(buildStr, 29, 2))
'outGuid.Data4(7) = CByte("&H" & Mid$(buildStr, 31, 2))
'Exit Sub
'zeroOut: 'error so zero it
'outGuid.Data1 = 0
'outGuid.Data2 = 0
'outGuid.Data3 = 0
'outGuid.Data4(0) = 0
'outGuid.Data4(1) = 0
'outGuid.Data4(2) = 0
'outGuid.Data4(3) = 0
'outGuid.Data4(4) = 0
'outGuid.Data4(5) = 0
'outGuid.Data4(6) = 0
'outGuid.Data4(7) = 0
'End Sub
'Public Function GetClsidFromFile(ByVal Filename As String, clsid As IClassFactory_TLB.Guid) As Boolean
'On Error GoTo getOut
''Uses the TLBINF32 library to figure out the clsid mapper plugin coclass
''Returns true if it succeeded, false if it didn't
'Dim tlbInfo As TypeLibInfo
'Set tlbInfo = TypeLibInfoFromFile(Filename)
'Dim clsInfo As CoClassInfo
'Set clsInfo = tlbInfo.CoClasses.NamedItem("CMapperPlugin")
'ConvGuid clsInfo.Guid, clsid
'GetClsidFromFile = True
'Exit Function
'getOut:
'Set clsInfo = Nothing
'Set tlbInfo = Nothing
'GetClsidFromFile = False
'GenerateError "Couldn't determine correct GUIDs for plugin.", "<PluginLoader>::GetClsidFromFile"
'End Function
'
'Public Sub PluginCtxRemoveItem(ItemArray() As PluginCtx, ByVal ItemElement As Long)
'
''PURPOSE:       Remove an item from an array, then
''               resize the array
'
''PARAMETERS:    ItemArray: Array, passed by reference, with
''               item to be removed.  Array must not be fixed
'
''               ItemElement: Element to Remove
'
''Modified to work on plugin array
'
'Dim lCtr As Long
'Dim lTop As Long
'Dim lBottom As Long
'
'lTop = UBound(ItemArray)
'lBottom = LBound(ItemArray)
'
'If ItemElement < lBottom Or ItemElement > lTop Then
'    Err.Raise 9, , "Subscript out of Range"
'    Exit Sub
'End If
'
'For lCtr = ItemElement To lTop - 1
'    ItemArray(lCtr) = ItemArray(lCtr + 1)
'    ItemArray(lCtr).appObject.pluginId = lCtr
'Next
'ReDim Preserve ItemArray(lBottom To lTop - 1)
'
'End Sub
'
'Public Sub RefreshPluginMenu()
'
'End Sub
'Public Sub ProcessPluginMenu(ByVal strName As String)
'
'End Sub
