VERSION 5.00
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Begin VB.Form TilesetEditMgr 
   Caption         =   "Tileset"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
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
   Icon            =   "TilesetEditMgr.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin VB.VScrollBar vsbScroll 
      Height          =   2655
      LargeChange     =   6
      Left            =   5640
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picTiles 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.Shape shpCursor 
         BorderColor     =   &H00FFFFFF&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
   Begin vbalTBar6.cReBar myRebar 
      Left            =   2760
      Top             =   0
      _ExtentX        =   3201
      _ExtentY        =   661
   End
   Begin vbalTBar6.cToolbar myToolbar 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
   End
End
Attribute VB_Name = "TilesetEditMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myName As String, myTitle As String, isFile As Boolean
Private myIml As New cVBALImageList
Private myTabId As Long


Private thisSet As TileSet

Private palWidth As Long, palHeight As Long, dcExists As Boolean
Private oldScrollVal As Long

Public tileID As Long

Private Sub Form_Load()
'Set up the imagelist
myIml.OwnerHDC = fMainForm.picStatus.hDC
myIml.ColourDepth = myIml.SystemColourDepth
myIml.IconSizeX = 16
myIml.IconSizeY = 16
myIml.Create
Dim hImg As Picture
Set hImg = LoadResPicture(103, vbResBitmap)
myIml.AddFromHandle hImg.handle, IMAGE_BITMAP, , &HFF00FF
Set hImg = Nothing

'Set up the toolbar
myToolbar.ImageSource = CTBExternalImageList
myToolbar.SetImageList myIml.hIml
myToolbar.CreateToolbar 16, , , True
myToolbar.AddButton "Add Tile(s)", 0, , , , CTBDropDownArrow, "addtile"
myToolbar.AddButton "Delete Tile", 1, , , , CTBNormal, "deltile"
myToolbar.AddButton , , , , , CTBSeparator
myToolbar.AddButton "Add Tiles from Bitmap", 2, , , , CTBNormal, "addbmp"
myToolbar.AddButton "Add Blank Tile", 3, , , , CTBNormal, "addplace"
myToolbar.AddButton , , , , , CTBSeparator
myToolbar.AddButton "Edit in External Editor", 4, , , , CTBNormal, "editext"
myToolbar.AddButton "Export Tile(s)", 5, , , , CTBDropDownArrow, "exporttile"
myToolbar.AddButton , , , , , CTBSeparator
myToolbar.AddButton "Edit Palette", 6, , , , CTBNormal, "editpal"
'myToolbar.AddButton "Save as PBMP", 7, , , , CTBCheck, "savepbmp"

'Link the toolbar to the rebar
myRebar.CreateRebar Me.hWnd
myRebar.AddBandByHwnd myToolbar.hWnd, , False, , "ToolsBar"

'Add this form to the tab bar
fMainForm.DTabControl.Tabs.Add "BMP" & CStr(fMainForm.DTabControl.Tabs.Count), , Me.Caption
myTabId = fMainForm.DTabControl.Tabs.Count - 1
Me.tag = "BMP" & CStr(myTabId)

End Sub

Public Function LoadTileset(ByVal Filename As String) As Boolean
On Error GoTo oops
Set thisSet = ResMan.LoadTileSetFile(Filename)
RedrawSelf

Me.ZOrder 0

isFile = True
LoadTileset = True

Exit Function
oops:
LoadTileset = False
GenerateError "Failed to load tileset. Perhaps corrupt BMP or PBMP format?", "TilesetEditMgr::LoadTileset"
End Function

Public Sub SetNewName(ByVal strName As String, ByVal sTitle As String)
On Error GoTo baha
'myName = long name as in full path to the file
'myTitle = short name, just the filetitle or other
myName = strName
myTitle = sTitle
Me.Caption = myTitle

'Set it on the tab bar
fMainForm.DTabControl.Tabs.Item("BMP" & CStr(myTabId)).Caption = Me.Caption

Exit Sub
baha:
GenerateError "SetNewName failed", "TilesetEditMgr::SetNewName"
End Sub

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
    'thisSet.SaveMap stream, TileGroups
    Set stream = Nothing
    SaveSelf = True
End If
End Function

Public Function SaveSelfAs() As Boolean
'Save the map file
Dim sFile As String
sFile = myName
If fMainForm.cCMDlg.VBGetSaveFileName(sFile, myTitle, True, "Tileset Bitmaps (*.bmp)|*.bmp", , mapsDir, "Save Tileset", "bmp") = False Then Exit Function
myName = sFile
Dim stream As StreamWriter
Set stream = ResMan.OpenStreamWrite(sFile)
'thisSet.SaveMap stream, TileGroups
Set stream = Nothing
SaveSelfAs = True
SetNewName myName, myTitle
End Function

Public Function CreateTileset() As Boolean

End Function

Public Sub RedrawSelf()
Dim i As Integer

palWidth = Int(picTiles.ScaleWidth / 32)

'Determine tile palette height (numrows*32)
If thisSet.NumTiles < palWidth Then
    palHeight = 32
Else
    palHeight = (Int(thisSet.NumTiles / palWidth) + 1) * 32
End If

picTiles.Cls

Dim curX As Long, curY As Long
'Draw the tiles into the DC
curX = 0: curY = 0
'For i = 0 To mgr.NumTileSets - 1
For i = vsbScroll.Value * palWidth To thisSet.NumTiles - 1
    thisSet.PasteTile picTiles.hDC, curX, curY, i
    curX = curX + 32
    If curX > picTiles.ScaleWidth - 32 Then curX = 0: curY = curY + 32
Next

'Resize the scroll bars
If Int(thisSet.NumTiles / palWidth) < Int(picTiles.ScaleHeight / 32) + 1 Then
    vsbScroll.Max = 0
Else
    vsbScroll.Max = (palHeight / 32) - Int(picTiles.ScaleHeight / 32)
End If
End Sub

Public Sub AddOne()

End Sub

Public Sub AddDir()

End Sub

Public Sub ExportOne()

End Sub

Public Sub ExportAll()

End Sub

Public Sub DeleteTile(ByVal id As Long)
Dim scanWidth As Long, buffer As String, nTilesLeft As Long, i As Long
Dim totalSize As Long
'Remove a tile from the tileset

'Move all tiles behind the tile back one space in memory

thisSet.SetNumTiles thisSet.NumTiles
RedrawSelf

If id > 0 Then
    'Figure out the byte width of one scanline, and the total size of the bitmap
    scanWidth = ((32 * thisSet.BitDepth + 31) And Not 31) \ 8
    totalSize = (scanWidth * 32 * thisSet.NumTiles)
    
    'Allocate buffer
    buffer = String$(scanWidth * 32, 0)
    
    For i = 1 To id
        'Get data, move it one "tile space" back
        thisSet.GetPixelData StrPtr(buffer), totalSize - (scanWidth * 32 * i), scanWidth * 32
        thisSet.SetPixelData StrPtr(buffer), totalSize - (scanWidth * 32 * (i + 1)), scanWidth * 32
        RedrawSelf
    Next
End If
'Contract the internal buffer

thisSet.SetNumTiles thisSet.NumTiles - 1
If thisSet.NumTiles = 0 Then 'move the cursor to 0,0 and pick tile 0
    tileID = 0
    shpCursor.Move 0, 0, 32, 32
End If
RedrawSelf
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
Select Case MsgBox("Do you want to save changes to " & myName & "?", vbQuestion Or vbYesNoCancel, "Save Changes?")
    Case vbYes
        If SaveSelf = False Then cancel = True: Exit Sub
    Case vbCancel
        cancel = True: Exit Sub
End Select
myRebar.RemoveAllRebarBands
End Sub

Private Sub Form_Resize()
On Error Resume Next
picTiles.Move 0, myRebar.RebarHeight, Me.ScaleWidth - vsbScroll.Width, Me.ScaleHeight - myRebar.RebarHeight
vsbScroll.TOp = myRebar.RebarHeight
vsbScroll.Left = picTiles.Width
vsbScroll.Height = picTiles.Height
'Resize the rebar
myRebar.RebarSize
'Redraw (this handles the scrollbar ranging automatically)
RedrawSelf
End Sub

Private Sub Form_Unload(cancel As Integer)
'Remove this from the tab bar
fMainForm.DTabControl.Tabs.Remove "BMP" & CStr(myTabId)
End Sub

Private Sub myToolbar_ButtonClick(ByVal lButton As Long)
Select Case myToolbar.ButtonKey(lButton)
    Case "deltile"
        DeleteTile tileID
    Case "addbmp"
    
    Case "addplace"
    
    Case "editext"
    
    Case "editpal"

End Select
End Sub

Private Sub myToolbar_DropDownPress(ByVal lButton As Long)
'Handle the dropdown menus
Select Case myToolbar.ButtonKey(lButton)
    Case "addtile"
        Me.SetFocus
        Me.PopupMenu fMainForm.mnuTilesetEditAdd
    Case "exporttile"
        Me.SetFocus
        Me.PopupMenu fMainForm.mnuTilesetEditExp
End Select
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim intX As Long, intY As Long
'Get and validate the values
intX = x \ 32
intY = y \ 32 + vsbScroll.Value
If intX < 0 Or intX > palWidth - 1 Then Exit Sub
If intY < 0 Or intY > palHeight - 1 Then Exit Sub

'Add together values for tile ID
tileID = intY * palWidth + intX

'If this goes outside the tileset boundaries then it can't be selected
If tileID > thisSet.NumTiles - 1 Then Exit Sub

shpCursor.Move intX * 32, (intY - vsbScroll.Value) * 32, 32, 32
SetStatusBar "Selected TileID " & CStr(tileID)
End Sub

Private Sub vsbScroll_Change()
'Calculate the distance to move the selection cursor
If oldScrollVal <> vsbScroll.Value Then 'We moved, so move cursor
    shpCursor.Move shpCursor.Left, ((shpCursor.TOp / 32) - (vsbScroll.Value - oldScrollVal)) * 32, 32, 32
End If
'Save old scroll value
oldScrollVal = vsbScroll.Value
'Redraw
RedrawSelf
End Sub

Private Sub vsbScroll_Scroll()
vsbScroll_Change
End Sub

Public Sub SaveAutosave()
Beep: SetStatusBar "TODO: Implement SaveAutosave()!"
End Sub
