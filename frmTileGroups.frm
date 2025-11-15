VERSION 5.00
Begin VB.Form frmTileGroups 
   Caption         =   "Tile Groups"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
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
   Icon            =   "frmTileGroups.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   296
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   Begin VB.PictureBox picGroup 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   373
      TabIndex        =   3
      Top             =   360
      Width           =   5655
   End
   Begin VB.HScrollBar hsbScroll 
      Height          =   255
      LargeChange     =   6
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   5655
   End
   Begin VB.VScrollBar vsbScroll 
      Height          =   3615
      LargeChange     =   6
      Left            =   5760
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.ComboBox cboTileGroup 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmTileGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const BLACKNESS = &H42

Public intDc As Long, tmpBm As Long

Public thisMap As MapFile
Private isGroup As Boolean, dcExists As Boolean

Public ownerForm As MapManager

Public curTileGroup As TileGroup
Public curGroupW As Long
Public curGroupH As Long
Public prepareClose As Boolean

Private Sub Form_Deactivate()
If ownerForm Is Nothing Then Exit Sub
ownerForm.ToolDeactivated
End Sub

Private Sub Form_Load()
Me.Left = GetSettingIni("Window", "GroupsLeft", 0)
Me.TOp = GetSettingIni("Window", "GroupsTop", 0)
Me.Width = GetSettingIni("Window", "GroupsWidth", 6240)
Me.Height = GetSettingIni("Window", "GroupsHeight", 4845)
Me.WindowState = GetSettingIni("Window", "GroupsWindowState", 0)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If prepareClose = False Then Cancel = True
End Sub

'Redraw the tile groups based on the map file
Public Sub SetNewMap(map As MapFile)
Dim i As Integer

Set thisMap = map
isGroup = True

'Determine the number of groups in all
cboTileGroup.Enabled = True
cboTileGroup.Clear

If map.NumTileGroups = 0 Then
    'Need to notify user somehow that no tile groups exist
    isGroup = False
    cboTileGroup.Enabled = False
    cboTileGroup.AddItem "<No tile groups>"
    cboTileGroup.ListIndex = 0
    hsbScroll.Max = 0
    vsbScroll.Max = 0
    Exit Sub
End If

For i = 0 To map.NumTileGroups - 1
    cboTileGroup.AddItem map.TileGroupName(i)
Next

cboTileGroup.ListIndex = 0

'Switch to group zero
ChangeGroup 0
End Sub

Public Sub ChangeGroup(ByVal idx As Long)
If isGroup = False Then Exit Sub

'Prepare to draw the new bitmap
If dcExists Then
    DeleteDC intDc
    DeleteObject tmpBm
End If
dcExists = True

'Create a DC and bitmap compatible with the picture box
intDc = CreateCompatibleDC(picGroup.hdc)
tmpBm = CreateCompatibleBitmap(picGroup.hdc, thisMap.TileGroup(idx).TileWidth * 32, thisMap.TileGroup(idx).TileHeight * 32)
'Out with the old... in with the new
DeleteObject SelectObject(intDc, tmpBm)

'Color fill it to black
BitBlt intDc, 0, 0, thisMap.TileGroup(idx).TileWidth * 32, thisMap.TileGroup(idx).TileHeight * 32, 0, 0, 0, BLACKNESS
'Draw the group
thisMap.TileGroup(idx).Draw intDc, 0, 0, thisMap.TileGroup(idx).TileWidth * 32, thisMap.TileGroup(idx).TileHeight * 32

Set curTileGroup = thisMap.TileGroup(idx)
curGroupW = curTileGroup.TileWidth
curGroupH = curTileGroup.TileHeight

RedrawSelf

'Resize the scrollbars
Form_Resize
End Sub

Private Sub cboTileGroup_Click()
'Show a different group
ChangeGroup cboTileGroup.ListIndex
End Sub

Private Sub Form_Resize()
On Error GoTo NeverMind
cboTileGroup.Width = Me.ScaleWidth
vsbScroll.Left = Me.ScaleWidth - vsbScroll.Width
hsbScroll.TOp = Me.ScaleHeight - hsbScroll.Height
picGroup.Width = vsbScroll.Left
picGroup.Height = hsbScroll.TOp - picGroup.TOp
vsbScroll.Height = picGroup.Height
hsbScroll.Width = picGroup.Width
'Resize scrollbars
If thisMap.TileGroup(cboTileGroup.ListIndex).TileWidth < Int(picGroup.ScaleWidth / 32) + 1 Then
    hsbScroll.Max = 0
Else
    hsbScroll.Max = thisMap.TileGroup(cboTileGroup.ListIndex).TileWidth - Int(picGroup.ScaleWidth / 32)
End If
If thisMap.TileGroup(cboTileGroup.ListIndex).TileHeight < Int(picGroup.ScaleHeight / 32) + 1 Then
    vsbScroll.Max = 0
Else
    vsbScroll.Max = thisMap.TileGroup(cboTileGroup.ListIndex).TileHeight - Int(picGroup.ScaleHeight / 32)
End If
NeverMind: 'trap 'invalid whatever' errors
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Save window state
If Me.WindowState <> vbMinimized Then
    SaveSettingIni "Window", "GroupsLeft", Me.Left
    SaveSettingIni "Window", "GroupsTop", Me.TOp
    SaveSettingIni "Window", "GroupsWidth", Me.Width
    SaveSettingIni "Window", "GroupsHeight", Me.Height
    SaveSettingIni "Window", "GroupsWindowState", Me.WindowState
End If
'Clean up bitmaps etc
Set thisMap = Nothing
If dcExists Then
    DeleteDC intDc
    DeleteObject tmpBm
End If
End Sub

Private Sub picGroup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

' ** TODO **
'If Button = vbRightButton Then Me.PopupMenu fMainForm.mnuGroup
End Sub

Private Sub vsbScroll_Change()
'Redraw
RedrawSelf
End Sub

Private Sub vsbScroll_Scroll()
vsbScroll_Change
End Sub

Private Sub hsbScroll_Change()
'Redraw
RedrawSelf
End Sub

Private Sub hsbScroll_Scroll()
hsbScroll_Change
End Sub

Public Sub RedrawSelf()
picGroup.Cls
BitBlt picGroup.hdc, 0, 0, picGroup.ScaleWidth, picGroup.ScaleHeight, intDc, hsbScroll.Value * 32, vsbScroll.Value * 32, vbSrcCopy
picGroup.Refresh
End Sub

Public Sub SetNewName(ByVal strName As String)
Me.Caption = "Tile Groups - " & strName
End Sub

Public Sub AddGroup(ByVal sName As String)
On Error GoTo oops
thisMap.AddTileGroup copyBuffer
thisMap.TileGroupName(thisMap.NumTileGroups - 1) = sName
SetNewMap thisMap
Exit Sub
oops:
GenerateError "AddTileGroup failed", "TileGroupTool::AddGroup"
End Sub
Public Sub DelCurGroup()
On Error GoTo oops
If thisMap.NumTileGroups = 0 Then Exit Sub
If MsgBox("Are you sure you want to permanently delete the selected tile group?", vbQuestion Or vbYesNo, "Delete Tile Group") = vbYes Then
    thisMap.RemoveTileGroup cboTileGroup.ListIndex
    SetNewMap thisMap
End If
Exit Sub
oops:
GenerateError "RemoveTileGroup failed", "TileGroupTool::DelCurGroup"
End Sub
