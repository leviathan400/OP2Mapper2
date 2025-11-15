VERSION 5.00
Begin VB.Form frmMinimap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mini Map"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   ControlBox      =   0   'False
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
   Icon            =   "frmMinimap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picMinimap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   167
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   231
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Label lblRedraw 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drawing..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
      Begin VB.Shape shpCur 
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMinimap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private thisMap As MapManager
Private mouseIsDown As Boolean

Public ownerForm As MapManager
Private uRec As UnitRec


Public Sub SetNewMap(map As MapManager)
On Error GoTo augh
Set thisMap = map

'Calculate the form's client area properly
Dim fWidth As Single, fHeight As Single
fWidth = map.myMap.TileWidth * Screen.TwipsPerPixelX 'let the system tell you how many twips are in a pixel
fHeight = map.myMap.TileHeight * Screen.TwipsPerPixelY
Me.Width = fWidth + (Me.Width - Me.ScaleWidth)
Me.Height = fHeight + (Me.Height - Me.ScaleHeight)
'Resize the picbox
picMinimap.Width = Me.ScaleWidth
picMinimap.Height = Me.ScaleHeight

lblRedraw.Visible = True
picMinimap.Cls
Dim i As Long, j As Long, theHdc As Long
theHdc = picMinimap.hDC
For i = 0 To map.myMap.TileWidth - 1
    For j = 0 To map.myMap.TileHeight - 1
        'Populate the picbox with pixels
        'Lookup and set the color on this (i,j) position
        SetPixelV theHdc, i, j, map.myMap.TileSetManager.TileSet(map.myMap.TileSetManager.tileSetIndex( _
            map.myMap.mappingIndex(i, j))).MiniMapColors( _
            map.myMap.TileSetManager.TileSetTileIndex(map.myMap.mappingIndex(i, j)))
    Next
    DoEvents
    'picMinimap.Refresh
Next

If map.toolMode = PlaceUnit Then
    'Check for units and draw them
    For i = 0 To map.numUnitRecs - 1
        'Access it
        thisMap.UnitRecord i, VarPtr(uRec)
        If uRec.uType.isGaia = True Then
            'Gaia units always have white
            SetPixelV picMinimap.hDC, uRec.locX, uRec.locY, vbWhite
        Else
            SetPixelV picMinimap.hDC, uRec.locX, uRec.locY, PRTFile.GetPlayerColor(uRec.playerNum)
        End If
    Next
End If

lblRedraw.Visible = False
picMinimap.Refresh
augh:
End Sub

Public Sub SetNewName(ByVal strName As String)
Me.Caption = "Mini Map - " & strName
End Sub

Private Sub Form_Deactivate()
If ownerForm Is Nothing Then Exit Sub
ownerForm.ToolDeactivated
End Sub

Private Sub Form_Load()
'Set window settings
Me.Left = GetSettingIni("Window", "MinimapLeft", 0)
Me.TOp = GetSettingIni("Window", "MinimapTop", 0)
End Sub

Private Sub Form_Unload(cancel As Integer)
'Save window settings
If Me.WindowState <> vbMinimized Then
    SaveSettingIni "Window", "MinimapLeft", Me.Left
    SaveSettingIni "Window", "MinimapTop", Me.TOp
End If
'Destroy mapobject
Set thisMap = Nothing
End Sub

Private Sub picMinimap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Me.SetFocus: Me.PopupMenu fMainForm.mnuMinimap: Exit Sub
mouseIsDown = True
shpCur.BorderColor = vbCyan
'Make sure it cant go beyond the constraints of the picbox
shpCur.Left = x - (shpCur.Width \ 2)
shpCur.TOp = y - (shpCur.Height \ 2)
If x < (shpCur.Width \ 2) Then shpCur.Left = 0
If y < (shpCur.Height \ 2) Then shpCur.TOp = 0
If (x > picMinimap.ScaleWidth - (shpCur.Width \ 2) + 1) Then shpCur.Left = picMinimap.ScaleWidth - shpCur.Width + 1
If (y > picMinimap.ScaleHeight - (shpCur.Height \ 2) + 1) Then shpCur.TOp = picMinimap.ScaleHeight - shpCur.Height + 1
thisMap.SetScroll shpCur.Left, shpCur.TOp
End Sub

Private Sub picMinimap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If mouseIsDown = True Then picMinimap_MouseDown Button, Shift, x, y
End Sub

Private Sub picMinimap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
mouseIsDown = False
shpCur.BorderColor = vbWhite
End Sub

Private Sub picMinimap_Resize()
lblRedraw.Width = picMinimap.ScaleWidth
lblRedraw.TOp = (picMinimap.ScaleHeight / 2) - lblRedraw.Height
End Sub

Public Sub SetExtents(ByVal horExtent As Long, ByVal verExtent As Long)
'Resize the selectorbox
shpCur.Width = horExtent
shpCur.Height = verExtent
End Sub

Public Sub SetLoc(ByVal x As Long, ByVal y As Long)
shpCur.Left = x
shpCur.TOp = y
End Sub

Public Sub UpdatePixel(ByVal x As Long, ByVal y As Long)
On Error GoTo getOut
Dim i As Long, theHdc As Long
theHdc = picMinimap.hDC
'Update a pixel on the minimap
'Lookup and set the color on this (x,y) position
SetPixelV theHdc, x, y, thisMap.myMap.TileSetManager.TileSet(thisMap.myMap.TileSetManager.tileSetIndex( _
    thisMap.myMap.mappingIndex(x, y))).MiniMapColors( _
    thisMap.myMap.TileSetManager.TileSetTileIndex(thisMap.myMap.mappingIndex(x, y)))

If thisMap.toolMode = PlaceUnit Then
    'Check for units and draw them
    For i = 0 To thisMap.numUnitRecs - 1
        'Access it
        thisMap.UnitRecord i, VarPtr(uRec)
        If uRec.locX = x And uRec.locY = y Then
            If uRec.uType.isGaia = True Then
                'Gaia units always have white
                SetPixelV picMinimap.hDC, uRec.locX, uRec.locY, vbWhite
            Else
                SetPixelV picMinimap.hDC, uRec.locX, uRec.locY, PRTFile.GetPlayerColor(uRec.playerNum)
            End If
        End If
    Next
End If

picMinimap.Refresh
getOut:
End Sub
