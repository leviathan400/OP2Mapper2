VERSION 5.00
Begin VB.Form frmCellType 
   Caption         =   "Cell Types"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2280
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
   Icon            =   "frmCellType.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   Begin VB.ListBox lstCellType 
      Height          =   3420
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmCellType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lookupTable(31) As CellTypes

Public ownerForm As MapManager

Public currentSel As CellTypes
Public prepareClose As Boolean

Private Sub Form_Deactivate()
If ownerForm Is Nothing Then Exit Sub
ownerForm.ToolDeactivated
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If prepareClose = False Then cancel = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
lstCellType.Width = Me.ScaleWidth
lstCellType.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(cancel As Integer)
If Me.WindowState <> vbMinimized Then
    SaveSettingIni "Window", "CellEditLeft", Me.Left
    SaveSettingIni "Window", "CellEditTop", Me.TOp
    SaveSettingIni "Window", "CellEditWidth", Me.Width
    SaveSettingIni "Window", "CellEditHeight", Me.Height
    SaveSettingIni "Window", "CellEditWindowState", Me.WindowState
End If
End Sub

Private Sub lstCellType_Click()
currentSel = lookupTable(lstCellType.ListIndex)
End Sub

Private Sub Form_Load()
'Set window settings
Me.Left = GetSettingIni("Window", "CellEditLeft", 0)
Me.TOp = GetSettingIni("Window", "CellEditTop", 0)
Me.Width = GetSettingIni("Window", "CellEditWidth", 2400)
Me.Height = GetSettingIni("Window", "CellEditHeight", 6800)
Me.WindowState = GetSettingIni("Window", "CellEditWindowState", 0)

'Populate the list and lookup table with celltypes
lstCellType.AddItem "Fast Passable 1 (F1)"
lookupTable(0) = FastPassable1
lstCellType.AddItem "Fast Passable 2 (F2)"
lookupTable(1) = FastPassable2
lstCellType.AddItem "Medium Passable 1 (M1)"
lookupTable(2) = MediumPassable1
lstCellType.AddItem "Medium Passable 2 (M2)"
lookupTable(3) = MediumPassable2
lstCellType.AddItem "Slow Passable 1 (S1)"
lookupTable(4) = SlowPassable1
lstCellType.AddItem "Slow Passable 2 (S2)"
lookupTable(5) = SlowPassable2
lstCellType.AddItem "Impassable 1 (I1)"
lookupTable(6) = Impassable1
lstCellType.AddItem "Impassable 2 (I2)"
lookupTable(7) = Impassable2
lstCellType.AddItem "North Cliffs (NC)"
lookupTable(8) = NorthCliffs
lstCellType.AddItem "Cliffs - High Side (CHS)"
lookupTable(9) = CliffsHighSide
lstCellType.AddItem "Cliffs - Low Side (CLS)"
lookupTable(10) = CliffsLowSide
lstCellType.AddItem "zPad 12 (Z12)"
lookupTable(11) = zPad12
lstCellType.AddItem "zPad 13 (Z13)"
lookupTable(12) = zPad13
lstCellType.AddItem "zPad 14 (Z14)"
lookupTable(13) = zPad14
lstCellType.AddItem "zPad 15 (Z15)"
lookupTable(14) = zPad15
lstCellType.AddItem "zPad 16 (Z16)"
lookupTable(15) = zPad16
lstCellType.AddItem "zPad 17 (Z17)"
lookupTable(16) = zPad17
lstCellType.AddItem "zPad 18 (Z18)"
lookupTable(17) = zPad18
lstCellType.AddItem "zPad 19 (Z19)"
lookupTable(18) = zPad19
lstCellType.AddItem "zPad 20 (Z20)"
lookupTable(19) = zPad20
lstCellType.AddItem "Bulldozed (D)"
lookupTable(20) = DozedArea
lstCellType.AddItem "Rubble (R)"
lookupTable(21) = rubble
lstCellType.AddItem "Fumaroles/Vents (V)"
lookupTable(22) = VentsAndFumaroles
lstCellType.AddItem "Normal Wall (NW)"
lookupTable(23) = normalWall
lstCellType.AddItem "Microbe Wall (MW)"
lookupTable(24) = microbeWall
lstCellType.AddItem "Lava Wall (LW)"
lookupTable(25) = lavaWall
lstCellType.AddItem "Tube 0 (T0)"
lookupTable(26) = Tube0
lstCellType.AddItem "Tube 1 (T1)"
lookupTable(27) = Tube1
lstCellType.AddItem "Tube 2 (T2)"
lookupTable(28) = Tube2
lstCellType.AddItem "Tube 3 (T3)"
lookupTable(29) = Tube3
lstCellType.AddItem "Tube 4 (T4)"
lookupTable(30) = Tube4
lstCellType.AddItem "Tube 5 (T5)"
lookupTable(31) = Tube5
lstCellType.ListIndex = 0
currentSel = lookupTable(0)
End Sub

Public Sub SetNewName(ByVal strName As String)
Me.Caption = "Cell Types - " & strName
End Sub


