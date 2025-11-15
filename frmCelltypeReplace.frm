VERSION 5.00
Begin VB.Form frmCelltypeReplace 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace Celltypes"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   Icon            =   "frmCelltypeReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboReplace 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Replace"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblWarn 
      Caption         =   "Warning: this operation is not undoable."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace with:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblSearch 
      Caption         =   "Search for:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmCelltypeReplace.frx":000C
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCelltypeReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lookupTable(31) As CellTypes
Public currentSel1 As CellTypes
Public currentSel2 As CellTypes
Public goAhead As Boolean

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub cmdOK_Click()
goAhead = True
currentSel1 = lookupTable(cboSearch.ListIndex)
currentSel2 = lookupTable(cboReplace.ListIndex)
Me.Hide
End Sub

Private Sub Form_Load()
'Populate the lists and lookup table with celltypes
cboSearch.AddItem "Fast Passable 1 (F1)"
cboReplace.AddItem "Fast Passable 1 (F1)"
lookupTable(0) = FastPassable1
cboSearch.AddItem "Fast Passable 2 (F2)"
cboReplace.AddItem "Fast Passable 2 (F2)"
lookupTable(1) = FastPassable2
cboSearch.AddItem "Medium Passable 1 (M1)"
cboReplace.AddItem "Medium Passable 1 (M1)"
lookupTable(2) = MediumPassable1
cboSearch.AddItem "Medium Passable 2 (M2)"
cboReplace.AddItem "Medium Passable 2 (M2)"
lookupTable(3) = MediumPassable2
cboSearch.AddItem "Slow Passable 1 (S1)"
cboReplace.AddItem "Slow Passable 1 (S1)"
lookupTable(4) = SlowPassable1
cboSearch.AddItem "Slow Passable 2 (S2)"
cboReplace.AddItem "Slow Passable 2 (S2)"
lookupTable(5) = SlowPassable2
cboSearch.AddItem "Impassable 1 (I1)"
cboReplace.AddItem "Impassable 1 (I1)"
lookupTable(6) = Impassable1
cboSearch.AddItem "Impassable 2 (I2)"
cboReplace.AddItem "Impassable 2 (I2)"
lookupTable(7) = Impassable2
cboSearch.AddItem "North Cliffs (NC)"
cboReplace.AddItem "North Cliffs (NC)"
lookupTable(8) = NorthCliffs
cboSearch.AddItem "Cliffs - High Side (CHS)"
cboReplace.AddItem "Cliffs - High Side (CHS)"
lookupTable(9) = CliffsHighSide
cboSearch.AddItem "Cliffs - Low Side (CLS)"
cboReplace.AddItem "Cliffs - Low Side (CLS)"
lookupTable(10) = CliffsLowSide
cboSearch.AddItem "zPad 12 (Z12)"
cboReplace.AddItem "zPad 12 (Z12)"
lookupTable(11) = zPad12
cboSearch.AddItem "zPad 13 (Z13)"
cboReplace.AddItem "zPad 13 (Z13)"
lookupTable(12) = zPad13
cboSearch.AddItem "zPad 14 (Z14)"
cboReplace.AddItem "zPad 14 (Z14)"
lookupTable(13) = zPad14
cboSearch.AddItem "zPad 15 (Z15)"
cboReplace.AddItem "zPad 15 (Z15)"
lookupTable(14) = zPad15
cboSearch.AddItem "zPad 16 (Z16)"
cboReplace.AddItem "zPad 16 (Z16)"
lookupTable(15) = zPad16
cboSearch.AddItem "zPad 17 (Z17)"
cboReplace.AddItem "zPad 17 (Z17)"
lookupTable(16) = zPad17
cboSearch.AddItem "zPad 18 (Z18)"
cboReplace.AddItem "zPad 18 (Z18)"
lookupTable(17) = zPad18
cboSearch.AddItem "zPad 19 (Z19)"
cboReplace.AddItem "zPad 19 (Z19)"
lookupTable(18) = zPad19
cboSearch.AddItem "zPad 20 (Z20)"
cboReplace.AddItem "zPad 20 (Z20)"
lookupTable(19) = zPad20
cboSearch.AddItem "Bulldozed (D)"
cboReplace.AddItem "Bulldozed (D)"
lookupTable(20) = DozedArea
cboSearch.AddItem "Rubble (R)"
cboReplace.AddItem "Rubble (R)"
lookupTable(21) = Rubble
cboSearch.AddItem "Fumaroles/Vents (V)"
cboReplace.AddItem "Fumaroles/Vents (V)"
lookupTable(22) = VentsAndFumaroles
cboSearch.AddItem "Normal Wall (NW)"
cboReplace.AddItem "Normal Wall (NW)"
lookupTable(23) = NormalWall
cboSearch.AddItem "Microbe Wall (MW)"
cboReplace.AddItem "Microbe Wall (MW)"
lookupTable(24) = MicrobeWall
cboSearch.AddItem "Lava Wall (LW)"
cboReplace.AddItem "Lava Wall (LW)"
lookupTable(25) = LavaWall
cboSearch.AddItem "Tube 0 (T0)"
cboReplace.AddItem "Tube 0 (T0)"
lookupTable(26) = Tube0
cboSearch.AddItem "Tube 1 (T1)"
cboReplace.AddItem "Tube 1 (T1)"
lookupTable(27) = Tube1
cboSearch.AddItem "Tube 2 (T2)"
cboReplace.AddItem "Tube 2 (T2)"
lookupTable(28) = Tube2
cboSearch.AddItem "Tube 3 (T3)"
cboReplace.AddItem "Tube 3 (T3)"
lookupTable(29) = Tube3
cboSearch.AddItem "Tube 4 (T4)"
cboReplace.AddItem "Tube 4 (T4)"
lookupTable(30) = Tube4
cboSearch.AddItem "Tube 5 (T5)"
cboReplace.AddItem "Tube 5 (T5)"
lookupTable(31) = Tube5
cboSearch.ListIndex = 0
cboReplace.ListIndex = 0
currentSel1 = lookupTable(0)
currentSel2 = lookupTable(0)
End Sub
