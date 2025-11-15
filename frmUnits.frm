VERSION 5.00
Begin VB.Form frmUnits 
   Caption         =   "Object Selector"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
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
   Icon            =   "frmUnits.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.ListBox lstUnits 
      Height          =   2700
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.ComboBox cboPlayer 
      Height          =   315
      ItemData        =   "frmUnits.frx":014A
      Left            =   0
      List            =   "frmUnits.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private uEntry() As UnitEntry
Public nEntries As Long

Private mCurSel As UnitEntry

Public ownerForm As MapManager

Public CurrentPlayer As Long
Public prepareClose As Boolean

Private Sub Form_Deactivate()
If ownerForm Is Nothing Then Exit Sub
ownerForm.ToolDeactivated
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
If prepareClose = False Then cancel = True
End Sub

Private Sub cboPlayer_Click()
If cboPlayer.ListIndex >= 0 Then
    If cboPlayer.ListIndex <= 6 Then
        UpdateListUnits
        CurrentPlayer = cboPlayer.ListIndex
    Else
        UpdateListGaia
        CurrentPlayer = 0
    End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo oops
'Set window settings
Me.Left = GetSettingIni("Window", "UnitsLeft", 0)
Me.TOp = GetSettingIni("Window", "UnitsTop", 0)
Me.Width = GetSettingIni("Window", "UnitsWidth", 4800)
Me.Height = GetSettingIni("Window", "UnitsHeight", 3600)
Me.WindowState = GetSettingIni("Window", "UnitsWindowState", 0)

UpdateListUnits

'Now set up the player combo
cboPlayer.AddItem "Player 0 (Human), Blue"
cboPlayer.AddItem "Player 1 (Human/AI), Red"
cboPlayer.AddItem "Player 2 (Human/AI), Green"
cboPlayer.AddItem "Player 3 (Human/AI), Yellow"
cboPlayer.AddItem "Player 4 (Human/AI), Cyan"
cboPlayer.AddItem "Player 5 (Human/AI), Magenta"
cboPlayer.AddItem "Player 6 (AI), Black"
cboPlayer.AddItem "Gaia/Game Objects"
cboPlayer.ListIndex = 0
Exit Sub
oops:
GenerateError "Couldn't load unit structures. Invalid ctl files?", "frmUnits::Load"
End Sub

Private Sub Form_Resize()
On Error Resume Next
cboPlayer.Width = Me.ScaleWidth
lstUnits.Move 0, cboPlayer.Height, Me.ScaleWidth, Me.ScaleHeight - cboPlayer.Height
End Sub

Public Sub SetNewName(ByVal strName As String)
Me.Caption = "Objects - " & strName
End Sub

Private Sub Form_Unload(cancel As Integer)
If Me.WindowState <> vbMinimized Then
    SaveSettingIni "Window", "UnitsLeft", Me.Left
    SaveSettingIni "Window", "UnitsTop", Me.TOp
    SaveSettingIni "Window", "UnitsWidth", Me.Width
    SaveSettingIni "Window", "UnitsHeight", Me.Height
    SaveSettingIni "Window", "UnitsWindowState", Me.WindowState
End If
End Sub

Private Sub lstUnits_Click()
'select the right record into the current selection
Dim i As Long
For i = 0 To nEntries - 1
    If uEntry(i).strName = lstUnits.Text Then 'this is it
        mCurSel = uEntry(i)
        UpdateRec
        Exit Sub
    End If
Next
End Sub

Public Sub UpdateRec()
'update the global record
curUnitSel = mCurSel
End Sub

Public Sub UpdateListUnits()
Erase uEntry()
lstUnits.Clear
'Create an array of unitrec's to store the list data
Dim i As Long, j As Long
For i = 0 To numUnitDefs
    If allUnitDefs(i).isGaia = False Then
        If allUnitDefs(i).canHaveTurret = True Then
            'Add individual entries for the turrets
            For j = 0 To numWeaponDefs
                ReDim Preserve uEntry(nEntries)
                nEntries = nEntries + 1
                uEntry(nEntries - 1).unitDef = allUnitDefs(i)
                uEntry(nEntries - 1).weaponDef = allWeaponDefs(j)
                uEntry(nEntries - 1).strName = allUnitDefs(i).strName & "/" & allWeaponDefs(j).strName
                lstUnits.AddItem uEntry(nEntries - 1).strName
            Next
        Else
            'Just one record
            ReDim Preserve uEntry(nEntries)
            nEntries = nEntries + 1
            uEntry(nEntries - 1).unitDef = allUnitDefs(i)
            uEntry(nEntries - 1).strName = allUnitDefs(i).strName
            lstUnits.AddItem uEntry(nEntries - 1).strName
        End If
    End If
Next
lstUnits.ListIndex = 0
'init the current sel
lstUnits_Click
End Sub

Public Sub UpdateListGaia()
Erase uEntry()
lstUnits.Clear
'Create an array of unitrec's to store the list data
Dim i As Long, j As Long
For i = 0 To numUnitDefs
    If allUnitDefs(i).isGaia = True Then
        'Gaia record
        ReDim Preserve uEntry(nEntries)
        nEntries = nEntries + 1
        uEntry(nEntries - 1).unitDef = allUnitDefs(i)
        uEntry(nEntries - 1).strName = allUnitDefs(i).strName
        lstUnits.AddItem uEntry(nEntries - 1).strName
    End If
Next
lstUnits.ListIndex = 0
'init the current sel
lstUnits_Click
End Sub
