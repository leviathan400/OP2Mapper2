VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
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
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.CommandButton cmdDown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      ToolTipText     =   "Move Down"
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      ToolTipText     =   "Move Up"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      ToolTipText     =   "Delete"
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      ToolTipText     =   "Add"
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox lstVols 
      Height          =   1785
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   $"frmOptions.frx":000C
      Top             =   1320
      Width           =   5415
   End
   Begin VB.CheckBox chkAutosave 
      Caption         =   "Autosave after how many minutes:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   $"frmOptions.frx":0132
      Top             =   4575
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.TextBox txtAutosaveMin 
      Height          =   315
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "5"
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox txtSavePath 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "This path specifies the default path Open and Save windows will switch to."
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton cmdSelSavePath 
      Caption         =   "..."
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.CheckBox chkEnableGrid 
      Caption         =   "Enable Grid by Default"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "The setting of this checkbox determines what the grid setting will be when opening/creating a map."
      Top             =   4200
      Value           =   1  'Checked
      Width           =   5895
   End
   Begin VB.CheckBox chkDisablePolling 
      Caption         =   "Disable Keyboard Polling / Scroll-by-ArrowKeys"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Checking this will prevent arrow key scrolling in maps."
      Top             =   3840
      Width           =   5895
   End
   Begin VB.CommandButton cmdSelectPath 
      Caption         =   "..."
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "This path specifies the path to your OP2 game files. It is critical that you set this path correctly."
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Tag             =   "Cancel"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblWarn 
      Caption         =   "Note: You will need to restart the Mapper for any changes in the list above to take effect."
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Label lblVols 
      Caption         =   "List of additional VOLs to look at when locating files:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label lblSavePath 
      Caption         =   "Default Save Dir:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   630
      Width           =   1935
   End
   Begin VB.Label lblPath 
      Caption         =   "Outpost 2 Dir:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   1935
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutosave_Click()
txtAutosaveMin.Enabled = IIf(chkAutosave.Value, True, False)
End Sub

Private Sub cmdAdd_Click()
Dim tmpStr As String
tmpStr = Trim$(InputBox("Enter the name of the VOL to load. Paths may be absolute or relative, if no path is given it is assumed the VOL resides in the OP2 dir.", "Enter Filename"))
If tmpStr = "" Then Exit Sub
lstVols.AddItem tmpStr
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDel_Click()
If lstVols.ListIndex >= 0 Then
    If MsgBox("Are you sure you want to remove " & lstVols.Text & "?", vbQuestion Or vbYesNo, "Remove File") = vbYes Then
        lstVols.RemoveItem lstVols.ListIndex
    End If
End If
End Sub

Private Sub cmdDown_Click()
Dim tmpStr As String
'Swap values
If lstVols.ListIndex >= 0 And lstVols.ListCount >= 2 And lstVols.ListIndex < lstVols.ListCount - 1 Then
    tmpStr = lstVols.List(lstVols.ListIndex + 1)
    lstVols.List(lstVols.ListIndex + 1) = lstVols.List(lstVols.ListIndex)
    lstVols.List(lstVols.ListIndex) = tmpStr
    lstVols.ListIndex = lstVols.ListIndex + 1
End If
End Sub

Private Sub cmdOK_Click()
Dim tmpStr As String, i As Long
SaveSettingIni "Paths", "Outpost2", txtPath.Text
SaveSettingIni "Paths", "Save", txtSavePath.Text
ResMan.RootPath = txtPath.Text
mapsDir = txtSavePath.Text
If lstVols.ListCount > 0 Then
    For i = 0 To lstVols.ListCount - 1
        If i = lstVols.ListCount - 1 Then
            'Don't add a | at the end
            tmpStr = tmpStr & lstVols.List(i)
        Else
            tmpStr = tmpStr & lstVols.List(i) & "|"
        End If
    Next
    SaveSettingIni "Paths", "CustomVOL", tmpStr
Else
    DeleteSettingIni "Paths", "CustomVOL"
End If
SaveSettingIni "Options", "DisableKBPoll", IIf(chkDisablePolling.Value, "1", "0")
SaveSettingIni "Options", "DisableGrid", IIf(chkEnableGrid.Value, "0", "1")
enableKbPoll = IIf(chkDisablePolling.Value, False, True)
enableGridDef = IIf(chkEnableGrid.Value, True, False)

If chkAutosave.Value = vbChecked Then
    SaveSettingIni "Options", "AutosaveMin", CLng(txtAutosaveMin.Text)
    autoSaveAfter = TimeSerial(0, txtAutosaveMin.Text, 0)
    fMainForm.tmrAutosave.Enabled = True
Else
    SaveSettingIni "Options", "AutosaveMin", "0"
    autoSaveAfter = TimeSerial(0, 0, 0)
    fMainForm.tmrAutosave.Enabled = False
End If
curAutosaveMin = 0
curAutosaveSec = 0
SetStatusAutosave fMainForm.tmrAutosave.Enabled, TimeSerial(0, curAutosaveMin, curAutosaveSec)

Unload Me
End Sub

Private Sub cmdSelectPath_Click()
Dim newStr As String
newStr = BrowseForFolder(Me.hWnd, "Please specify the folder where your OP2 mapsXX.vol files are located.")
If newStr = "" Then Exit Sub
txtPath.Text = newStr
End Sub

Private Sub cmdSelSavePath_Click()
Dim newStr As String
newStr = BrowseForFolder(Me.hWnd, "Please specify the default folder where you will be saving files.")
If newStr = "" Then Exit Sub
txtSavePath.Text = newStr
End Sub

Private Sub cmdUp_Click()
Dim tmpStr As String
'Swap values
If lstVols.ListIndex >= 1 And lstVols.ListCount >= 2 Then
    tmpStr = lstVols.List(lstVols.ListIndex - 1)
    lstVols.List(lstVols.ListIndex - 1) = lstVols.List(lstVols.ListIndex)
    lstVols.List(lstVols.ListIndex) = tmpStr
    lstVols.ListIndex = lstVols.ListIndex - 1
End If
End Sub

Private Sub Form_Load()
'Load the settings from ini
Dim tmpStr As String, tmpArr() As String, i As Long
txtPath.Text = GetSettingIni("Paths", "Outpost2")
txtSavePath.Text = GetSettingIni("Paths", "Save")
tmpStr = GetSettingIni("Paths", "CustomVOL")
If tmpStr <> "" Then
    tmpArr = Split(tmpStr, "|")
    For i = 0 To UBound(tmpArr)
        lstVols.AddItem tmpArr(i)
    Next
End If
chkDisablePolling.Value = IIf(GetSettingIni("Options", "DisableKBPoll", "0") = "1", vbChecked, vbUnchecked)
chkEnableGrid.Value = IIf(GetSettingIni("Options", "DisableGrid", "0") = "1", vbUnchecked, vbChecked)
chkAutosave.Value = IIf(GetSettingIni("Options", "AutosaveMin", "5") = "0", vbUnchecked, vbChecked)
If chkAutosave.Value = vbChecked Then txtAutosaveMin.Text = GetSettingIni("Options", "AutosaveMin", "5")
End Sub

