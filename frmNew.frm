VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "VBALTR~1.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New File"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin vbalTreeViewLib6.vbalTreeView tvObjects 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      FullRowSelect   =   -1  'True
      HotTracking     =   0   'False
      SingleSel       =   -1  'True
      Style           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin vbalDTab6.vbalDTabControl DTabControl 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      AllowScroll     =   0   'False
      TabAlign        =   0
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
   Begin VB.Label lblTopInfo 
      Caption         =   "Create new file of type:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim filesIml As New cVBALImageList

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If tvObjects.SelectedItem Is Nothing Then Exit Sub
Select Case tvObjects.SelectedItem.Key
    Case "map"
        Unload Me
        frmNewMapSet.Show 1, fMainForm
        
    Case "vol"
        Unload Me
        Dim frmVol As VolManager
        Set frmVol = New VolManager
        If frmVol.CreateVol Then
            docNumber = docNumber + 1
            frmVol.SetNewName "Untitled" & CStr(docNumber), "Untitled" & CStr(docNumber)
            frmVol.Show
        Else
            Set frmVol = Nothing
        End If
End Select
End Sub

Private Sub DTabControl_TabSelected(theTab As vbalDTab6.cTab)
Select Case theTab.Key
    Case "file"
        tvObjects.Nodes.Clear
        tvObjects.Nodes.Add , etvwNext, "map", "Map File", 0
        'tvObjects.Nodes.Add , etvwNext, "script", "Mission Script", 1
        'tvObjects.Nodes.Add , etvwNext, "techtree", "Tech Tree", 2
        tvObjects.Nodes.Add , etvwNext, "tileset", "Tile Set", 3
        tvObjects.Nodes.Add , etvwNext, "vol", "VOL Archive", 4
    Case "proj"
        tvObjects.Nodes.Clear
        tvObjects.Nodes.Add , etvwNext, "multi", "Multiplayer Scenario", 5
        tvObjects.Nodes.Add , etvwNext, "colony", "Colony Game", 5
        tvObjects.Nodes.Add , etvwNext, "tutor", "Tutorial", 5
        tvObjects.Nodes.Add , etvwNext, "campaign", "Campaign Mission", 5
        tvObjects.Nodes.Add , etvwNext, "demo", "AutoDemo", 5
        tvObjects.Nodes.Add , etvwNext, "blank", "(Empty Project)", 5
End Select
End Sub

Private Sub Form_Load()
filesIml.OwnerHDC = Me.hdc
filesIml.ColourDepth = filesIml.SystemColourDepth
filesIml.IconSizeX = 16
filesIml.IconSizeY = 16
filesIml.Create
Dim hImg As Picture
Set hImg = LoadResPicture(91, vbResIcon)
filesIml.AddFromHandle hImg.Handle, IMAGE_ICON
Set hImg = LoadResPicture(92, vbResIcon)
filesIml.AddFromHandle hImg.Handle, IMAGE_ICON
Set hImg = LoadResPicture(93, vbResIcon)
filesIml.AddFromHandle hImg.Handle, IMAGE_ICON
Set hImg = LoadResPicture(94, vbResIcon)
filesIml.AddFromHandle hImg.Handle, IMAGE_ICON
Set hImg = LoadResPicture(95, vbResIcon)
filesIml.AddFromHandle hImg.Handle, IMAGE_ICON
Set hImg = LoadResPicture(96, vbResIcon)
filesIml.AddFromHandle hImg.Handle, IMAGE_ICON
Set hImg = Nothing

tvObjects.ImageList = filesIml.hIml

DTabControl.Tabs.Add "file", , "Files"
'DTabControl.Tabs.Add "proj", , "Projects"
DTabControl.Tabs.Item("file").Selected = True
End Sub

Private Sub tvObjects_NodeDblClick(node As vbalTreeViewLib6.cTreeViewNode)
cmdOK_Click
End Sub
