VERSION 5.00
Begin VB.Form frmNewMapSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Map Settings"
   ClientHeight    =   2295
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
   Icon            =   "frmNewMapSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkNoTilesets 
      Caption         =   "Start with no tilesets or terrains?"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox cboH 
      Height          =   315
      ItemData        =   "frmNewMapSet.frx":000C
      Left            =   960
      List            =   "frmNewMapSet.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox cboW 
      Height          =   315
      ItemData        =   "frmNewMapSet.frx":0010
      Left            =   960
      List            =   "frmNewMapSet.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   "Note: Maps with a width of 512 or higher are automatically Around the World maps."
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label lblH 
      Caption         =   "Height:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblW 
      Caption         =   "Width:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmNewMapSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim mapForm As MapManager, w As Long, h As Long, tilesets As Long
w = CLng(cboW.Text)
h = CLng(cboH.Text)
tilesets = IIf(chkNoTilesets.Value = vbChecked, 0, 1)

Unload Me
Set mapForm = New MapManager
mapForm.CreateMap w, h, tilesets
docNumber = docNumber + 1
mapForm.SetNewName "Untitled" & CStr(docNumber), "Untitled" & CStr(docNumber)
mapForm.Show
End Sub

Private Sub Form_Load()
cboW.AddItem "32"
cboW.AddItem "64"
cboW.AddItem "128"
cboW.AddItem "256"
cboW.AddItem "512"
cboW.AddItem "1024"

cboH.AddItem "32"
cboH.AddItem "64"
cboH.AddItem "128"
cboH.AddItem "256"
cboH.AddItem "512"
cboH.AddItem "1024"

cboW.ListIndex = 0
cboH.ListIndex = 0
End Sub
