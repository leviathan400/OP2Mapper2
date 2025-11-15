VERSION 5.00
Begin VB.Form frmGenUnitCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate C++ Code"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
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
   Icon            =   "frmGenUnitCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtFilename 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Frame grpCodeType 
      Caption         =   "Code Generation Template"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6255
      Begin VB.ComboBox cboTemplate 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select a template to use when generating code."
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblNote 
         Caption         =   "Extra templates must be stored in the map editor's directory."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label lblTemplate 
         Caption         =   "Use template:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   255
         Width           =   2805
      End
   End
   Begin VB.Label lblSaveFile 
      Caption         =   "Save the generated file as:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1695
      Width           =   2535
   End
   Begin VB.Label lblInfo0 
      Caption         =   $"frmGenUnitCode.frx":000C
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmGenUnitCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public templateName As String
Public outputName As String

Private Sub cmdCancel_Click()
'cancel operation
templateName = ""
outputName = ""
Me.Hide
End Sub

Private Sub cmdGenerate_Click()
'validate inputted data
templateName = cboTemplate.Text
If Trim(txtFilename.Text) = "" Then Beep: txtFilename.SetFocus: Exit Sub
outputName = Trim(txtFilename.Text)
Me.Hide
End Sub

Private Sub Form_Load()
'Enumerate all .tpl files in the Mapper dir
templateName = ""
outputName = ""
Dim theFile As String
theFile = Dir(App.Path & "\*.tpl")
Do Until theFile = ""
    cboTemplate.AddItem theFile
    theFile = Dir()
Loop
If cboTemplate.ListCount = 0 Then
    MsgBox "There are no template files to generate code from. Place a .tpl file in the map editor directory.", vbInformation, "No Template Files"
    Unload Me
Else
    cboTemplate.ListIndex = 0
End If
End Sub
