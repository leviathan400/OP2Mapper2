VERSION 5.00
Begin VB.Form ScriptManager 
   Caption         =   "Script"
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
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin Mapper2.vbalTreeView tvScript 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   6588
      _ExtentY        =   4048
      HotTracking     =   0   'False
      LineStyle       =   0
      SingleSel       =   -1  'True
      ScaleMode       =   3
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
End
Attribute VB_Name = "ScriptManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum ScriptTypes
    MultiLOS = 0
    '**TODO** implement the scripttypes below
    MultiMidas = 1
    MultiRR = 2
    MultiSR = 3
    MultiLR = 4
    SingleTut = 5
    SingleDemo = 6
    SingleEden = 7
    SinglePly = 8
    SingleColonyPop = 9
    SingleColonyStar = 10
    CustomScript = 11
End Enum
'Public Enum ScriptCmdTypes
'    Set
'End Enum
Private scriptCmds As New Collection

'Private Type ScriptCmd
'    cmdType As ScriptCmdTypes
'End Type
Private Sub Form_Resize()
'cboEditWhat.Width = Me.ScaleWidth
'tvScript.Width = Me.ScaleWidth
'tvScript.Height = Me.ScaleHeight - cboEditWhat.Height
End Sub

Public Sub CreateNew(ByVal scrType As ScriptTypes, ByVal nPlayers As Integer)
Select Case scrType
    Case MultiLOS
        
    Case Else 'nothing else implemented - **TODO**
        GenerateError "TODO: implement other script types", "ScriptManager.CreateNew"
End Select
End Sub

Public Sub RefreshTree()

End Sub
