VERSION 5.00
Begin VB.Form frmPluginMsg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plugin Msg/Input"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblMessage 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmPluginMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNo_Click()
Me.Tag = ""
Me.Hide
End Sub

Private Sub cmdYes_Click()
Me.Tag = "1"
Me.Hide
End Sub

Public Sub SetMsg(ByVal msgText As String)
lblMessage.Caption = msgText
Me.Refresh
End Sub
