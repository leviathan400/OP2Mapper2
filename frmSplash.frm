VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6000
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   5775
   End
   Begin VB.Image imgEditor 
      Height          =   3375
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
