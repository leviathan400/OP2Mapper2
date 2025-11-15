VERSION 5.00
Begin VB.Form frmTileInfo 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTileInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000017&
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmTileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub txtInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub
