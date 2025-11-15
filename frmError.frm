VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
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
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Copy the text of the error to the clipboard so you can submit it to OPU staff."
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      ToolTipText     =   "Close the error dialog."
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtError 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   6495
   End
   Begin VB.Label lblInfo3 
      Caption         =   "As a precaution, we recommend that you save your work as soon as possible to prevent any data loss."
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   6495
   End
   Begin VB.Label lblInfo2 
      Caption         =   $"frmError.frx":5D52
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label lblInfo1 
      Caption         =   "Application Error Occurred!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmError.frx":5DDC
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtError.Text
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub
