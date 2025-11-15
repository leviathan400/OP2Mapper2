VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TileSetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim LoadedSets() As TileSet 'array of valid TileSet objects
Dim SetDescs() As String 'descriptions of tilesets; for the interface
Dim curPos As Long 'current window position

Public Sub AddSet()
'application calls to add a tileset to the manager

End Sub

Public Sub DelSet()
'application calls to remove a tileset from the manager

End Sub

Public Sub Redraw(lPos As Long)
'application calls to redraw as needed for the set window
'self calls when a tileset is added or removed

End Sub
