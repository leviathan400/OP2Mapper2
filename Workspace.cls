VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Workspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Workspace manager class.
Private maps() As MapManager
Private nMaps As Integer

Public Enum WorkspaceObjType
    map = 0
    Script = 1
    Techtree = 2
End Enum

'For the add functions, if the name is blank, an empty file is created

Public Function AddScript(ByVal strName As String) As Integer

End Function

Public Function AddMap(ByVal strName As String, ByVal strTitle As String) As Integer
Dim newMap As MapManager
Set newMap = New MapManager
If strName = "" Then
    'create a new one
    'If newMap.SelfMakeMap = False Then Exit Function
Else
    'load existing
    If newMap.LoadMap(strName) = False Then Exit Function
End If
'set the name on it
newMap.SetNewName strName, strTitle
newMap.Show
'Set the map to the array
ReDim Preserve maps(nMaps)
Set maps(nMaps) = newMap
maps(nMaps).Show

'Return the id
AddMap = nMaps

'Increment the counter
nMaps = nMaps + 1
End Function

Public Function AddTechtree(ByVal strName As String) As Integer

End Function

Public Function RemoveItem(ByVal objtype As WorkspaceObjType, ByVal id As Integer) As Boolean

End Function

Public Function SaveWorkspace(ByVal strFile As String) As Boolean
'Save to a simple format
'**TODO**
End Function

Public Function LoadWorkspace(ByVal strFile As String) As Boolean

End Function

Public Function ShutDown() As Boolean
'**TODO** Pop up a box asking if they want to save the project
ShutDown = True
End Function

Private Sub Class_Terminate()
'Close the open workspace and prompt to save all files
End Sub
