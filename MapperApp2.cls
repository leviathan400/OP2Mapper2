VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapperApp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'MapperApp2 module for plugins to use.
'Exports all the functions they'll ever need
'Implements IAppObject

'BASE FUNCTIONS

'Not really something for the plugin module itself, but it's there
'Public pluginId As Long

'<All Versions> Returns the version of the library/mapper
' in number format (200 = 2.0.0)
'Private Function IAppObject_GetLibVersion() As Long
'GetLibVersion = CLng((App.Major * 100) + (App.Minor * 10) + (App.Revision))
'End Function
'
''<All Versions> Produces a shell error
'Private Sub IAppObject_RaiseError(ByVal desc As String)
'GenerateError desc, "<Plugin>::AppObject"
'End Sub
'
''<All Versions> Attempts to enter privileged mode, the user may need
'' to authorize this, returns 0 on failure
'Private Function IAppObject_EnterPrivilegedMode() As Long
'If privMode = True Then EnterPrivilegedMode = 1: Exit Function
'Dim fMsg As New frmPluginMsg
'fMsg.SetMsg "The currently running plugin is requesting to enter privileged mode. This means it may be able to execute commands that could potentially harm open files." & vbNewLine & "If this was requested, or you know what the plugin will do, it it safe to allow this."
'fMsg.Show 1
'If fMsg.Tag = "1" Then
'    'Enter privileged mode
'    privMode = True
'    EnterPrivilegedMode = 1
'Else
'    privMode = False
'    EnterPrivilegedMode = 0
'End If
'Unload fMsg
'Set fMsg = Nothing
'End Function
'
''<All Versions> Leaves privileged mode
'Private Sub IAppObject_LeavePrivilegedMode()
'privMode = False
'End Sub
'
''<All Versions> Installs a menu item
'Private Sub IAppObject_InstallMenuItem(ByVal strName As String, ByVal lngCode As Long)
'ReDim Preserve loadedPlugins(pluginId).menuItems(loadedPlugins(pluginId).numMenuItems)
'loadedPlugins(pluginId).menuItems(loadedPlugins(pluginId).numMenuItems).strTitle = strName
'loadedPlugins(pluginId).menuItems(loadedPlugins(pluginId).numMenuItems).lngCode = lngCode
'loadedPlugins(pluginId).numMenuItems = loadedPlugins(pluginId).numMenuItems + 1
'RefreshPluginMenu
'End Sub
'
''<All Versions> Installs an event hook
'Private Sub IAppObject_InstallEventHook(ByVal lngEventType As Long, ByVal lngCode As Long)
'ReDim Preserve loadedPlugins(pluginId).eventHooks(loadedPlugins(pluginId).numEventHooks)
'loadedPlugins(pluginId).eventHooks(loadedPlugins(pluginId).numEventHooks).lngHookType = lngEventType
'loadedPlugins(pluginId).eventHooks(loadedPlugins(pluginId).numEventHooks).lngCode = lngCode
'loadedPlugins(pluginId).numEventHooks = loadedPlugins(pluginId).numEventHooks + 1
''**TODO**: handle events.
'End Sub
'
''<All Versions> Asks a yes/no question to the user
''Returns 1 on Yes, 0 on No
'Private Function IAppObject_YesNoBox(ByVal strMsg As String) As Long
'Dim fMsg As New frmPluginMsg
'fMsg.SetMsg strMsg
'fMsg.Show 1
'If fMsg.Tag = "1" Then
'    YesNoBox = 1
'Else
'    privMode = False
'    YesNoBox = 0
'End If
'Unload fMsg
'Set fMsg = Nothing
'End Function
'
''MAP EDITING
'
''<All Versions, Privileged> Gets the MapFile object from the currently active map
'' so the plugin can work on it, returns Nothing if failed
'Private Function IAppObject_GetMapObject() As Object
'If PrivModeCheck() = False Then Exit Function
'GetMapObject = Nothing
'If fMainForm.ActiveForm Is Nothing Then Exit Function
'If fMainForm.ActiveForm.Name = "MapManager" Then GetMapObject = fMainForm.ActiveForm.myMap
'End Function
'
''<All Versions> Redraws the current map
'' returns 0 if failed (no active map)
'Private Function IAppObject_RedrawMap() As Long
'RedrawMap = 0
'If fMainForm.ActiveForm Is Nothing Then Exit Function
'If fMainForm.ActiveForm.Name = "MapManager" Then fMainForm.ActiveForm.RedrawSelf: RedrawMap = 1
'End Function
'
''<All Versions> Changes the currently open editing tool
'Private Sub IAppObject_ChangeMapTool(ByVal tool As TOOL_MODE)
'If fMainForm.ActiveForm Is Nothing Then Exit Sub
'If fMainForm.ActiveForm.Name = "MapManager" Then fMainForm.ActiveForm.ChangeTool tool
'End Sub
'
''Other functions not in the IAppObject interface
'Private Function PrivModeCheck() As Boolean
'If privMode = False Then
'    'Priv mode check?
'    GenerateError "Plugin attempted to call privileged method when it did not have permissions to do so", "<Plugin>::AppObject"
'    PrivModeCheck = False
'    Exit Function
'Else
'    PrivModeCheck = True
'End If
'End Function
