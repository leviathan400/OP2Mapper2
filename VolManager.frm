VERSION 5.00
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form VolManager 
   Caption         =   "VOL"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
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
   Icon            =   "VolManager.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   Begin vbalTBar6.cReBar myRebar 
      Left            =   2760
      Top             =   0
      _ExtentX        =   3201
      _ExtentY        =   661
   End
   Begin vbalTBar6.cToolbar myToolbar 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
   End
   Begin vbalTreeViewLib6.vbalTreeView tvFiles 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4683
      FullRowSelect   =   -1  'True
      HotTracking     =   0   'False
      SingleSel       =   -1  'True
      Style           =   1
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
Attribute VB_Name = "VolManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public myName As String, myTitle As String, isFile As Boolean
Private myIml As New cVBALImageList
Private myTabId As Long
Private nFiles As Long
Private nextListId As Long

Private Type VOLRecord
    strName As String
    strNewName As String 'Has a value if the file is being renamed
    isRenamed As Boolean 'true if the file is being renamed (to strNewName)
    isDeleted As Boolean 'true if the file is to be deleted at save
    isNewlyAdded As Boolean 'true if the file doesn't already exist in the vol. strName = whole path, strNewName = short name
    listId As Long 'Key in the list
End Type
Private volRecs() As VOLRecord

Private Sub Form_Load()
'Set up the imagelist
myIml.OwnerHDC = fMainForm.picStatus.hDC
myIml.ColourDepth = myIml.SystemColourDepth
myIml.IconSizeX = 16
myIml.IconSizeY = 16
myIml.Create
Dim hImg As Picture
Set hImg = LoadResPicture(104, vbResBitmap)
myIml.AddFromHandle hImg.handle, IMAGE_BITMAP, , &HFF00FF
'Add icons for all the different filetypes
Set hImg = LoadResPicture(91, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(92, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(93, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(94, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(95, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(96, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(97, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(98, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(99, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(100, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(101, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(102, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = LoadResPicture(901, vbResIcon)
myIml.AddFromHandle hImg.handle, IMAGE_ICON
Set hImg = Nothing

tvFiles.ImageList = myIml.hIml

'Set up the toolbar
myToolbar.ImageSource = CTBExternalImageList
myToolbar.SetImageList myIml.hIml
myToolbar.CreateToolbar 16, , , True
myToolbar.AddButton "Add File", 0, , , , CTBNormal, "addfile"
myToolbar.AddButton "Add Directory", 1, , , , CTBNormal, "adddir"
myToolbar.AddButton "Delete File", 2, , , , CTBNormal, "delfile"
myToolbar.AddButton , , , , , CTBSeparator
myToolbar.AddButton "Extract File", 3, , , , CTBNormal, "extfile"
myToolbar.AddButton "Extract All Files", 4, , , , CTBNormal, "extallfiles"
myToolbar.AddButton , , , , , CTBSeparator
myToolbar.AddButton "Rename File", 5, , , , CTBNormal, "renfile"

'Link the toolbar to the rebar
myRebar.CreateRebar Me.hWnd
myRebar.AddBandByHwnd myToolbar.hWnd, , False, , "ToolsBar"

'Add this form to the tab bar
fMainForm.DTabControl.Tabs.Add "VOL" & CStr(fMainForm.DTabControl.Tabs.Count), , Me.Caption
myTabId = fMainForm.DTabControl.Tabs.Count - 1
Me.tag = "VOL" & CStr(myTabId)

End Sub

Public Function LoadVol(ByVal Filename As String) As Boolean
On Error GoTo oops
'Create an ArchiveReader and enumerate the files
tvFiles.Nodes.Clear
nFiles = 0: nextListId = 0
Dim arch As ArchiveReader, i As Long
Set arch = ResMan.LoadVolFile(Filename, True)
For i = 0 To arch.NumFiles - 1
    ReDim Preserve volRecs(nFiles) As VOLRecord
    volRecs(nFiles).strName = arch.Filename(i)
    volRecs(nFiles).listId = nextListId
    tvFiles.Nodes.Add , etvwNext, CStr(nextListId), arch.Filename(i), SelectIcon(arch.Filename(i))
    nFiles = nFiles + 1: nextListId = nextListId + 1
Next
SetStatusBar CStr(nFiles) & " files in archive."
Set arch = Nothing

isFile = True
LoadVol = True
Exit Function
oops:
LoadVol = False
GenerateError "VOL loading failed. Perhaps the VOL is corrupt?", "VolManager::LoadVol"
End Function

Public Sub SetNewName(ByVal strName As String, ByVal sTitle As String)
On Error GoTo baha
'myName = long name as in full path to the file
'myTitle = short name, just the filetitle or other
myName = strName
myTitle = sTitle
Me.Caption = myTitle

'Set it on the tab bar
fMainForm.DTabControl.Tabs.Item("VOL" & CStr(myTabId)).Caption = Me.Caption

Exit Sub
baha:
GenerateError "SetNewName failed", "VolManager::SetNewName"
End Sub

Public Function SaveSelf() As Boolean
Dim tmpStr As String
If isFile = False Then
    'Save as for this file
    If SaveSelfAs = True Then isFile = True Else Exit Function
    SetNewName myName, myTitle
    SaveSelf = True
Else
    'Save the active file
    'Get a temp filename
    tmpStr = String$(261, 0)
    GetTempFileName App.Path, "vol", 0, tmpStr
    tmpStr = Left$(tmpStr, lstrlen(tmpStr))
    
    'Save
    Dim stream As StreamWriter
    Set stream = ResMan.OpenStreamWrite(tmpStr)
    SaveVol myName, stream
        
    Set stream = Nothing
    
    'Now move the temp file to the real location
    Kill myName
    FileCopy tmpStr, myName
    Kill tmpStr

    SaveSelf = True
End If
End Function

Public Function SaveSelfAs() As Boolean
'Save the map file
Dim sFile As String, tmpStr As String
sFile = myName
If fMainForm.cCMDlg.VBGetSaveFileName(sFile, myTitle, True, "VOL Archives (*.vol)|*.vol", , mapsDir, "Save VOL Archive", "vol") = False Then Exit Function

'Get a temp filename
tmpStr = String$(261, 0)
GetTempFileName App.Path, "vol", 0, tmpStr
tmpStr = Left$(tmpStr, lstrlen(tmpStr))

'Save
Dim stream As StreamWriter
Set stream = ResMan.OpenStreamWrite(tmpStr)
SaveVol sFile, stream

myName = sFile

Set stream = Nothing

'Now move the temp file to the real location
If Dir(sFile) <> "" Then Kill sFile
FileCopy tmpStr, sFile
Kill tmpStr

SaveSelfAs = True
SetNewName myName, myTitle
End Function

Public Function SaveVol(ByVal sName As String, strm As StreamWriter)
'Read each vol record and save it
Dim i As Long
On Error GoTo oops
Dim inArch As ArchiveReader
Dim outArch As ArchiveWriter
Dim inStrm As SeekableStreamReader
Set inArch = ResMan.LoadVolFile(sName, True)
Set outArch = ResMan.CreateVolFile
For i = 0 To nFiles - 1
    If volRecs(i).isDeleted = False Then
        If volRecs(i).isNewlyAdded = True Then
            'Load from file (use resource manager functions since it's not in the vol)
            Set inStrm = ResMan.OpenStreamRead(volRecs(i).strName)
            outArch.AddToArchive volRecs(i).strNewName, inStrm, 0
            Set inStrm = Nothing
        Else
            'Load from existing vol
            If volRecs(i).isRenamed = True Then
                'VOL name as dir in path gets automatically handled
                Set inStrm = inArch.OpenStreamRead(volRecs(i).strNewName)
                outArch.AddToArchive volRecs(i).strNewName, inStrm, 0
            Else
                Set inStrm = inArch.OpenStreamRead(volRecs(i).strName)
                outArch.AddToArchive volRecs(i).strName, inStrm, 0
            End If
            Set inStrm = Nothing
        End If
    End If
Next
outArch.WriteArchive strm
Set outArch = Nothing
Set inArch = Nothing
SaveVol = True
Exit Function
oops:
SaveVol = False
GenerateError "VOL Save failed", "VolManager::SaveVol"
End Function

Public Function CreateVol() As Boolean
'Clear data
tvFiles.Nodes.Clear
nFiles = 0: nextListId = 0
Erase volRecs

SetStatusBar "Empty VOL created."
CreateVol = True
End Function

Public Sub AddFile(ByVal sPath As String, ByVal sTitle As String)
'Create a new VOL record for a file
Dim i As Long
'check that the file doesn't already exist
For i = 0 To nFiles - 1
    If (volRecs(i).isRenamed = True And volRecs(i).strNewName = sTitle) Or (volRecs(i).strName = sTitle) Then
        Beep
        SetStatusBar "The name " & sTitle & " is already used by another file."
        Exit Sub
    End If
Next
ReDim Preserve volRecs(nFiles) As VOLRecord
volRecs(nFiles).strName = sPath
volRecs(nFiles).strNewName = sTitle
volRecs(nFiles).isNewlyAdded = True
volRecs(nFiles).listId = nextListId
tvFiles.Nodes.Add , etvwNext, CStr(nextListId), sTitle, SelectIcon(sTitle)
nFiles = nFiles + 1: nextListId = nextListId + 1
SetStatusBar sTitle & " added successfully."
End Sub

Public Sub ExtractFile(ByVal sPathOut As String, ByVal sTitle As String)
On Error GoTo oops
'Extract a file. if the file doesn't exist yet in the volume give the user a message
Dim i As Long, arch As ArchiveReader, strm As SeekableStreamReader, f As Integer
Dim tmpBuff As String, outStrm As StreamWriter
For i = 0 To nFiles - 1
    'Match a file if: New Name is changed, and matches;
    'Or if new name has not changed, see if Name matches
    If (volRecs(i).isRenamed = True And volRecs(i).strNewName = sTitle) Or (volRecs(i).strName = sTitle) Then
        'If the vol file actually exists, and the file is in the VOL
        If isFile And (volRecs(i).isNewlyAdded = False) Then
            'Open a VOL reader and extract the file
            Set arch = ResMan.LoadVolFile(myName, True)
            'Read the file into a string
            Set strm = arch.OpenStreamRead(sTitle)
            tmpBuff = String$(strm.StreamSize, 0)
            strm.Read strm.StreamSize, StrPtr(tmpBuff)
            'Write the string
            Set outStrm = ResMan.OpenStreamWrite(sPathOut)
            outStrm.Write strm.StreamSize, StrPtr(tmpBuff)
            tmpBuff = Empty
            Set outStrm = Nothing
            Set strm = Nothing
            Set arch = Nothing
            SetStatusBar sTitle & " extracted successfully."
        Else
            'Cant do it
            MsgBox "File " & sTitle & " does not actually exist in the VOL file yet, and can't be extracted.", vbExclamation, "Cannot Extract"
        End If
    End If
Next
Exit Sub
oops:
GenerateError "Extraction failed", "VolManager::ExtractFile"
End Sub

Public Function SelectIcon(ByVal Filename As String) As Long
'Pick an icon for the file
Select Case LCase$(Mid$(Filename, InStrRev(Filename, ".", , vbTextCompare) + 1))
    Case "map" 'Map files
        SelectIcon = 6
    Case "scr", "dll", "c", "cpp", "h", "hpp", "tpl" 'Scripts
        SelectIcon = 7
    Case "tek" 'Techtrees
        SelectIcon = 8
    Case "bmp", "pal", "gif", "jpg", "jpeg", "jpe", "png", "psd", "tif", "tga", "pcx" 'Tilesets/Images
        SelectIcon = 9
    Case "vol", "zip", "rar", "ace", "cab", "7z" 'Archives
        SelectIcon = 10
    Case "txt", "def", "diz", "nfo", "1st" 'Text
        SelectIcon = 12
    Case "wav", "aud", "snd", "mp3", "mod", "s3m", "669", "xm", "it", "mo3", "ogg" 'sound/music
        SelectIcon = 13
    Case "rtf", "doc" 'documents
        SelectIcon = 14
    Case "ini", "inf", "set", "cfg" 'ini files
        SelectIcon = 15
    Case "prt", "ctl", "bin", "dat" 'control files
        SelectIcon = 16
    Case "avi", "mpg", "mpeg", "mpe", "mov", "rm", "wma", "asf", "ram" 'movie files
        SelectIcon = 17
    Case Else 'other file
        SelectIcon = 18
End Select
End Function

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
Select Case MsgBox("Do you want to save changes to " & myName & "?", vbQuestion Or vbYesNoCancel, "Save Changes?")
    Case vbYes
        If SaveSelf = False Then cancel = True: Exit Sub
    Case vbCancel
        cancel = True: Exit Sub
End Select
myRebar.RemoveAllRebarBands
End Sub

Private Sub Form_Resize()
On Error Resume Next
tvFiles.Move 0, myRebar.RebarHeight, Me.ScaleWidth, Me.ScaleHeight - myRebar.RebarHeight
'Resize the rebar
myRebar.RebarSize
End Sub

Private Sub Form_Unload(cancel As Integer)
'Remove this from the tab bar
fMainForm.DTabControl.Tabs.Remove "VOL" & CStr(myTabId)
End Sub

Private Sub myToolbar_ButtonClick(ByVal lButton As Long)
Dim sFile As String, sTitle As String, i As Long
Select Case myToolbar.ButtonKey(lButton)
    Case "addfile"
        'Add file to the VOL
        If fMainForm.cCMDlg.VBGetOpenFileName(sFile, sTitle, , , , True, "All Files (*.*)|*.*", , mapsDir, "Add File") = False Then Exit Sub
        AddFile sFile, sTitle
    Case "adddir"
        'Add a whole dir to the VOL
        sFile = BrowseForFolder(Me.hWnd, "Select the folder containing the files you wish to add.")
        If sFile = "" Then Exit Sub
        sTitle = Dir(sFile & "\*.*")
        Do Until sTitle = ""
            AddFile sFile & "\" & sTitle, sTitle
            sTitle = Dir()
            i = i + 1
        Loop
        SetStatusBar CStr(i) & " files added."
    Case "delfile"
        'Remove a file from the VOL
        If tvFiles.SelectedItem Is Nothing Then Exit Sub
        'Mark it as deleted in the VOL list
        volRecs(CLng(tvFiles.SelectedItem.Key)).isDeleted = True
        SetStatusBar tvFiles.SelectedItem.Text & " deleted."
        tvFiles.SelectedItem.Delete
    Case "extfile"
        'Extract a file from the VOL
        If tvFiles.SelectedItem Is Nothing Then Exit Sub
        sFile = tvFiles.SelectedItem.Text
        If fMainForm.cCMDlg.VBGetSaveFileName(sFile, tvFiles.SelectedItem.Text, , "All Files (*.*)|*.*", , mapsDir, "Extract File", LCase$(Mid$(tvFiles.SelectedItem.Text, InStrRev(tvFiles.SelectedItem.Text, ".", , vbTextCompare) + 1))) = False Then Exit Sub
        ExtractFile sFile, tvFiles.SelectedItem.Text
    Case "extallfiles"
        'Extract all files
        sFile = BrowseForFolder(Me.hWnd, "Select the folder which you would like to extract all files into. Note: Existing files will be overwritten!")
        If sFile = "" Then Exit Sub
        sTitle = Dir(sFile & "\*.*")
        For i = 0 To nFiles - 1
            If Dir(sFile & "\" & tvFiles.Nodes(i).Text) Then Kill sFile & "\" & tvFiles.Nodes(i).Text
            ExtractFile sFile & "\" & tvFiles.Nodes(i).Text, tvFiles.Nodes(i).Text
        Next
        SetStatusBar CStr(nFiles) & " files extracted."
    Case "renfile"
        'Rename selected file
        If tvFiles.SelectedItem Is Nothing Then Exit Sub
        sTitle = InputBox("Please enter the new name you want for the file below:", "Rename File", tvFiles.SelectedItem.Text)
        If sTitle = "" Then Exit Sub
        'Check to be sure that another file doesn't have the same name
        For i = 0 To nFiles - 1
            If (volRecs(i).isRenamed = True And volRecs(i).strNewName = sTitle) Or (volRecs(i).strName = sTitle) Then
                Beep
                SetStatusBar "The name " & sTitle & " is already used by another file."
                Exit Sub
            End If
        Next
        volRecs(CLng(tvFiles.SelectedItem.Key)).strNewName = sTitle
        volRecs(CLng(tvFiles.SelectedItem.Key)).isRenamed = True
        tvFiles.SelectedItem.Text = sTitle
        SetStatusBar sTitle & " renamed successfully."
End Select
End Sub

Public Sub SaveAutosave()
Beep: SetStatusBar "TODO: Implement SaveAutosave()!"
End Sub

