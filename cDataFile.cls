VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This module defines a standard data file format which the mapper uses

'The data files are segmented. The beginning of the file is specified as below:
Private Type DataFileHeader
    signature As String * 10 'Should equal the value "MapperData" followed by an EOF byte, followed by a NUL byte
    fileVersion As Integer 'Version number of the file. Currently 200
    numSegs As Integer 'Number of segments in the file
End Type

'Next comes the first segment. This is a header followed by raw data.
Private Type DataSegHeader
    signature As String * 4 'Identifies what type of data
    sequenceId As Byte 'Unused for now. set to 0
    dataSize As Long 'Number of bytes that follow this segment header
End Type

'X number of bytes: segment dependent data
'numSegs - 1 sections follow: headers, then data
'END OF FILE

'Some segment types
'"UNIT" - unit data for maps
'"SCRP" - script data
'"CTLD" - .ctl file contents
'"QUOT" - text records
'"MAPD" - map data
'"VOLD" - vol data
'"TECH" - techtree text file
'"TILS" - tileset
'"UNDO" - undo data
'"GRPS" - tile group data

'"UNKN" - unknown / unused...

Private fileNum As Integer

Public Function OpenDataFile(ByVal strFile As String) As Boolean
    On Error GoTo errhandler
    fileNum = FreeFile
    Open strFile For Binary As #fileNum
    ' get header
    Dim hdr As DataFileHeader
    Get #fileNum, , hdr
    ' check signature
    If hdr.signature <> "MapperData" & Chr$(26) & Chr$(0) Then
        Err.Raise vbObjectError, "cDataFile", "Signature does not match"
        OpenDataFile = False
    End If
    
    ' check version number
    If hdr.fileVersion > 200 Then
        Err.Raise vbObjectError, "cDataFile", "Version too high"
        OpenDataFile = False
    End If
    
    Dim i As Long
    ' now read section info
    
errhandler:
    GenerateError "Could not open data file " & strFile, "cDataFile::OpenDataFile"
End Function
