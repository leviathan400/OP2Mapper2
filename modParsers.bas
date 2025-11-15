Attribute VB_Name = "modParsers"
Option Explicit

'Template parser is a finite state machine: these are the possible states
Public Enum UnitTplMode
    InvalidTplMode
    ReadProlog 'reading the code prolog
    ReadEpilog 'reading the code epilog
    ReadPlayerProlog 'reading player prolog
    ReadPlayerEpilog 'reading player epilog
    ReadUnit 'reading unit lines
    ReadBeacon 'reading beacon lines
    ReadWall 'reading wall lines
    ReadWreck 'reading wreck lines
End Enum

Public Function ParseUnitTable(strFileName As String) As Boolean
On Error GoTo TableProblems
'Returns true on success
Dim f As Integer, tmpLine As String, inQuote As Boolean
Dim tmpUnit As UnitType, buildStr As String, i As Long, curField As Integer
Dim lineNum As Long
f = FreeFile
Open strFileName For Input As #f
'Loop thru the file, reading a line at a time
lineNum = 1
Do Until EOF(f)
    'Init default values
    tmpUnit.mapID = 0
    tmpUnit.artID = -1
    tmpUnit.strName = ""
    tmpUnit.isStructure = False
    tmpUnit.sizeX = 0
    tmpUnit.sizeY = 0
    tmpUnit.hasTubes = False
    tmpUnit.verTubeLoc = 0
    tmpUnit.horTubeLoc = 0
    tmpUnit.canHaveTurret = False
    inQuote = False
    buildStr = ""
    curField = 0
    
    'Read a line
    Line Input #f, tmpLine
    'Clean up the line
    tmpLine = Trim$(tmpLine)
    tmpLine = Replace(tmpLine, vbTab, "", , , vbTextCompare)
    
    '---SPECIAL CASES---
    'Is the first character ; (comment)?
    'If so, skip the line
    If Left$(tmpLine, 1) = ";" Then GoTo NextLine
    
    'Does the line start with INCLUDE?
    'If so, read the filename and call ParseUnitTable recursively
    If UCase$(Left$(tmpLine, 7)) = "INCLUDE" Then
        For i = 8 To Len(tmpLine)
            If inQuote = False Then
                If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                    inQuote = True
                End If
            Else
                'In quotes now
                If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                    inQuote = False
                    'Call this function recursively
                    ParseUnitTable App.Path & "\" & buildStr
                    Exit For
                Else
                    'Copy current character into the string
                    buildStr = buildStr & Mid$(tmpLine, i, 1)
                End If
            End If
        Next
        GoTo NextLine
    End If
    
    'Is the line empty?
    If Len(tmpLine) = 0 Then GoTo NextLine
    
    inQuote = False
    buildStr = ""
    
    'Add a comma on the end of the string, so the last value gets copied
    tmpLine = tmpLine & ","
    
    'Read the line character by character and scan in values
    For i = 1 To Len(tmpLine)
        If inQuote = False Then
            'Not inside quotes so characters other than numeric, quote and comma are not allowed
            '---Handle Numbers---
            If IsNumeric(Mid$(tmpLine, i, 1)) Then
                'add it to the temporary string
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            ElseIf Mid$(tmpLine, i, 1) = "-" Then
                'add it to the temporary string
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            ElseIf Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                inQuote = True
            ElseIf Mid$(tmpLine, i, 1) = "," Then 'comma
                'copy the buildStr into whatever is wanted right now
                Select Case curField
                    Case 0 'mapID
                        tmpUnit.mapID = CLng(buildStr)
                    'Case 1 is handled below
                    Case 2 'sizeX
                        'store in ART ID, this could be an art id if it's a vehicle
                        tmpUnit.artID = CLng(buildStr)
                    Case 3 'sizeY
                        tmpUnit.isStructure = True
                        tmpUnit.sizeX = tmpUnit.artID
                        tmpUnit.sizeY = CLng(buildStr)
                        tmpUnit.artID = -1
                    Case 4 'verTube
                        'tmpUnit.hasTubes = True
                        'tmpUnit.verTubeLoc = CLng(buildStr)
                        'If the horizontal tube data never appears, the unit
                        ''never has' tubes so this becomes the art id
                        tmpUnit.artID = CLng(buildStr)
                    Case 5 'horTube
                        tmpUnit.hasTubes = True
                        tmpUnit.horTubeLoc = CLng(buildStr)
                        'The unit has tubes now; so the 'art id' is really the V tube data
                        tmpUnit.verTubeLoc = tmpUnit.artID
                        tmpUnit.artID = -1
                    Case 6 'artID
                        tmpUnit.artID = CLng(buildStr)
                End Select
                buildStr = ""
                curField = curField + 1
                If curField = 7 Then Exit For 'all done with this line
            ElseIf Mid$(tmpLine, i, 1) = " " Then 'space
                'error
                ParseUnitTable = False
                GenerateError "Line " & CStr(lineNum) & ": Spaces are not allowed in non-string table fields", "Unit Table Parser: " & strFileName
                Close #f: Exit Function
            Else
                'error
                ParseUnitTable = False
                GenerateError "Line " & CStr(lineNum) & ": Unrecognized character", "Unit Table Parser: " & strFileName
                Close #f: Exit Function
            End If
        Else 'inQuote=True
            If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                inQuote = False
                If i < Len(tmpLine) Then
                    If Mid$(tmpLine, i + 1, 1) = "," Then i = i + 1
                End If
                If curField <> 1 Then
                    'error
                    ParseUnitTable = False
                    GenerateError "Line " & CStr(lineNum) & ": String where non-string expected", "Unit Table Parser: " & strFileName
                    Close #f: Exit Function
                Else
                    tmpUnit.strName = buildStr
                    buildStr = ""
                    curField = curField + 1
                End If
            Else
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            End If
        End If
    Next
    'Copy the unit structure into the master array
    numUnitDefs = numUnitDefs + 1
    ReDim Preserve allUnitDefs(numUnitDefs)
    allUnitDefs(numUnitDefs) = tmpUnit

NextLine:
    lineNum = lineNum + 1
Loop
Close #f
ParseUnitTable = True
Exit Function
TableProblems:
ParseUnitTable = False
GenerateError "Line " & CStr(lineNum) & ": Intrinsic error #" & Err.Number & " occurred", "Unit Table Parser: " & strFileName
Close #f
End Function

Public Function ParseWeaponTable(strFileName As String) As Boolean
On Error GoTo TableProblems
'Returns true on success
Dim f As Integer, tmpLine As String, inQuote As Boolean
Dim tmpWeapon As WeaponType, buildStr As String, i As Long, curField As Integer
Dim lineNum As Long
f = FreeFile
Open strFileName For Input As #f
'Loop thru the file, reading a line at a time
lineNum = 1
Do Until EOF(f)
    tmpWeapon.mapID = 0
    tmpWeapon.strName = ""
    tmpWeapon.artID = -1
    inQuote = False
    buildStr = ""
    curField = 0
    
    'Read a line
    Line Input #f, tmpLine
    'Clean up the line
    tmpLine = Trim$(tmpLine)
    tmpLine = Replace(tmpLine, vbTab, "", , , vbTextCompare)
    
    '---SPECIAL CASES---
    'Is the first character ; (comment)?
    'If so, skip the line
    If Left$(tmpLine, 1) = ";" Then GoTo NextLine
    
    'Does the line start with INCLUDE?
    'If so, read the filename and call ParseWeaponTable recursively
    If UCase$(Left$(tmpLine, 7)) = "INCLUDE" Then
        For i = 8 To Len(tmpLine)
            If inQuote = False Then
                If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                    inQuote = True
                End If
            Else
                'In quotes now
                If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                    inQuote = False
                    'Call this function recursively
                    ParseWeaponTable App.Path & "\" & buildStr
                    Exit For
                Else
                    'Copy current character into the string
                    buildStr = buildStr & Mid$(tmpLine, i, 1)
                End If
            End If
        Next
        GoTo NextLine
    End If
    
    'Does the line start with HASTURRET?
    'If so, set the boolean in the weapon record
    If UCase$(Left$(tmpLine, 9)) = "HASTURRET" Then
        'Parse the digits after it
        For i = 10 To Len(tmpLine)
            If IsNumeric(Mid$(tmpLine, i, 1)) Then
                'add it to the temporary string
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            End If
        Next
        'Now scan the unit table for a corresponding entry and set it's fields accordingly
        For i = 0 To numUnitDefs
            If allUnitDefs(i).mapID = CLng(buildStr) Then
                allUnitDefs(i).canHaveTurret = True
                GoTo NextLine
            End If
        Next
        'Execution should not reach here, if it does then the user entered a non-existant mapid
        ParseWeaponTable = False
        GenerateError "Line " & CStr(lineNum) & ": Unit ID " & CLng(buildStr) & " does not exist", "Weapon Table Parser: " & strFileName
        Close #f: Exit Function
    End If
    
    'Is the line empty?
    If Len(tmpLine) = 0 Then GoTo NextLine
    
    inQuote = False
    buildStr = ""
    
    'Add a comma on the end of the string, so the last value gets copied
    tmpLine = tmpLine & ","
    
    'Read the line character by character and scan in values
    For i = 1 To Len(tmpLine)
        If inQuote = False Then
            'Not inside quotes so characters other than numeric, quote and comma are not allowed
            '---Handle Numbers---
            If IsNumeric(Mid$(tmpLine, i, 1)) Then
                'add it to the temporary string
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            ElseIf Mid$(tmpLine, i, 1) = "-" Then
                'add it to the temporary string
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            ElseIf Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                inQuote = True
            ElseIf Mid$(tmpLine, i, 1) = "," Then 'comma
                'copy the buildStr into whatever is wanted right now
                Select Case curField
                    Case 0 'mapID
                        tmpWeapon.mapID = CLng(buildStr)
                    'Case 1 is handled below
                    Case 2 'artID
                        tmpWeapon.artID = CLng(buildStr)
                End Select
                buildStr = ""
                curField = curField + 1
                If curField = 3 Then Exit For 'all done with this line
            ElseIf Mid$(tmpLine, i, 1) = " " Then 'space
                'error
                ParseWeaponTable = False
                GenerateError "Line " & CStr(lineNum) & ": Spaces are not allowed in non-string table fields", "Weapon Table Parser: " & strFileName
                Close #f: Exit Function
            Else
                'error
                ParseWeaponTable = False
                GenerateError "Line " & CStr(lineNum) & ": Unrecognized character", "Weapon Table Parser: " & strFileName
                Close #f: Exit Function
            End If
        Else 'inQuote=True
            If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                inQuote = False
                If i < Len(tmpLine) Then
                    If Mid$(tmpLine, i + 1, 1) = "," Then i = i + 1
                End If
                If curField <> 1 Then
                    'error
                    ParseWeaponTable = False
                    GenerateError "Line " & CStr(lineNum) & ": String where non-string expected", "Weapon Table Parser: " & strFileName
                    Close #f: Exit Function
                Else
                    tmpWeapon.strName = buildStr
                    buildStr = ""
                    curField = curField + 1
                End If
            Else
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            End If
        End If
    Next
    'Copy the weapon structure into the master array
    numWeaponDefs = numWeaponDefs + 1
    ReDim Preserve allWeaponDefs(numWeaponDefs)
    allWeaponDefs(numWeaponDefs) = tmpWeapon

NextLine:
    lineNum = lineNum + 1
Loop
Close #f
ParseWeaponTable = True
Exit Function
TableProblems:
ParseWeaponTable = False
GenerateError "Line " & CStr(lineNum) & ": Intrinsic error #" & Err.Number & " occurred", "Weapon Table Parser: " & strFileName
Close #f
End Function

Public Function ParseObjectTable(strFileName As String) As Boolean
On Error GoTo TableProblems
'Returns true on success
Dim f As Integer, tmpLine As String, inQuote As Boolean
Dim tmpUnit As UnitType, buildStr As String, i As Long, curField As Integer
Dim lineNum As Long
f = FreeFile
Open strFileName For Input As #f
'Loop thru the file, reading a line at a time
lineNum = 1
Do Until EOF(f)
    'Init default values
    tmpUnit.mapID = 0
    tmpUnit.artID = -1
    tmpUnit.strName = ""
    tmpUnit.isGaia = True
    tmpUnit.gaiaType = -1
    tmpUnit.extra1 = -1
    tmpUnit.extra2 = -1
    tmpUnit.extra3 = -1
    
    inQuote = False
    buildStr = ""
    curField = 0
    
    'Read a line
    Line Input #f, tmpLine
    'Clean up the line
    tmpLine = Trim$(tmpLine)
    tmpLine = Replace(tmpLine, vbTab, "", , , vbTextCompare)
    
    '---SPECIAL CASES---
    'Is the first character ; (comment)?
    'If so, skip the line
    If Left$(tmpLine, 1) = ";" Then GoTo NextLine
    
    'Does the line start with INCLUDE?
    'If so, read the filename and call ParseObjectTable recursively
    If UCase$(Left$(tmpLine, 7)) = "INCLUDE" Then
        For i = 8 To Len(tmpLine)
            If inQuote = False Then
                If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                    inQuote = True
                End If
            Else
                'In quotes now
                If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                    inQuote = False
                    'Call this function recursively
                    ParseObjectTable App.Path & "\" & buildStr
                    Exit For
                Else
                    'Copy current character into the string
                    buildStr = buildStr & Mid$(tmpLine, i, 1)
                End If
            End If
        Next
        GoTo NextLine
    End If
    
    'Is the line empty?
    If Len(tmpLine) = 0 Then GoTo NextLine
    
    inQuote = False
    buildStr = ""
    
    'Add a comma on the end of the string, so the last value gets copied
    tmpLine = tmpLine & ","
    
    'Read the line character by character and scan in values
    For i = 1 To Len(tmpLine)
        If inQuote = False Then
            'Not inside quotes so characters other than numeric, quote and comma are not allowed
            '---Handle Numbers and Negative Char---
            If IsNumeric(Mid$(tmpLine, i, 1)) Then
                'add it to the temporary string
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            ElseIf Mid$(tmpLine, i, 1) = "-" Then
                'add it to the temporary string
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            ElseIf Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                inQuote = True
            ElseIf Mid$(tmpLine, i, 1) = "," Then 'comma
                'copy the buildStr into whatever is wanted right now
                Select Case curField
                    Case 0 'mapID
                        tmpUnit.mapID = CLng(buildStr)
                    'Case 1 is handled below
                    Case 2 'type
                        'store in ART ID, this could be an art id if it's a vehicle
                        tmpUnit.gaiaType = CLng(buildStr)
                    Case 3 'artID
                        tmpUnit.artID = CLng(buildStr)
                    Case 4 'extra1
                        tmpUnit.extra1 = CLng(buildStr)
                    Case 5 'extra2
                        tmpUnit.extra2 = CLng(buildStr)
                    Case 6 'extra3
                        tmpUnit.extra3 = CLng(buildStr)
                End Select
                buildStr = ""
                curField = curField + 1
                If curField = 7 Then Exit For 'all done with this line
            ElseIf Mid$(tmpLine, i, 1) = " " Then 'space
                'error
                ParseObjectTable = False
                GenerateError "Line " & CStr(lineNum) & ": Spaces are not allowed in non-string table fields", "Object Table Parser: " & strFileName
                Close #f: Exit Function
            Else
                'error
                ParseObjectTable = False
                GenerateError "Line " & CStr(lineNum) & ": Unrecognized character", "Object Table Parser: " & strFileName
                Close #f: Exit Function
            End If
        Else 'inQuote=True
            If Asc(Mid$(tmpLine, i, 1)) = 34 Then 'Quote "" marks
                inQuote = False
                If i < Len(tmpLine) Then
                    If Mid$(tmpLine, i + 1, 1) = "," Then i = i + 1
                End If
                If curField <> 1 Then
                    'error
                    ParseObjectTable = False
                    GenerateError "Line " & CStr(lineNum) & ": String where non-string expected", "Object Table Parser: " & strFileName
                    Close #f: Exit Function
                Else
                    tmpUnit.strName = buildStr
                    buildStr = ""
                    curField = curField + 1
                End If
            Else
                buildStr = buildStr & Mid$(tmpLine, i, 1)
            End If
        End If
    Next
    'Copy the unit structure into the master array
    numUnitDefs = numUnitDefs + 1
    ReDim Preserve allUnitDefs(numUnitDefs)
    allUnitDefs(numUnitDefs) = tmpUnit

NextLine:
    lineNum = lineNum + 1
Loop
Close #f
ParseObjectTable = True
Exit Function
TableProblems:
ParseObjectTable = False
GenerateError "Line " & CStr(lineNum) & ": Intrinsic error #" & Err.Number & " occurred", "Object Table Parser: " & strFileName
Close #f
End Function


Public Sub ParseCommandLine()
'There is something on the command line to be read. Like a file name
Dim theCmd As String, cmdArr() As String, i As Variant
theCmd = Trim$(Command)
'cmdArr = Split(theCmd, " ", , vbTextCompare)
'For Each i In cmdArr

'**TODO** Figure out how to parse names with spaces better, so other
'command line options can be supported in future
i = theCmd
    'File extension? "."
    If Mid$(i, Len(i) - 3, 1) = "." Then
        Select Case LCase$(Right$(i, 3)) 'Decide how to open the file
            Case "map" 'Map file
                fMainForm.OpenMap i
            Case "bmp" 'Tileset file
                '**TODO**
            Case "vol"
                fMainForm.OpenVol i
            Case Else
                'Extension not recognized
                SetStatusBar "Extension ." & LCase$(Right$(i, 3)) & " on the command line was not recognized."
        End Select
    End If
    
'Next

End Sub

Public Function ParseAndGenerateUnits(ByVal templateName As String, unitRecs() As UnitRec, ByVal numUnitRecs As Long, ByVal mapTitle As String) As String
On Error GoTo oops

'Parse a unit template file and generate appropriate code
Dim f As Integer, tmpLine As String, lineNum As Long, buildStr As String
Dim prolog As String, epilog As String 'Master prolog / epilog
Dim plrProlog As String, plrEpilog As String 'player prolog / epilog
Dim unitLine As String 'unit line data
Dim beaconLine As String 'beacon line data
Dim wallLine As String 'wall line data
Dim wreckLine As String 'wreck line data
Dim modeSel As UnitTplMode, numPlayersUsed As Long, numGaiaObjects As Long, i As Long, j As Long
f = FreeFile
modeSel = InvalidTplMode

'Open it, begin reading
Open templateName For Input As #f
lineNum = 1
Do Until EOF(f)
    Line Input #f, tmpLine
    ' Handle various cases if this is a new section
    Select Case LCase$(Trim$(tmpLine))
        Case "$$rem" 'Do nothing, its a comment
            GoTo NextLine
        Case "$$begin" 'Put the parser into prolog state
            modeSel = ReadProlog
        Case "$$end" 'Put the parser into epilog state
            modeSel = ReadEpilog
        Case "$$player" 'Put the parser into player prolog state
            modeSel = ReadPlayerProlog
        Case "$$endplayer" 'Put the parser into player epilog state
            modeSel = ReadPlayerEpilog
        Case "$$unit" 'Put the parser into unit state
            modeSel = ReadUnit
        Case "$$beacon" 'Put the parser into beacon state
            modeSel = ReadBeacon
        Case "$$wall" 'Put the parser into wall state
            modeSel = ReadWall
        Case "$$wreck" 'Put the parser into wreck state
            modeSel = ReadWreck
        Case Else
            'Nothing matched a magic keyword; follow the normal states parsing
            Select Case modeSel
                Case ReadProlog
                    'Read prolog lines
                    prolog = prolog & tmpLine & vbNewLine
                Case ReadEpilog
                    'Read epilog lines
                    epilog = epilog & tmpLine & vbNewLine
                Case ReadPlayerProlog
                    'Read player prolog lines
                    plrProlog = plrProlog & tmpLine & vbNewLine
                Case ReadPlayerEpilog
                    'Read player epilog lines
                    plrEpilog = plrEpilog & tmpLine & vbNewLine
                Case ReadUnit
                    'Read unit lines
                    unitLine = unitLine & tmpLine & vbNewLine
                Case ReadBeacon
                    'Read beacon lines
                    beaconLine = beaconLine & tmpLine & vbNewLine
                Case ReadWall
                    'Read wall lines
                    wallLine = wallLine & tmpLine & vbNewLine
                Case ReadWreck
                    'Read wreck lines
                    wreckLine = wreckLine & tmpLine & vbNewLine
                Case Else
                    'not defined yet -- we'll let it slide with any garbage in here
                    'GenerateError "Parser error: Invalid data for current state on line " & CStr(lineNum), "ParseAndGenerateUnits"
            End Select
    End Select
NextLine: 'skip
lineNum = lineNum + 1
Loop

'Done with the file now
Close #f

'Calculate the number of players used
numPlayersUsed = 1
For i = 0 To 6
    For j = 0 To numUnitRecs - 1
        If unitRecs(j).playerNum = numPlayersUsed And unitRecs(j).uType.isGaia = False Then
            numPlayersUsed = numPlayersUsed + 1
            Exit For
        End If
    Next
Next
'Calculate the number of gaia objects
numGaiaObjects = 0
For j = 0 To numUnitRecs - 1
    If unitRecs(j).uType.isGaia = True Then
        numGaiaObjects = numGaiaObjects + 1
        Exit For
    End If
Next

'Now write out the data

'Write out a small header that tells the stats of the current file
buildStr = buildStr & "// Generated by OP2Mapper " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision) & vbNewLine & _
    "// at " & CStr(Now) & vbNewLine & _
    "// using template " & templateName & vbNewLine & _
    "// Total Number of Objects: " & CStr(numUnitRecs) & vbNewLine & vbNewLine

'Handle prolog - substitute data
tmpLine = prolog
tmpLine = Replace(tmpLine, "$totalunits", CStr(numUnitRecs), , , vbTextCompare)
tmpLine = Replace(tmpLine, "$playersused", CStr(numPlayersUsed), , , vbTextCompare)
tmpLine = Replace(tmpLine, "$template", templateName, , , vbTextCompare)
tmpLine = Replace(tmpLine, "$mapname", mapTitle, , , vbTextCompare)

'No newline, it has been added on previously
buildStr = buildStr & tmpLine

'Handle the Gaia objects first
For j = 0 To numUnitRecs - 1
    If unitRecs(j).uType.isGaia = True Then
        'Handle this unit
        Select Case unitRecs(j).uType.gaiaType
            Case 0 'beacon
                tmpLine = beaconLine
            Case 1 'wall
                tmpLine = wallLine
            Case 2 'wreck
                tmpLine = wreckLine
        End Select
        tmpLine = Replace(tmpLine, "$totalunits", CStr(numUnitRecs), , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$playersused", CStr(numPlayersUsed), , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$template", templateName, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$mapname", mapTitle, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$mapid", unitRecs(j).uType.mapID, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$x", unitRecs(j).locX + 1, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$y", unitRecs(j).locY + 1, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$extra1", unitRecs(j).uType.extra1, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$extra2", unitRecs(j).uType.extra2, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$extra3", unitRecs(j).uType.extra3, , , vbTextCompare)
        tmpLine = Replace(tmpLine, "$gaiaunits", CStr(numGaiaObjects), , , vbTextCompare)
        
        buildStr = buildStr & tmpLine
    End If
Next

'Start with lowest numbered player
For i = 0 To numPlayersUsed - 1
    'Handle this player
    tmpLine = plrProlog
    tmpLine = Replace(tmpLine, "$totalunits", CStr(numUnitRecs), , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$playersused", CStr(numPlayersUsed), , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$template", templateName, , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$mapname", mapTitle, , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$playerid", CStr(i), , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$playerunits", ParserGetNumPlayerUnits(i, unitRecs, numUnitRecs), , , vbTextCompare)
    
    buildStr = buildStr & tmpLine
    
    For j = 0 To numUnitRecs - 1
        If unitRecs(j).playerNum = i And unitRecs(j).uType.isGaia = False Then
            'Handle this unit
            tmpLine = unitLine
            tmpLine = Replace(tmpLine, "$totalunits", CStr(numUnitRecs), , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$playersused", CStr(numPlayersUsed), , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$template", templateName, , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$mapname", mapTitle, , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$playerid", CStr(i), , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$playerunits", ParserGetNumPlayerUnits(i, unitRecs, numUnitRecs), , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$mapid", unitRecs(j).uType.mapID, , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$weaponmapid", unitRecs(j).wType.mapID, , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$x", unitRecs(j).locX + 1, , , vbTextCompare)
            tmpLine = Replace(tmpLine, "$y", unitRecs(j).locY + 1, , , vbTextCompare)

            buildStr = buildStr & tmpLine
        End If
    Next
    
    'Player Epilog
    tmpLine = plrEpilog
    tmpLine = Replace(tmpLine, "$totalunits", CStr(numUnitRecs), , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$playersused", CStr(numPlayersUsed), , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$template", templateName, , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$mapname", mapTitle, , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$playerid", CStr(i), , , vbTextCompare)
    tmpLine = Replace(tmpLine, "$playerunits", ParserGetNumPlayerUnits(i, unitRecs, numUnitRecs), , , vbTextCompare)

    buildStr = buildStr & tmpLine
Next

'Handle epilog
tmpLine = epilog
tmpLine = Replace(tmpLine, "$totalunits", CStr(numUnitRecs), , , vbTextCompare)
tmpLine = Replace(tmpLine, "$playersused", CStr(numPlayersUsed), , , vbTextCompare)
tmpLine = Replace(tmpLine, "$template", templateName, , , vbTextCompare)
tmpLine = Replace(tmpLine, "$mapname", mapTitle, , , vbTextCompare)

buildStr = buildStr & tmpLine

'End it
ParseAndGenerateUnits = buildStr
Exit Function

oops:
ParseAndGenerateUnits = ""
GenerateError "TPL Parser error occurred.", "Parser::ParseAndGenerateUnits"
End Function

Public Function ParserGetNumPlayerUnits(player As Long, unitRecs() As UnitRec, ByVal numUnitRecs As Long) As Long
'figure out the number of units for a player
ParserGetNumPlayerUnits = 0
Dim i As Long
For i = 0 To numUnitRecs - 1
    If unitRecs(i).playerNum = player And unitRecs(i).uType.isGaia = False Then
        ParserGetNumPlayerUnits = ParserGetNumPlayerUnits + 1
    End If
Next
End Function
