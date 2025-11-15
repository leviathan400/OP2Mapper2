VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrtFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const DIB_RGB_COLORS = 0
Private Const BI_RGB = 0&
Private Const BI_BITFIELDS = 3&

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long



' Structures for the image port of the .prt file (second section)

Private Type IMAGEINFO
    ScanLineByteWidth As Long
    DataPtr As Long
    Height As Long
    Width As Long
    Type As Integer
    PaletteNum As Integer
End Type
' The field "Type" has the following values/meanings
' 0 - menu graphic
' 1 - in game graphic
' 2
' 3
' 4 - Unit shadows - (1 bpp)
' 5 - Unit shadows - (1 bpp)
' ... more, many more


' Structures for the Animation section of the .prt file (third section)

Private Type RECT
    Left As Long
    TOp As Long
    Right As Long
    Bottom As Long
End Type

' Animation header as seen in the file.
' This data is follow by the data for each frame
Private Type AnimationHeaderInfo
    unknown As Long             ' Set to 1 **TODO** Find out use
    SelectionBox As RECT
    PixelXDisplacement As Long
    PixelYDisplacement As Long
    unknown2 As Long            ' Set to 0 (usually) **TODO** Find out use
    NumAnimationFrames As Long
End Type

Private Type AnimationAppendixInfo  ' Name this better? **TODO**
    unknown1 As Long
    unknown2 As Long
    unknown3 As Long
    unknown4 As Long
End Type

Private Type FrameHeaderInfo
    numFrameComponents As Byte
    unknown As Byte
End Type


Private Type FrameOptionalInfo
    unknown1 As Integer
    unknown2 As Integer
End Type

Private Type FrameComponentInfo
    ImageInfoIndex As Integer   ' Index in the image() array for this component
    Unused As Byte              ' Set to 0xFF (Unused/Ignored by Outpost2)
    FrameComponentIndex As Byte ' ??? What's the point?
    PixelXOffset As Integer     ' Offset to left edge of image
    PixelYOffset As Integer     ' Offset to top edge of image
End Type


' Structures to hold the animation data in memory

Private Type AnimationInfo
    Header As AnimationHeaderInfo
    StartingFrameIndex As Long
    NumUnknown As Long
    StartingUnknownIndex As Long
End Type

Private Type FrameInfo
    Header As FrameHeaderInfo
    Optional As FrameOptionalInfo
    StartingFrameComponentIndex As Long
End Type



' PRT data
Private numLoadedPalettes As Long           ' Total number of palettes for all images
Private Palette() As RGBQUAD                ' Array of all palettes needed for the images
Private playerColors() As RGBQUAD           ' Player color palette array
Private numLoadedImages As Long             ' Total number of stored images
Private images() As IMAGEINFO               ' Array of image info structs describing all images in this file
Private numLoadedAnimations As Long         ' Number of animation sequences
Private numLoadedFrames As Long             ' Total number of frames in all animations
Private numLoadedFrameComponents As Long    ' Total number of all components of all frames of all animations
Private numLoadedFrameOptional As Long      ' Additional info attached to select frames
Private animations() As AnimationInfo       ' Array describing all animation sequences
Private frames() As FrameInfo               ' Array describing all frames of animation
Private frameComponents() As FrameComponentInfo ' Array describing all components of each fram of animation

' Bitmap data
Private bmFileHeader As BITMAPFILEHEADER
Private bmInfo As BITMAPINFO
Private bitmapData() As Byte

' Palette data
Private Type SectionHeader
    tag As Long
    sectionSize As Long
End Type

Public Function ReadPaletteData(palId As Long, stream As StreamReader) As Long
    Dim numTagsLeft As Long
    Dim sectHead As SectionHeader
    Dim i As Long, temp As Long, buff As String

    ' Initialize tag count
    numTagsLeft = -2
    ' Read all needed tags
    Do
        ' Decrease the tag count
        numTagsLeft = numTagsLeft - 1

        ' Read in the section header
        stream.Read Len(sectHead), VarPtr(sectHead)

        ' Determine which type of section it was
        Select Case sectHead.tag
        Case 1279348816 ' "PPAL"
            ' Error check the section size
            If sectHead.sectionSize <> 1048 Then
                Debug.Print "Palette: Warning! 'PPAL' section size is not 1048"
            End If
        Case 1179011410 ' "RIFF"
            Debug.Print "Palette: Warning! 'RIFF' tag unhandled"
        Case 1684104552 ' "head"
            ' Error check the section size
            If sectHead.sectionSize <> 4 Then
                Debug.Print "Palette: Warning! 'head' section size is not 4"
            End If
            ' Read numTagsLeft (4 bytes of data)
            stream.Read 4, VarPtr(numTagsLeft)
        Case 1635017060 ' "data"
            ' Error check the section size
            If sectHead.sectionSize <> 1024 Then
                Debug.Print "Palette: Warning! 'data' section size is not 1024"
            End If
            ' Read the palette data
            For i = 0 To 255
                stream.Read 4, VarPtr(Palette(palId, i))
            Next
        Case Else
            Debug.Print "Palette: Warning! Unrecognized section tag: " + sectHead.tag
            ReadPaletteData = -1    ' Error
            Exit Function
        End Select
    Loop While numTagsLeft <> 0
    
    ' Reverse the color components of the palette data
    ' (For some reason, the custom Outpost 2 format stores
    '  it backwards from the standard Windows format)
    For i = 0 To 255
        temp = Palette(palId, i).rgbBlue
        Palette(palId, i).rgbBlue = Palette(palId, i).rgbRed
        Palette(palId, i).rgbRed = temp
    Next
    
    ' Return success
    ReadPaletteData = 0
End Function

Public Function OpenBitmapFile(strFileName As String) As Long
    Dim lngFileNum As Long

    ' Check if the file exists
    If Len(Dir(strFileName)) = 0 Then
        ' Error, file doesn't exist
        OpenBitmapFile = -1
        Exit Function
    End If
    
    ' Open the bitmap file
    lngFileNum = FreeFile
    Open strFileName For Binary As #lngFileNum
    
    ' Read in the bitmap data
    OpenBitmapFile = ReadBitmapData(lngFileNum)
    
    ' Close the bitmap file
    Close #lngFileNum
End Function

Public Function OpenColorFile(stream As StreamReader) As Long
    ' Read in the bitmap data
    OpenColorFile = ReadColorData(stream)
End Function

Public Function ReadBitmapData(lngFileNum As Long) As Long
    ' Read the bitmap header
    Get #lngFileNum, , bmFileHeader
    Get #lngFileNum, , bmInfo

    ' Validate file size
    If LOF(lngFileNum) <= 0 Then
        ReadBitmapData = -1
        Exit Function
    End If
    
    ' Allocate space for the bitmap pixel data
    ReDim bitmapData(LOF(lngFileNum) - Loc(lngFileNum)) As Byte
    ' Load the bitmap pixel data
    Get #lngFileNum, , bitmapData()

    ' Return success
    ReadBitmapData = 0
End Function

Public Function ReadColorData(stream As StreamReader) As Long
    Dim i As Long, j As Long
    
    ' Read the bitmap header
    stream.Read Len(bmFileHeader), VarPtr(bmFileHeader)
    stream.Read Len(bmInfo), VarPtr(bmInfo)
    
    ' Copy player palettes
    ReDim playerColors(6, 24) As RGBQUAD
    
    For i = 0 To 6 '6 players, yes? ;)
        For j = 0 To 24 '24 colors
            playerColors(i, j) = bmInfo.bmiColors(i * 24 + j)
        Next
    Next
    
    ' Return success
    ReadColorData = 0
End Function

Public Function OpenPrtFile(stream As StreamReader) As Long
    ' Read in the PRT data (and assign return code)
    OpenPrtFile = ReadPrtData(stream)
End Function

Public Function ReadPrtData(stream As StreamReader) As Long
    Dim i As Long, j As Long, k As Long
    Dim theTag As Long

    ' Read in the palette section
    ' ***************************

    ' Read the palette section header tag
    stream.Read 4, VarPtr(theTag)
    If theTag <> 1279348803 Then ' "CPAL"
        ReadPrtData = -1    ' Return failure
        Exit Function
    End If

    ' Read the number of palettes
    stream.Read 4, VarPtr(numLoadedPalettes)
    
    ' Allocate the palette array
    ReDim Palette(numLoadedPalettes - 1, 255) As RGBQUAD
    ' Read in all the palette info
    For i = 0 To numLoadedPalettes - 1
        ' Read in the palette data
        If ReadPaletteData(i, stream) <> 0 Then
            ' Error reading in palette data. Abort.
            Debug.Print "CPrtFile: Error reading in palette #" + Format$(i)
            ReadPrtData = -1
            Exit Function
        End If
    Next i
    
    ' Read in the images info
    ' ***********************
    
    ' Read in the number of image info structs
    stream.Read 4, VarPtr(numLoadedImages)
    
    ' Allocate array to hold image structs
    ReDim images(numLoadedImages - 1) As IMAGEINFO
    ' Read in the image info structs
    stream.Read Len(images(0)) * numLoadedImages, VarPtr(images(0))
    
    ' Read in the Animation info
    ' ***********************
    
    ' Read in the structure counts
    stream.Read 4, VarPtr(numLoadedAnimations)
    stream.Read 4, VarPtr(numLoadedFrames)
    stream.Read 4, VarPtr(numLoadedFrameComponents)
    stream.Read 4, VarPtr(numLoadedFrameOptional)
    
    ' Allocate space for the data
    ReDim animations(numLoadedAnimations - 1)
    ReDim frames(numLoadedFrames - 1)
    ReDim frameComponents(numLoadedFrameComponents - 1)
    
    Dim lngCurrentFrame As Long
    Dim lngCurrentFrameComponent As Long
    
    ' Read in the Animation data
    For i = 0 To numLoadedAnimations - 1
        ' Read the animation header
        stream.Read Len(animations(i).Header), VarPtr(animations(i).Header)
        
        ' Set the starting frame index
        animations(i).StartingFrameIndex = lngCurrentFrame

        ' Read in all frame info
        For j = 0 To animations(i).Header.NumAnimationFrames - 1
            ' Read in all frame component info
            stream.Read Len(frames(lngCurrentFrame).Header), VarPtr(frames(lngCurrentFrame).Header)
            
            ' Check if first block of optional data needs to be read in
            If (frames(lngCurrentFrame).Header.numFrameComponents And &H80) = &H80 Then
                ' Read in the optional data
                stream.Read 2, VarPtr(frames(lngCurrentFrame).Optional.unknown1)
            End If
            ' Check if second block of optional data needs to be read in
            If (frames(lngCurrentFrame).Header.unknown And &H80) = &H80 Then
                ' Read in the optional data
                stream.Read 2, VarPtr(frames(lngCurrentFrame).Optional.unknown2)
            End If
            
            ' Set the starting frame component index
            frames(lngCurrentFrame).StartingFrameComponentIndex = lngCurrentFrameComponent
            
            For k = 0 To (frames(lngCurrentFrame).Header.numFrameComponents And &H7F) - 1
                ' Read the frame component
                stream.Read Len(frameComponents(lngCurrentFrameComponent)), VarPtr(frameComponents(lngCurrentFrameComponent))
                lngCurrentFrameComponent = lngCurrentFrameComponent + 1
            Next k
            lngCurrentFrame = lngCurrentFrame + 1
        Next j
        
        ' Read number of (appedicies?)
        stream.Read 4, VarPtr(animations(i).NumUnknown)

        ' Read in all (appendix?) data
        For j = 0 To animations(i).NumUnknown - 1
            ' Read in the (appendix?) structure (and forget it! ;)  **TODO** Improve
            Dim temp As AnimationAppendixInfo
            stream.Read Len(temp), VarPtr(temp)
        Next j
    Next i


    ' Check if the correct amount of data was loaded
    If numLoadedFrames <> lngCurrentFrame Then
        Debug.Print "Incorrect number of frames loaded"
    End If
    If numLoadedFrameComponents <> lngCurrentFrameComponent Then
        Debug.Print "Incorrect number of frame components loaded"
    End If

    ' Return success
    ReadPrtData = 0
End Function


' Image access functions
' **********************

Public Property Get numImages() As Long
    numImages = numLoadedImages
End Property

Public Property Get ImageWidth(imageNum As Long) As Long
    ImageWidth = images(imageNum).Width
End Property

Public Property Get ImageHeight(imageNum As Long) As Long
    ImageHeight = images(imageNum).Height
End Property

Public Property Get ImageType(imageNum As Long) As Long
    ImageType = images(imageNum).Type
End Property

' Animation access functions
' **************************
' **TODO** Finish writing these/add more

Public Property Get NumAnimations() As Long
    NumAnimations = numLoadedAnimations
End Property

Public Property Get NumFrames(animationIndex As Long) As Long
    NumFrames = animations(animationIndex).Header.NumAnimationFrames
End Property

' Pasting functions
' *****************

Public Function Paste(ByVal destDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal imageNum As Long, ByVal playerNum As Long) As Long
    Dim i As Long
    
    ' Fill in temp fields
    bmInfo.bmiHeader.biWidth = images(imageNum).Width
    bmInfo.bmiHeader.biHeight = -images(imageNum).Height
    
    ' **TODO** Fix this cheap hack to adjust bit depth
    If images(imageNum).Type < 4 Then
        bmInfo.bmiHeader.biBitCount = 8
    ElseIf images(imageNum).Type > 5 Then
        bmInfo.bmiHeader.biBitCount = 8
    Else
        bmInfo.bmiHeader.biBitCount = 1
    End If

    ' Fill in the player palette
    For i = 0 To 23
        bmInfo.bmiColors(i) = playerColors(playerNum, i)
    Next
    ' Fill in the rest of the palette
    For i = 24 To 255
        bmInfo.bmiColors(i) = Palette(images(imageNum).PaletteNum, i)
    Next

    ' Paste the image
    Paste = SetDIBitsToDevice(destDC, destX, destY, images(imageNum).Width, images(imageNum).Height, 0, 0, 0, images(imageNum).Height, bitmapData(images(imageNum).DataPtr), bmInfo, DIB_RGB_COLORS)
End Function

Public Function PasteMask(ByVal destDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal imageNum As Long) As Long
    'Same as paste, but it pastes a black and white bitmap for use as a mask
    Dim i As Long
    
    ' Fill in temp fields
    bmInfo.bmiHeader.biWidth = images(imageNum).Width
    bmInfo.bmiHeader.biHeight = -images(imageNum).Height
    
    ' **TODO** Fix this cheap hack to adjust bit depth
    If images(imageNum).Type < 4 Then
        bmInfo.bmiHeader.biBitCount = 8
    ElseIf images(imageNum).Type > 5 Then
        bmInfo.bmiHeader.biBitCount = 8
    Else
        bmInfo.bmiHeader.biBitCount = 1
    End If

    ' Fill in the palette (first index is trans, others are not)
    bmInfo.bmiColors(0).rgbBlue = &HFF
    bmInfo.bmiColors(0).rgbGreen = &HFF
    bmInfo.bmiColors(0).rgbRed = &HFF
    For i = 1 To 255
        bmInfo.bmiColors(i).rgbBlue = 0
        bmInfo.bmiColors(i).rgbGreen = 0
        bmInfo.bmiColors(i).rgbRed = 0
    Next

    ' Paste the mask image
    PasteMask = SetDIBitsToDevice(destDC, destX, destY, images(imageNum).Width, images(imageNum).Height, 0, 0, 0, images(imageNum).Height, bitmapData(images(imageNum).DataPtr), bmInfo, DIB_RGB_COLORS)
End Function

Public Sub PasteFrame(ByVal destDC As Long, ByVal destX As Long, ByVal destY As Long, ByVal frameNum As Long, ByVal playerNum As Long)
    'Paste a whole frame to the screen somewhere
    Dim i As Long
    'Go thru each frame component and draw it to the screen
    For i = frames(frameNum).StartingFrameComponentIndex To frames(frameNum).StartingFrameComponentIndex + (frames(frameNum).Header.numFrameComponents And &H7F) - 1
        'Exclude the shadows from the final drawing (shadows are type 4 and 5)
        If images(frameComponents(i).ImageInfoIndex).Type <> 4 And images(frameComponents(i).ImageInfoIndex).Type <> 5 Then
            Paste destDC, frameComponents(i).PixelXOffset + destX, frameComponents(i).PixelYOffset + destY, frameComponents(i).ImageInfoIndex, playerNum
        End If
    Next
End Sub

Public Sub PasteAnim(ByVal destDC As Long, ByVal destX1 As Long, ByVal destY1 As Long, ByVal destX2 As Long, ByVal destY2 As Long, ByVal animNum As Long, ByVal playerNum As Long)
    'Paste the last frame of a building animation
    Dim frame As Long
    Dim xDisp As Long
    Dim yDisp As Long
    
    xDisp = ((destX2 - destX1) \ 2) - animations(animNum).Header.PixelXDisplacement
    yDisp = ((destY2 - destY1) \ 2) - animations(animNum).Header.PixelYDisplacement
    frame = animations(animNum).StartingFrameIndex + animations(animNum).Header.NumAnimationFrames - 1
    'The tile being pasted to is the center. So add the displacement data
    PasteFrame destDC, destX1 + xDisp, destY1 + yDisp, frame, playerNum
End Sub

Public Sub PasteAnimTrans(ByVal destDC As Long, ByVal destX1 As Long, ByVal destY1 As Long, ByVal destX2 As Long, ByVal destY2 As Long, ByVal animNum As Long, ByVal playerNum As Long)
    On Error Resume Next
    'Same as PasteAnim, but treat the first index in the palette as the transparent color
    Dim frame As Long
    Dim xDisp As Long
    Dim yDisp As Long
    
    xDisp = ((destX2 - destX1) \ 2) - animations(animNum).Header.PixelXDisplacement
    yDisp = ((destY2 - destY1) \ 2) - animations(animNum).Header.PixelYDisplacement
    frame = animations(animNum).StartingFrameIndex + animations(animNum).Header.NumAnimationFrames - 1
    
    'Create a temp DC
    Dim tmpDC As Long, tmpBm As Long, garbage As Long
    tmpDC = CreateCompatibleDC(destDC)
    
    Dim i As Long
    'Go thru each frame component and draw its mask to the screen
    For i = frames(frame).StartingFrameComponentIndex To frames(frame).StartingFrameComponentIndex + (frames(frame).Header.numFrameComponents And &H7F) - 1
        'Exclude the shadows from the final drawing (shadows are type 4 and 5)
        If images(frameComponents(i).ImageInfoIndex).Type <> 4 And images(frameComponents(i).ImageInfoIndex).Type <> 5 Then
            'Setup the mask bitmap
            tmpBm = CreateCompatibleBitmap(destDC, images(frameComponents(i).ImageInfoIndex).Width, images(frameComponents(i).ImageInfoIndex).Height)
            garbage = SelectObject(tmpDC, tmpBm)
            
            PasteMask tmpDC, 0, 0, frameComponents(i).ImageInfoIndex
            'Blit this with SRCAND (for masking)
            BitBlt destDC, frameComponents(i).PixelXOffset + destX1 + xDisp, frameComponents(i).PixelYOffset + destY1 + yDisp, images(frameComponents(i).ImageInfoIndex).Width, images(frameComponents(i).ImageInfoIndex).Height, tmpDC, 0, 0, vbSrcAnd
            
            'Now paste the real deal and blit it with OR (SRCPAINT)
            Paste tmpDC, 0, 0, frameComponents(i).ImageInfoIndex, playerNum
            BitBlt destDC, frameComponents(i).PixelXOffset + destX1 + xDisp, frameComponents(i).PixelYOffset + destY1 + yDisp, images(frameComponents(i).ImageInfoIndex).Width, images(frameComponents(i).ImageInfoIndex).Height, tmpDC, 0, 0, vbSrcPaint
            
            'Destroy the crap bitmap
            SelectObject tmpDC, garbage
            DeleteObject tmpBm
        End If
    Next
    
    'Destroy the crap DC
    DeleteDC tmpDC
End Sub

Public Function GetPlayerColor(ByVal playerId As Long) As Long
    On Error Resume Next
    GetPlayerColor = RGB(playerColors(playerId, 19).rgbRed, playerColors(playerId, 19).rgbGreen, playerColors(playerId, 19).rgbBlue)
End Function
