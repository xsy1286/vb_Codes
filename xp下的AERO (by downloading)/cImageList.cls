VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Review Usage_ImageList for overview of class methods and properties

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' 32 bytes as used. Can be used for 1D and/or 2D arrays
End Type

Private Type ItemDataStruct
    Key As Long     ' pointer to the key if used
    Tag As String   ' tag if used
End Type

Private m_Items() As ItemDataStruct     ' collection if image data
Private m_Keys() As String              ' collection of keys used
Private m_IndexedKey() As Long          ' cross-reference from Key to Item
Private m_CountItems As Long            ' number of items
Private m_CountKey As Long              ' number of keys
Private m_DelayLoad As Boolean          ' option to not create DIB until needed
Private m_Width As Long                 ' width of images within this DIB
Private m_Height As Long                ' height of images within this DIB

Private cImages As c32bppDIB            ' the imagelist

' AddImage... functions are wrappers around the c32bppDIB class
' ----------------------------------------------------------
Public Function AddImage_fromFile(ByVal FileName As String, _
                                Optional ByVal Key As String = vbNullString, _
                                Optional ByVal Tag As String = vbNullString, _
                                Optional ByVal iconBitDepth As Long = 32, _
                                Optional ByVal Scaled As Boolean = True) As Boolean
    ' Parameters:
    '   FileName :: full path & file (unicode supported) containing an image
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   iconBitDepth :: when file is an icon, desired resolution
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_File(FileName, m_Width, m_Height, False, iconBitDepth) Then
        AddImage_fromFile = AddItem(newImg, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromStream(inStream() As Byte, _
                                    Optional ByVal Key As String = vbNullString, _
                                    Optional ByVal Tag As String = vbNullString, _
                                    Optional ByVal streamStart As Long = 0&, _
                                    Optional ByVal streamLength As Long = 0&, _
                                    Optional ByVal iconBitDepth As Long = 32, _
                                    Optional ByVal Scaled As Boolean = True) As Boolean
    ' Parameters:
    '   inStream :: a single dimensioned array containing a complete image
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   streamStart :: byte position in array where image begins
    '   streamLength :: number of bytes that comprise the image
    '   iconBitDepth :: when file is an icon, desired resolution
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_Stream(inStream, m_Width, m_Height, streamStart, streamLength, False, iconBitDepth) Then
        AddImage_fromStream = AddItem(newImg, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromResource(ByVal ResIndex As Variant, ByVal ResSection As Variant, _
                            VBglobal As IUnknown, _
                            Optional ByVal Key As String = vbNullString, _
                            Optional ByVal Tag As String = vbNullString, _
                            Optional ByVal streamStart As Long = 0&, _
                            Optional ByVal streamLength As Long = 0&, _
                            Optional ByVal iconBitDepth As Long = 32, _
                            Optional ByVal Scaled As Boolean = True) As Boolean
    ' Parameters:
    '   ResIndex :: the resource index containing the image
    '   ResSection :: the resource section (i.e., "Custom", vbResIcon)
    '   VBglobal :: always pass as VB.Global
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   streamStart :: byte position in resource array where image begins (usually LBound)
    '   streamLength :: number of bytes that comprise the image (usually UBound+1)
    '   iconBitDepth :: when file is an icon, desired resolution
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    If TypeOf VBglobal Is VB.Global Then
        Dim newImg As c32bppDIB
        Set newImg = New c32bppDIB
        If newImg.LoadPicture_Resource(ResIndex, ResSection, VBglobal, m_Width, m_Height, streamStart, streamLength, iconBitDepth) Then
            AddImage_fromResource = AddItem(newImg, Key, Tag, Scaled, 1&)
        End If
    End If
End Function
Public Function AddImage_fromStdPicture(Picture As StdPicture, _
                                Optional ByVal Key As String = vbNullString, _
                                Optional ByVal Tag As String = vbNullString, _
                                Optional ByVal Scaled As Boolean = True) As Boolean
    ' Parameters:
    '   Picture :: a VB standard picture object
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_StdPicture(Picture) Then
        AddImage_fromStdPicture = AddItem(newImg, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromHandle(Handle As Long, _
                                Optional ByVal Key As String = vbNullString, _
                                Optional ByVal Tag As String = vbNullString, _
                                Optional ByVal Scaled As Boolean = True) As Boolean
    ' Parameters:
    '   Handle :: a memory handle to the image
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_ByHandle(Handle) Then
        AddImage_fromHandle = AddItem(newImg, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromClipBoard(Optional ByVal Key As String = vbNullString, _
                                    Optional ByVal Tag As String = vbNullString, _
                                    Optional ByVal Scaled As Boolean = True) As Boolean
    ' Parameters:
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_ClipBoard() Then
        AddImage_fromClipBoard = AddItem(newImg, Key, Tag, Scaled, 1&)
    End If
End Function

' Render function is a wrapper around the c32bppDIB class
' --------------------------------------------------------
Public Function Render(Index As Variant, ByVal destinationDC As Long, _
                Optional ByVal destX As Long, Optional ByVal destY As Long, _
                Optional ByVal destWidth As Long, Optional ByVal destHeight As Long, _
                Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, _
                Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                Optional ByVal Opacity As Long = 100&, _
                Optional ByVal grayScale As eGrayScaleFormulas = eGrayScaleFormulas.gsclNone, _
                Optional ByVal LightAdjustment As Single = 0!, _
                Optional ByVal Angle As Single = 0!, _
                Optional ByVal CenterOnDestXY As Boolean = False) As Boolean
                
    ' Parameters:
    '   Index :: Ordinal or Key for the image to be rendered
    '   destinationDC :: target DC to render image to
    '   destX,destY :: target DC coordinates to be rendered at
    '   destWidth,destHeight :: dimensions to be rendered to
    '   SrcX,SrcY :: image coordinates to being rendering from
    '   srcWdith,srcHeight :: image dimensions that will be rendered from SrcX,Srcy to destX,destY
    '   Opacity :: value btwn 0 & 100. 100 is fully opaque, 0 is fully transparent
    '   grayScale :: one of several gray scale formulas
    '   LigthAdjustment :: values btwn -100 to 100 which is a percentage of lightness(+) or darkness(-) to add to image
    '   Angle :: values btwn -360 and 360 to rotate image. Rotation is clockwise use negative values for counterclockwise rotation
    '   CenterOnDestXY :: if True, image is rendered around destX,destY else image is rendered from destX,destY

    If m_CountItems = 0& Then Exit Function
    If cImages.ImageType = imgError Then
        ' when DelayLoad is active, no cImage class will be created yet
        If AdjustImageListSize(0&, 0&) = False Then Exit Function
    End If
    
    Dim iIndex As Long, mirrOffset As Long
    iIndex = KeyToIndex(Index)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.Render", "Invalid Key or Index"
    Else
        ' validate optional parameters for source image, translating Index to slot in image list
        ' Regarding mirroring. If the source or destination width or height is negative, mirroring will be applied
        If destWidth = 0& Then destWidth = m_Width
        If destHeight = 0& Then destHeight = m_Height
        If srcWidth = 0& Then
            srcWidth = m_Width
        ElseIf srcWidth < 0& Then                       ' locally, only the destWidth can be negative for mirroring, not srcHeight
            destWidth = -destWidth
            srcWidth = -srcWidth
        End If
        If srcHeight = 0& Then
            srcHeight = m_Height
        ElseIf srcHeight < 0& Then                      ' locally, only the destHeight can be negative for mirroring, not srcHeight
            destHeight = -destHeight
            srcHeight = -srcHeight
        End If
        If destWidth < 0& Then
            ' Mirroring: when using GDI+, we have to tweak the offset.
            ' All images are one bitmap. When GDI+ mirrors the entire bitmap then the
            ' index will be pointing to the wrong subimage. Fix it here.
            If cImages.isGDIplusEnabled Then iIndex = UBound(m_Items) - iIndex + 1
            ' Note: when GDI+ not available, the manual routines do this naturally
        End If
        If SrcX < 0& Then SrcX = 0&
        If SrcY < 0& Then SrcY = 0&
        SrcX = SrcX + m_Width * (iIndex - 1&)
        
        If SrcX + srcWidth > m_Width * iIndex Then srcWidth = m_Width * iIndex - SrcX
        If SrcY + srcHeight > m_Height Then srcHeight = m_Height - SrcY
        
        Render = cImages.Render(destinationDC, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, Opacity, , , , grayScale, LightAdjustment, Angle, CenterOnDestXY)
    End If
    
End Function

Public Function RemoveItem(Index As Variant) As Boolean

    ' Function removes an image form the image list
    ' Parameters:
    '   Index :: The Key (String) or Index (Long) of the item to be removed

    Dim iIndex As Long
    Dim kIndex As Long
    
    iIndex = KeyToIndex(Index)  ' Validate passed Index/Key
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.RemoveItem", "Invalid Key or Index"
        Exit Function
    End If
    
    
    If m_CountItems = 1& Then   ' deleting last one in the list
        Erase m_Keys            ' simply purge everything
        Erase m_Items
        Erase m_IndexedKey
        m_CountItems = 0&
        m_CountKey = 0&
    
    Else
        ' call function to remove the image & slide rest of images down
        If AdjustImageListSize(m_CountItems, iIndex) = False Then
            Err.Raise 1967, "ImageList.RemoveItem", "Failed to remove image"   ' can't think of any reason why this could happen
            Exit Function
        End If
        
        ' does deleted item have a key?
        
        If RemoveKey(iIndex) Then
            ' re-indexing, unless key deleted was to the last m_Items
            If iIndex < m_CountItems Then
                For kIndex = 1 To m_CountKey
                    If m_IndexedKey(kIndex) > iIndex Then m_IndexedKey(kIndex) = m_IndexedKey(kIndex) - 1
                Next
            End If
        End If
    
        m_Items(iIndex).Tag = vbNullString ' free memory as needed
        If iIndex < m_CountItems Then
            CopyMemory ByVal VarPtr(m_Items(iIndex)), ByVal VarPtr(m_Items(iIndex + 1)), (m_CountItems - iIndex) * 8&
            CopyMemory ByVal VarPtr(m_Items(m_CountItems).Tag), 0&, 4&  ' reset the tag pointer
        End If
        m_Items(m_CountItems).Key = 0&
        m_CountItems = m_CountItems - 1
    
        ' array size management; don't let it remain too big after many items were deleted
        If UBound(m_Items) - m_CountItems > 5 Then AdjustImageListSize m_CountItems + 1, 0&
    
    End If
    RemoveItem = True

End Function

Public Function MoveItem(Index As Variant, newIndex As Variant) As Boolean

    ' Function moves an image to another location within the imagelist
    ' Parameters:
    '   Index :: Key or Ordinal of the image to be moved
    '   NewIndex :: Key or Ordinal of the image in the position that Index will be moved to

    Dim fromIndex As Long, toIndex As Long
    Dim bOverlay() As Byte, oSA As SafeArray
    Dim locMoveFrom As Long, locMoveTo As Long
    Dim scanWidth As Long
    Dim kIndex As Long, sKey As String
    Dim bSwap() As Byte
    
    Dim X As Long, y As Long, nrBytesToMove As Long
    
    ' validate Index & NewIndex are within current imagelist bounds
    fromIndex = KeyToIndex(Index)
    If fromIndex = 0 Then
        Err.Raise 1961, "ImageList.MoveItem", "Invalid Key or Index"
        Exit Function
    End If
    toIndex = KeyToIndex(newIndex)
    If toIndex = 0 Then
        Err.Raise 1961, "ImageList.MoveItem", "Invalid Key or Index for New Index"
        Exit Function
    End If

    ' now determine if move needs to be made
    If toIndex = fromIndex Then  ' same position, no move
        MoveItem = True
    Else
        ' ok, now lets get ready to move & shift some stuff around
        scanWidth = m_Width * 4&    ' calc nr bytes per image
        kIndex = fromIndex * scanWidth - scanWidth      ' get byte location of the image that will be moved
        
        If toIndex < fromIndex Then                     ' moving fromIndex to the Left
            locMoveTo = toIndex * scanWidth           ' calc byte location for placement of shifted bytes
            locMoveFrom = locMoveTo - scanWidth         ' calc byte location for start of shifted bytes
            For y = 1 To m_CountKey
                If m_IndexedKey(y) < fromIndex Then      ' reindex keys as needed
                    If m_IndexedKey(y) > toIndex Then m_IndexedKey(y) = m_IndexedKey(y) + 1
                End If
            Next
        Else                                            ' moving fromIndex to the Right
            locMoveTo = kIndex                          ' calc byte location for placement of shifted bytes
            locMoveFrom = kIndex + scanWidth            ' calc byte location for start of shifted bytes
            For y = 1 To m_CountKey
                If m_IndexedKey(y) > fromIndex Then     ' reindex keys as needed
                    If m_IndexedKey(y) < toIndex Then m_IndexedKey(y) = m_IndexedKey(y) - 1
                End If
            Next
        End If
        
        iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 2, cImages.Height, cImages.scanWidth, cImages.BitsPointer
        ReDim bSwap(0 To scanWidth - 1, 0 To m_Height - 1)  ' build array to hold fromIndex image
        For y = 0 To m_Height - 1                           ' and copy the image into the array
            CopyMemory bSwap(0, y), bOverlay(kIndex, y), scanWidth
        Next
        kIndex = toIndex * scanWidth - scanWidth            ' calc position where fromIndex will be relocated
        nrBytesToMove = Abs((toIndex - fromIndex)) * scanWidth   ' and number of bytes to shift the DIB
        
        ' shift DIB, then relocate fromIndex to its new location
        For y = 0 To m_Height - 1
            CopyMemory bOverlay(locMoveTo, y), bOverlay(locMoveFrom, y), nrBytesToMove
            CopyMemory bOverlay(kIndex, y), bSwap(0, y), scanWidth
        Next
        iparseOverlayHost_Byte bOverlay(), 0, 0, 0, 0, 0    ' remove overlay
        
        If Not m_Items(fromIndex).Key = 0 Then              ' change index of moved item's key if needed
            CopyMemory ByVal VarPtr(sKey), m_Items(fromIndex).Key, 4&
            kIndex = FindKey(sKey, False)
            m_IndexedKey(kIndex) = toIndex                  ' set new index
        End If
        If m_Items(toIndex).Key = 0 Then                    ' get index of key at fromIndex' new location
            kIndex = 0
        Else
            CopyMemory ByVal VarPtr(sKey), m_Items(toIndex).Key, 4&
            kIndex = FindKey(sKey, False)                   ' cachce the index for now
        End If
        CopyMemory ByVal VarPtr(sKey), 0&, 4&               ' remove soft pointer of local sKey variable
        
        ' will be shifting array that contains pointers; gotta account for them
        nrBytesToMove = Abs((toIndex - fromIndex)) * 8&             ' calc byte length of structures to move
        
        ReDim bSwap(0 To 7)                                         ' resize to fit ItemDataStructure (8 bytes)
        CopyMemory bSwap(0), ByVal VarPtr(m_Items(fromIndex)), 8&   ' cache fromIndex structure with pointers
        If toIndex < fromIndex Then                                 ' shifting arrays right
            If kIndex Then m_IndexedKey(kIndex) = m_IndexedKey(kIndex) + 1    ' set Index of shifted item & shift the array
            CopyMemory ByVal VarPtr(m_Items(toIndex + 1)), ByVal VarPtr(m_Items(toIndex)), nrBytesToMove
        Else ' fromIndex > toIndex                                  ' shifting array left
            If kIndex Then m_IndexedKey(kIndex) = m_IndexedKey(kIndex) - 1    ' set index of shifted item & shift the array
            CopyMemory ByVal VarPtr(m_Items(fromIndex)), ByVal VarPtr(m_Items(fromIndex + 1)), nrBytesToMove
        End If
        CopyMemory ByVal VarPtr(m_Items(toIndex)), bSwap(0), 8&     ' move fromIndex structure including pointers
        
        MoveItem = True ' done
        
    End If

End Function

Public Function ExtractIcon(Index As Variant, Optional Rtn_BitDepth As Long) As Long

    ' Function extracts an image and converts it to an icon
    ' Return value of the function is the icon handle if it was created.
    ' Icons must be destroyed by you with a call to DestroyIcon API
    
    ' Parameters:
    '   Index :: Ordinal or Key for the image to be iconized
    
    ' Notes:
    
    ' 1. Depending on the per-pixel translucency and number of colors
    ' within the image, the returned icon can exist in one of several
    ' bit depths: 1,4,8 bit paletted, 24 bit true color, or 32 bit alpha-blended
    
    ' 2. When high rotation and/or quality interpolation is used and an image is stretched
    ' or otherwise resized to fit the image list, that image will probably be
    ' alphablended along its edges even if the image previously was 100% opaque.
    ' Therefore, resized images very well may return a 32bpp icon vs 24bpp or lower

    If m_CountItems = 0& Then Exit Function
    If cImages.ImageType = imgError Then
        ' when DelayLoad is active, no cImage class will be created yet
        If AdjustImageListSize(0&, 0&) = False Then Exit Function
    End If
    Dim iIndex As Long
    iIndex = KeyToIndex(Index)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.ExtractIcon", "Invalid Key or Index"
    Else
        ExtractIcon = ExtractIconLocal(cImages, iIndex, Rtn_BitDepth, m_Width, m_Height)
    End If
End Function

Public Function ExtractIconEx(Index As Variant, _
                Optional ByVal MirrorAxisX As Boolean = False, _
                Optional ByVal MirrorAxisY As Boolean = False, _
                Optional ByVal Opacity As Long = 100&, _
                Optional ByVal grayScale As eGrayScaleFormulas = eGrayScaleFormulas.gsclNone, _
                Optional ByVal LightAdjustment As Single = 0!, _
                Optional ByVal Angle As Single = 0!, _
                Optional Rtn_iconWidth As Long, Optional Rtn_iconHeight As Long, Optional Rtn_BitDepth As Long) As Long
                
    
    ' Function will return an icon in a rendered appearance other than the original image
    ' See notes in ExtractIcon also
                
    ' Parameters:
    '   Index :: Ordinal or Key for the image to be rendered
    '   MirroredAxisX :: Set to true if image is to be mirrored horizontally
    '   MirroredAxisY :: Set to true if image is to be mirrored vertically
    '   Opacity :: value btwn 0 & 100. 100 is fully opaque, 0 is fully transparent
    '   grayScale :: one of several gray scale formulas
    '   LigthAdjustment :: values btwn -100 to 100 which is a percentage of lightness(+) or darkness(-) to add to image
    '   Angle :: values btwn -360 and 360 to rotate image. Rotation is clockwise use negative values for counterclockwise rotation
    '   ... Rotation will create a larger icon than the source with the source image rotated around the center of the icon
    '   Rtn_iconWidth [out]:: will be the width of the created icon.
    '   Rtn_iconHeight [out]:: will be the height of the created icon
    '   Rtn_BitDepth [out]:: will be the bit depth of the icon: 1,4,8,24,32

    If m_CountItems = 0& Then Exit Function
    If cImages.ImageType = imgError Then
        ' when DelayLoad is active, no cImage class will be created yet
        If AdjustImageListSize(0&, 0&) = False Then Exit Function
    End If
    
    Dim iIndex As Long
    iIndex = KeyToIndex(Index)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.ExtractIconEx", "Invalid Key or Index"
    Else
        ' validate optional parameters for source image, translating Index to slot in image list
        ' Regarding mirroring. If the source or destination width or height is negative, mirroring will be applied
        Dim cX As Long, cY As Long, X As Long
        Dim destX As Long, destY As Long
        Dim tImage As c32bppDIB
        If MirrorAxisY Then cY = -m_Height Else cY = m_Height
        If MirrorAxisX Then
            cX = -m_Width
            ' Mirroring: when using GDI+, we have to tweak the offset.
            ' All images are one bitmap. When GDI+ mirrors the entire bitmap then the
            ' index will be pointing to the wrong subimage. Fix it here.
            If cImages.isGDIplusEnabled Then iIndex = UBound(m_Items) - iIndex + 1
            ' Note: when GDI+ not available, the manual routines do this naturally
        Else
            cX = m_Width
        End If
        X = m_Width * (iIndex - 1&)

        Set tImage = New c32bppDIB
        
        ' reduce icon size if possible. Can be done when rotation at 90 degree increments
        If Angle = Int(Angle) Then                      ' else degree fraction in use
            If Angle Mod 360 = 0 Then                   ' no rotation?
                tImage.InitializeDIB m_Width, m_Height
                Angle = 0!
            ElseIf Angle Mod 180 = 0 Then                   ' vertical rotation only?
                tImage.InitializeDIB m_Width, m_Height
                destY = m_Height \ 2: destX = m_Width \ 2
            ElseIf Angle Mod 90 = 0 Then                    ' 90 degree rotation?
                tImage.InitializeDIB m_Height, m_Width
                destY = m_Width \ 2: destX = m_Height \ 2   ' calc center point
            End If
        End If
        If tImage.ImageType = imgError Then             ' angle did not fit one of above categories
            destX = Sqr(m_Width * m_Width + m_Height * m_Height)    ' create icon to allow any angle rotation
            tImage.InitializeDIB destX, destX                       ' create the temp image
            destY = destX \ 2: destX = destY                        ' calc center point
        End If
        tImage.gdiToken = cImages.gdiToken
            
        ' render image to a temporary dib
        cImages.Render 0&, destX, destY, m_Width, m_Height, X, 0, cX, m_Height, Opacity, , , tImage, grayScale, LightAdjustment, Angle, (Angle <> 0!)
        Rtn_iconWidth = tImage.Width
        Rtn_iconHeight = tImage.Height
        ' call local function to return icon from that dib
        ExtractIconEx = ExtractIconLocal(tImage, 1, Rtn_BitDepth, tImage.Width, tImage.Height)
        Set tImage = Nothing
    End If
End Function

Public Function IsKeyAssigned(ByVal Key As String) As Long
    ' function will return the Index of the image the key is associated with
    ' If the function returns < 0, then no image is associated with the key
    Dim bIsUnique As Boolean
    Dim Index As Long
    Index = FindKey(Key, bIsUnique)
    If Not bIsUnique Then IsKeyAssigned = m_IndexedKey(Index)
End Function

' Property Returns the number of images currently in this list
Public Property Get ImageCount() As Long
    ImageCount = m_CountItems
End Property
Public Property Get MaxImages() As Long
    ' most images that can exist on a single imagelist
    ' Idea is to keep DIBs less than 4096kb to prevent CreateDIBSection from failing
    MaxImages = GetMaxImageCount(m_Width, m_Height)
End Property

Public Property Get Width() As Long ' individual width of each image on imagelist
    Width = m_Width
End Property
Public Property Get Height() As Long ' individual height of each image on imagelist
    Height = m_Height
End Property

' Property Returns/Sets the Key associated with an image in the imagelist
' Parameters:
'   Index :: The Key (String) or Index (Long) of the item to be removed
'   NewKey (Property Let) :: The key to be assigned to the image
Public Property Get Key(Index As Variant) As String
    Dim iIndex As Long
    iIndex = KeyToIndex(Index)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.ImageKey", "Invalid Key or Index"
    Else
        Dim itemKey As String
        If m_Items(iIndex).Key Then
            CopyMemory ByVal VarPtr(itemKey), m_Items(iIndex).Key, 4&
            Key = itemKey
            CopyMemory ByVal VarPtr(itemKey), 0&, 4&
        End If
    End If
End Property
Public Property Let Key(Index As Variant, ByVal newKey As String)

    Dim kIndex As Long, bIsNew As Boolean, OldKey As String
    Dim newIndex As Long, iIndex As Long
    
    iIndex = KeyToIndex(Index)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.ImageKey", "Invalid Key or Index"
        Exit Property
    End If
    
    If newKey = vbNullString Then
        RemoveKey iIndex
    Else
        ' if item already has a key, get its key index
        If m_Items(iIndex).Key Then
            CopyMemory ByVal VarPtr(OldKey), m_Items(iIndex).Key, 4&
            kIndex = FindKey(OldKey, False)
            CopyMemory ByVal VarPtr(OldKey), 0&, 4&
        End If
        
        newIndex = FindKey(newKey, bIsNew)
        ' validate new key not duplicated
        If bIsNew = False Then      ' key exists
            ' error if not for the same key index
            If newIndex = kIndex Then
                m_Keys(kIndex) = newKey ' change it per user request
                m_Items(iIndex).Key = StrPtr(m_Keys(kIndex))
            Else
                Err.Raise 1962, "ImageList.ImageKey", "Key is already associated with another image in the collection"
            End If
            
        ElseIf (newIndex = kIndex) Or (newIndex - kIndex = 1) And kIndex > 0& Then
            ' new key is unique & will be placed in same position as oldkey
            m_Keys(kIndex) = newKey
            m_Items(iIndex).Key = StrPtr(m_Keys(kIndex))
            
        Else
            ' new key is unique and will be placed elsewhere
            If kIndex Then
                RemoveKey iIndex    ' remove old key if any
                AddKey newKey, 0&, iIndex          ' pass 0 to re-calc position; array changed with key removal
            Else
                AddKey newKey, newIndex, iIndex    ' add new key
            End If
        End If
        
    End If
End Property

' Property Returns/Sets the Tag associated with an image in the imagelist
' Parameters:
'   Index :: The Key (String) or Index (Long) of the item to be removed
'   NewTag (Property Let) :: The tag to be assigned to the image
Public Property Get Tag(Index As Variant) As String
    Dim iIndex As Long
    iIndex = KeyToIndex(Index)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.ImageTag", "Invalid Key or Index"
    Else
        Tag = m_Items(iIndex).Tag
    End If
End Property
Public Property Let Tag(Index As Variant, NewTag As String)
    Dim iIndex As Long
    iIndex = KeyToIndex(Index)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.ImageTag", "Invalid Key or Index"
    Else
        m_Items(iIndex).Tag = NewTag
    End If
End Property


' *********************************************************
'       INTERNAL AND PROPERTY-PAGE RELATED ROUTINES
' *********************************************************

Public Property Get DelayLoad() As Boolean ' design time only via the property page
    DelayLoad = m_DelayLoad
End Property
Public Property Let DelayLoad(bDelay As Boolean)
    m_DelayLoad = bDelay
End Property
Friend Property Get ImageClass() As c32bppDIB
    Set ImageClass = cImages
End Property

Private Function AddItem(cHost As c32bppDIB, Key As String, Tag As String, Scaled As Boolean, ppgBuffer As Long) As Boolean
    
    ' Function Adds a new image to the imagelist
    ' Parameters:
    '   Key :: Key to reference this image
    '   Tag :: Any string data you want to associate with the image
    
    Dim kIndex As Long, incrSize As Long
    Dim bIsNew As Boolean
    Dim imgCx As Long, imgCy As Long
    Dim imgX As Long, imgY As Long
    Dim tDC As Long
    
    ' first see if a passed key is already registered
    If Len(Key) Then
        kIndex = FindKey(Key, bIsNew)
        If bIsNew = False Then
            Err.Raise 1962, "ImageList.AddImage", "Key is already associated with another image in the collection"
            Exit Function    ' duplicated key
        End If
    Else
        Key = vbNullString
    End If
        
    ' when working with the GUI property page, we buffer extra images expecting several additions
    incrSize = ppgBuffer
    imgCx = Me.MaxImages
    If imgCx < m_CountItems + incrSize Then
        If imgCx < m_CountItems + 1 Then
            If m_CountItems = 0& Then
                Err.Raise 1968, "ImageList.AddImage", "Image list cannot support images of that size"
            Else
                Err.Raise 1969, "ImageList.AddImage", "Image list cannot append additional images. Create new image list."
            End If
            Exit Function
        End If
        incrSize = 1
    End If
    
    ' resize arrays if needed
'    If m_CountItems Then
        If m_CountItems = UBound(m_Items) Then
            incrSize = m_CountItems + incrSize
        Else
            incrSize = 0&
        End If
'    End If
    If AdjustImageListSize(incrSize, 0&) = False Then  ' resize & simultaneously handle for delay load
        Err.Raise 1970, "ImageList.AddImage", "System resources preventing changing the size of the image list."
        Exit Function
    End If
    m_CountItems = m_CountItems + 1
    
    ' scale keeping aspect ratio if requested
    If Scaled Then
        cHost.ScaleImage m_Width, m_Height, imgCx, imgCy, eScaleOptions.ScaleToSize
        imgX = (m_Width - imgCx) \ 2
        imgY = (m_Height - imgCy) \ 2
    Else    ' stretching
        imgCx = m_Width
        imgCy = m_Height
    End If
        
    ' render the image to our DIB
    'tDC = cImages.LoadDIBinDC(True)
    If cHost.Render(tDC, (m_CountItems - 1) * m_Width + imgX, imgY, imgCx, imgCy, , , , , , , , cImages) Then
        ' add tag property as needed
        'Call cImages.LoadDIBinDC(False)
        If Not Tag = vbNullString Then
            CopyMemory ByVal VarPtr(m_Items(m_CountItems).Tag), StrPtr(Tag), 4&
            CopyMemory ByVal VarPtr(Tag), 0&, 4&
        End If
        
        ' add the key to the collection
        If kIndex Then AddKey Key, kIndex, m_CountItems
        
        AddItem = True
        
    Else
        'Call cImages.LoadDIBinDC(False)
        m_CountItems = m_CountItems - 1
        Err.Raise 1966, "ImageList.AddImage", "Failed to render image to the image list"
    End If
    

End Function

Friend Function InitializeImageList(nrImages As Long, Width As Long, Height As Long) As Boolean
    ' helper function to ensure imagelist can be created. Can be called from property page also
    Dim bInIDE As Boolean
    Set cImages = New c32bppDIB
    Debug.Assert IsUnCompiled(bInIDE)
    cImages.KeepGDIplusActive = Not bInIDE
    
    If m_CountItems = 0& Then '
        If nrImages <= GetMaxImageCount(Width, Height) Then
            m_Width = Width
            m_Height = Height
            If nrImages = 0& Then
                InitializeImageList = True
            Else
                m_CountItems = 0&
                Erase m_Keys()
                Erase m_IndexedKey()
                ReDim m_Items(1 To nrImages)
                InitializeImageList = True
            End If
        End If
    End If
End Function

Private Function RemoveKey(ItemIndex As Long) As Boolean
    
    ' Helper function to remove a specific key
    
    If m_Items(ItemIndex).Key = 0& Then Exit Function
    
    Dim Key As String, kIndex As Long
    
    If m_CountKey = 1& Then     ' easy, last key -- simply erase the array
        Erase m_Keys
        Erase m_Keys
        m_Items(m_CountKey).Key = 0&
        m_CountKey = 0&
    Else
        ' copy the key from the key array to our Key variable
        CopyMemory ByVal VarPtr(Key), m_Items(ItemIndex).Key, 4&
        ' get the position of the key within the array
        kIndex = FindKey(Key, False)
        ' shift the key array down one as needed
        If kIndex < m_CountKey Then
            CopyMemory ByVal VarPtr(m_Keys(kIndex)), ByVal VarPtr(m_Keys(kIndex + 1)), (m_CountKey - kIndex) * 4&
            CopyMemory m_IndexedKey(kIndex), m_IndexedKey(kIndex + 1), (m_CountKey - kIndex) * 4&
        End If
        m_Items(ItemIndex).Key = 0& ' dereference
        ' zero out the final key, our Key variable has the value now
        CopyMemory ByVal VarPtr(m_Keys(m_CountKey)), 0&, 4&
        m_CountKey = m_CountKey - 1
        
        ' array size management; don't let it remain too big when many items are deleted
        If UBound(m_Keys) - m_CountKey > 5 Then
            ReDim Preserve m_Keys(1 To m_CountKey + 1)
            ReDim Preserve m_IndexedKey(1 To m_CountKey + 1)
        End If
        
        RemoveKey = True
        
    End If

End Function

Private Sub AddKey(Key As String, KeyIndex As Long, ItemIndex As Long)

    ' Helper function to add a eky

    If KeyIndex = 0& Then KeyIndex = FindKey(Key, False)
    
    ' increment & resize array if needed, maintaining small buffer
    m_CountKey = m_CountKey + 1
    If m_CountKey = 1& Then
        ReDim m_Keys(1 To m_CountKey + 3)
        ReDim m_IndexedKey(1 To m_CountKey + 3)
    ElseIf m_CountKey > UBound(m_Keys) Then
        ReDim Preserve m_Keys(1 To m_CountKey + 3)
        ReDim Preserve m_IndexedKey(1 To m_CountKey + 3)
    End If
    If KeyIndex < m_CountKey Then ' need to shift array to maintain ASC sort order
        CopyMemory ByVal VarPtr(m_Keys(KeyIndex + 1)), ByVal VarPtr(m_Keys(KeyIndex)), (m_CountKey - KeyIndex) * 4&
        CopyMemory m_IndexedKey(KeyIndex + 1), m_IndexedKey(KeyIndex), (m_CountKey - KeyIndex) * 4&
    End If
    m_IndexedKey(KeyIndex) = ItemIndex      ' set reference to image index
    m_Items(ItemIndex).Key = StrPtr(Key)    ' cross-reference item to key
    CopyMemory ByVal VarPtr(m_Keys(KeyIndex)), StrPtr(Key), 4& ' move Key pointer vs Copying
    CopyMemory ByVal VarPtr(Key), 0&, 4&    ' remove Key strPtr

End Sub

Private Function KeyToIndex(Key As Variant) As Long
    
    ' Helper function to validate passed Key/Index
    
    Dim kIndex As Long, bIsNew As Boolean
    
    Select Case VarType(Key)
    Case vbString                           ' Strings are always treated as Keys
        If Not Key = vbNullString Then
            kIndex = FindKey(CStr(Key), bIsNew)
            If bIsNew = False Then KeyToIndex = m_IndexedKey(kIndex)
        End If
    Case vbLong, vbInteger, vbDouble, vbByte, vbSingle ' Numerical varTypes are treated as Indexes
        On Error Resume Next
        kIndex = Int(Key)
        If kIndex > 0& And kIndex <= m_CountItems Then KeyToIndex = kIndex
        If Err Then Err.Clear
    End Select
    
End Function

Private Function FindKey(Key As String, ByRef isNew As Boolean) As Long

    ' BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between just 1 and 20 iterations
    
    ' [in] Key. A value to search for. Order is always ascending
    ' [out] isNew. If Key not found, isNew is True else False
    ' [out] Return value: The Index where Key was found or where the new Key should be inserted

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If m_CountKey = 0& Then
        newIndex = 1&
        isNew = True
    Else
    
        UB = m_CountKey
        LB = 1&
        
        Do Until LB > UB
            newIndex = LB + ((UB - LB) \ 2&)
            Select Case StrComp(Key, m_Keys(newIndex), vbTextCompare)
            Case 0 ' key found
                Exit Do
            Case Is < 0 ' key is < current key
                UB = newIndex - 1&
            Case Else ' key is > current key
                LB = newIndex + 1&
            End Select
        Loop
    
        If LB > UB Then  ' key was not found
            If StrComp(Key, m_Keys(newIndex), vbTextCompare) > 0 Then newIndex = newIndex + 1
            isNew = True
        Else
            isNew = False
        End If
    
    End If
    
    FindKey = newIndex

End Function

Private Function AdjustImageListSize(newCount As Long, ByVal DeleteIndex As Long) As Boolean
        
    ' Helper function. Manages size of imagelist in response to additions/deletions
    ' Parameters:
    '   newCount :: can be zero first time AddItem is called and delay load is in effect
    '               else resize imagelist to support this number of images
    '   DeleteIndex :: Index of the item to be deleted, called from RemoveItem only
        
    Dim bResize As Boolean, bDummy() As Byte
    Dim bInIDE As Boolean
    If cImages Is Nothing Then
        Set cImages = New c32bppDIB
        Debug.Assert IsUnCompiled(bInIDE)
        cImages.KeepGDIplusActive = Not bInIDE
    End If
    
    ' check for delay load. cImages will be created, but not initialized
    ' When ImageType=imgError, this should be the case
    If cImages.ImageType = imgError Then
        ' see if image waiting to be processed
        bResize = cImages.LoadPicture_FromOrignalFormat()
        If bResize = False Then
            ' see if our m_items array is empty, shoudln't be if DelayLoad=True
            If iparseIsArrayEmpty(VarPtrArray(m_Items)) Then
                ' create the blank image
                bResize = cImages.InitializeDIB(m_Width * UBound(m_Items), m_Height)
            End If
        Else
            cImages.SetOriginalFormat bDummy    ' don't need cached bytes any long
        End If
        If bResize = False Then 'failure, abort
            m_CountItems = 0&
            m_CountKey = 0&
            Erase m_IndexedKey()
            Erase m_Keys()
            Erase m_Items()
            Exit Function
        End If
        If newCount = 0& Then bResize = False
    ElseIf newCount Then    ' can be zero only when triggering delay load
        bResize = (UBound(m_Items) <> newCount)
    End If
    
    If bResize = True Or DeleteIndex > 0& Then
    
        Dim bSrc() As Long, bDst() As Long
        Dim SAsrc As SafeArray, SAdst As SafeArray
        Dim scanWidth As Long, lRow As Long, lCol As Long
        Dim tDIB As c32bppDIB
        
        If DeleteIndex Then ' deleting item, not resizing imagelist
    
            Dim lastCol As Long, fillSize As Long
            
            ' use array overlays, so we don't need to worry about potential of incrementing a negative memory ponter
            iparseOverlayHost_Long bSrc(), VarPtr(SAsrc), 2, m_Height, UBound(m_Items) * m_Width, cImages.BitsPointer
                        
            scanWidth = (m_CountItems - DeleteIndex) * 4 * m_Width  ' nr of bytes to shift
            lastCol = m_CountItems * m_Width - m_Width              ' position of last image in the list
            fillSize = m_Width * 4&                                 ' nr of bytes to zeroize
            
            If scanWidth Then                                       ' last image being removed?
                DeleteIndex = DeleteIndex * m_Width - m_Width       ' nope, calc its byte offset in the imagelist
                lCol = DeleteIndex + m_Width                        ' calc bytre position of next image in the list
                
                For lRow = 0 To m_Height - 1                        ' loop thru rows
                    CopyMemory bSrc(DeleteIndex, lRow), bSrc(lCol, lRow), scanWidth ' shift images one left
                    FillMemory bSrc(lastCol, lRow), fillSize, 0&                    ' zeroize last position
                Next
            Else                                                    ' removing last image
                For lRow = 0 To m_Height - 1                        ' just zeroize its bytes
                    FillMemory bSrc(lastCol, lRow), fillSize, 0&
                Next
            End If
            
            iparseOverlayHost_Long bSrc, 0, 0, 0, 0, 0              ' remove overlay & return
            AdjustImageListSize = True
        
        Else    ' resizing the imagelist
            
            Set tDIB = New c32bppDIB                                ' create new DIB of correct size; abort if failure
            If tDIB.InitializeDIB(newCount * m_Width, m_Height) Then
        
                ' use array overlays, to transfer bytes from one to the other
                iparseOverlayHost_Long bSrc(), VarPtr(SAsrc), 2, m_Height, UBound(m_Items) * m_Width, cImages.BitsPointer
                iparseOverlayHost_Long bDst(), VarPtr(SAdst), 2, m_Height, newCount * m_Width, tDIB.BitsPointer
                
                ' calc how many bytes need to be transfered per row
                If cImages.Width > tDIB.Width Then  ' reducing the size
                    scanWidth = tDIB.scanWidth
                Else                                ' increasing the size
                    scanWidth = cImages.scanWidth
                End If
                
                For lRow = 0 To m_Height - 1        ' loop thru and transfer the bytes
                    CopyMemory bDst(lCol, lRow), bSrc(lCol, lRow), scanWidth
                Next
                iparseOverlayHost_Long bSrc, 0, 0, 0, 0, 0  ' remove overlays
                iparseOverlayHost_Long bDst, 0, 0, 0, 0, 0
            
                cImages.SwapDIBs tDIB                       ' move copy to our class, preserving class pointer
                ReDim Preserve m_Items(1 To newCount)       ' resize the items array & return
                
                AdjustImageListSize = True
            End If
        
        End If
    Else
        AdjustImageListSize = True
    End If

End Function

' Exports imagelist data to a byte array for storage from usercontrol's WriteProperty
' Also called by usercontrol's ExportImageList function
Friend Function ExportImageList(bExport() As Byte, HeaderBytes As Long, _
                            ValidationOffset As Long, _
                            ImageOffset As Long, _
                            ByVal CompressOption As Boolean) As Boolean

    Dim lSizeArray As Long, lImageSize As Long
    Dim lPointer As Long
    Dim iItem As Long
    Dim bImage() As Byte, bKeys() As Byte
    Dim oSA As SafeArray, bOverlay() As Byte
    Dim Key As String
    Dim keyLen As Long, tagLen As Long
    Dim bCompressMethod As Byte ' 0=PNG, or uncompressed BMP , 1=BMP compressed
    Dim cCompressor As cHuffman
    
    Const xtraBytes As Long = 3
    Const vVersion As Long = 2      ' must match the version in lvImageList control
    
    ' ^^  version 1: delay load only extra property
    ' ^^  version 2: + 1 byte: 0=png format, 1=huffmann compressed bitmap, 2=uncompressed bitmap
    
    ' Export data will look like following:
    ' Nr Bytes  Data
    '   nnn     header bytes used by the usercontrol (image list name, version, other as needed)
    '   4       nrItems :: overall DIB width/nrItems = individual image width.  ImageHeight is overall DIB height
    '   4       nrKeys  :: speeds up importing because we can pre-initialize key arrays
    '   8       width & height of images
    '   for each item, following is defined
    '       4   length of key & tag (zero if not keyed/tagged) HiWord=Tag length, LoWord=Key Length
    '       nnn key in unicode if present
    '       nnn tag in unicode of present
    '   4       extra offset bytes to beginning of image. Validation Offset
    '   nnn     extra bytes if used
    '   ---     image begins. Image Offset
    
    If m_CountItems Then
        
        Call AdjustImageListSize(m_CountItems, 0&)          ' trim the imagelist to remove extra space if needed
        ' save imagelist to PNG/BMP in the bImage() array
        If cImages.SaveToStream_PNG(bImage) = False Then    ' if no GDI+ or zLib installed, then
            If cImages.SaveToStream(bImage) = False Then Exit Function
            ' to compress or not?
            If CompressOption Then   ' yep try it
                Set cCompressor = New cHuffman
                If cCompressor.EncodeListData(Nothing, bImage, bOverlay) Then
                    Erase bImage
                    CopyMemory ByVal VarPtrArray(bImage), ByVal VarPtrArray(bOverlay), 4&
                    CopyMemory ByVal VarPtrArray(bOverlay), vbDefault, 4&
                    bCompressMethod = 2
                End If
            End If
        End If
        lImageSize = UBound(bImage) - LBound(bImage) + 1
        
        ' collecting key/tag info. Loop thru once to count total size of all
        For iItem = 1 To m_CountItems  ' start counting how many bytes are needed for non-image data
            If m_Items(iItem).Key Then
                CopyMemory ByVal VarPtr(Key), m_Items(iItem).Key, 4& ' create soft reference to key -- must be removed by end of routine
                lSizeArray = lSizeArray + Len(Key) * 2  ' unicode size of key
            End If
            ' add space for the tag if it is used
            If m_Items(iItem).Tag <> vbNullString Then
                lSizeArray = lSizeArray + Len(m_Items(iItem).Tag) * 2 ' unicode size of key
            End If
        Next
        
        If lSizeArray > 1024 Then       ' compressing less than 1kb generally not much savings
            bCompressMethod = (bCompressMethod Or 4)        ' add flag to show compressed
            lSizeArray = lSizeArray + m_CountItems * 4      ' calc size of array needed
            ReDim bKeys(1 To lSizeArray)                    ' size the array
            ' overlay array over bKeys. This allows the writing to occur to this array
            iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, lSizeArray, 0, VarPtr(bKeys(1))
            lPointer = 0 ' bOverlay is zero-bound
        End If
        If (bCompressMethod And 4) = 0 Then ' no compression of key/tag data
            ' calc size of array needed for the entire export
            lSizeArray = lSizeArray + HeaderBytes + m_CountItems * 4 + xtraBytes + 20
            ' ^^ (20= 4: item count, 4:key count, 4:width, 4:height, 4:xtraByte entry)
            ReDim bExport(1 To lSizeArray + lImageSize) ' now we can resize our array
            ' overlay array over bExport. Allows writing to occur to this array
            iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, lSizeArray + lImageSize, 0, VarPtr(bExport(1))
            lPointer = HeaderBytes + 16 ' bOverlay is zero-bound
        End If
        
        ' loop thru key/tags again, this time writing them to bOverlay
        For iItem = 1 To m_CountItems
            If m_Items(iItem).Key Then                                      ' keyed?
                CopyMemory ByVal VarPtr(Key), m_Items(iItem).Key, 4&        ' create soft reference to key -- must be removed by end of routine
                keyLen = Len(Key) * 2                                       ' cache length & then store key
                CopyMemory bOverlay(lPointer + 4), ByVal m_Items(iItem).Key, keyLen
            Else ' no key
                keyLen = 0&
            End If
            If m_Items(iItem).Tag <> vbNullString Then                      ' tagged?
                tagLen = Len(m_Items(iItem).Tag) * 2                        ' cache length & then store tag
                CopyMemory bOverlay(lPointer + keyLen + 4), ByVal StrPtr(m_Items(iItem).Tag), tagLen
            Else
                tagLen = 0&
            End If
            CopyMemory bOverlay(lPointer), (keyLen Or (&H10000 * tagLen)), 4& ' store lengths of key/tag
            lPointer = lPointer + tagLen + keyLen + 4                        ' position to next available byte
        Next
        iparseOverlayHost_Byte bOverlay(), 0, 0, 0, 0, 0    ' remove overlay
        CopyMemory ByVal VarPtr(Key), 0&, 4&                ' remove soft reference inside the Key variable
        
        If (bCompressMethod And 4) = 4 Then ' we are going to compress the keys/tags
            If cCompressor Is Nothing Then Set cCompressor = New cHuffman
            cCompressor.EncodeListData Nothing, bKeys(), bOverlay()
            ' calculate entire size of the export array
            lSizeArray = HeaderBytes + UBound(bOverlay) + 1 + xtraBytes + lImageSize + 20
            ReDim bExport(1 To lSizeArray)
            ' copy the compressed keys/tags to Export
            CopyMemory bExport(HeaderBytes + 17), bOverlay(0), UBound(bOverlay) + 1
            ' set pointer for start of image data
            lPointer = HeaderBytes + UBound(bOverlay) + 17
            bCompressMethod = (bCompressMethod Or 4)    ' update compression options used
            Erase bKeys     ' clear unneeded arrays
            Erase bOverlay
        Else
            lPointer = lPointer + 1 ' move from zero-bound to one-bound pointer
        End If
        
        ' place header information in the Export data
        CopyMemory bExport(HeaderBytes + 1), m_CountItems, 4&                   ' add nr items
        If m_CountKey Then CopyMemory bExport(HeaderBytes + 5), m_CountKey, 4&  ' add nr keys
        CopyMemory bExport(HeaderBytes + 9), m_Width, 4&                        ' add width & height of images
        CopyMemory bExport(HeaderBytes + 13), m_Height, 4&
        
        ValidationOffset = lPointer     ' location where xtraBytes data can be found
        CopyMemory bExport(lPointer), xtraBytes, 4&
        ' Version 1:
            CopyMemory bExport(lPointer + 4), m_DelayLoad, 2&
        ' Version 2: add compression to bitmaps if cannot be converted to PNG
            CopyMemory bExport(lPointer + 6), bCompressMethod, 1&
        ' Version 3+: add extra bytes & adjust pointer appropriately
        '   ...
        ImageOffset = lPointer + xtraBytes  ' location where image data resides
        CopyMemory bExport(ImageOffset + 4), bImage(LBound(bImage)), lImageSize ' copy the image over
        ExportImageList = True
        
    End If
    
End Function

' Imports imagelist data from a byte array from usercontrol's ReadProperty
' Also called by usercontrol's ImportImageList function
Friend Function ImportImageList(bImport() As Byte, HeaderBytes As Long, ValidationOffset As Long, iVersion As Long) As Boolean

    Dim lPointer As Long, lBoundOffset As Long
    Dim iItem As Long
    Dim Key As String, strLen As Long, keyLen As Long
    Dim bInIDE As Boolean
    Dim bCompressMethod As Byte ' 0=either PNG or raw Bimtap, 1=compressed bits only, 2=compressed bmp w/header, 4=compressed keys/tags
    Dim xtraBytes As Long
    Dim bOverlay() As Byte, oSA As SafeArray
    Dim bKeys() As Byte
    
    Dim cCompressor As cHuffman
    
    ' Import data is same as described in ExportImageList
    ' Nr Bytes  Data
    '   nnn     header bytes used by the usercontrol (image list name, version, other as needed)
    '   4       nrItems :: overall DIB width/nrItems = individual image width.  ImageHeight is overall DIB height
    '   4       nrKeys  :: speeds up importing because we can pre-initialize key arrays
    '   8       width & height of images
    '   for each item, following is defined
    '       4   length of key & tag (zero if not keyed/tagged) HiWord=Tag length, LoWord=Key Length
    '       nnn key in unicode if present
    '       nnn tag in unicode of present
    '   4       extra offset bytes to beginning of image. 2 for Version 1
    '   nnn     extra bytes if used
    '   ---     image begins
    
    lPointer = LBound(bImport) + HeaderBytes                ' start of array
    lBoundOffset = LBound(bImport) - 1
    m_CountKey = 0&
    CopyMemory m_CountItems, bImport(lPointer), 4&         ' number of sub images?
    If m_CountItems Then
        ReDim m_Items(1 To m_CountItems)
        CopyMemory iItem, bImport(lPointer + 4), 4&         ' number of keyed images?
        If iItem Then
            ReDim m_Keys(1 To iItem)                        ' resize arrays
            ReDim m_IndexedKey(1 To iItem)
        End If
        CopyMemory m_Width, bImport(lPointer + 8), 4&       ' width of images
        CopyMemory m_Height, bImport(lPointer + 12), 4&     ' height of images
        
        ' get version data...
        CopyMemory xtraBytes, bImport(ValidationOffset + lBoundOffset), 4&
        If iVersion > 1 Then    ' we have image/key/tag compression options
            CopyMemory bCompressMethod, bImport(ValidationOffset + 6 + lBoundOffset), 1&
            If (bCompressMethod And 4) = 4 Then ' keys/tags are huffmann compressed
                Set cCompressor = New cHuffman
                ' compressed bytes will always start 16bytes from the header bytes
                cCompressor.DecodeByte bImport, lPointer + 16, bKeys(), Nothing
                ' overlay array on returned array; neat hack in using safearray as "Any" array
                iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, Abs(UBound(bKeys) - LBound(bKeys)) + 1, 0, VarPtr(bKeys(LBound(bKeys)))
                lPointer = 0    ' overlays are always zero-bound
            End If
        End If
        If cCompressor Is Nothing Then  ' keys/tags are not compressed
            ' overlay array over our import data
            iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, Abs(UBound(bImport) - LBound(bImport)) + 1, 0, VarPtr(bImport(LBound(bImport)))
            lPointer = HeaderBytes + 16 ' where keys/tags begin
        End If
        
        ' routine extracts data from 1 of 2 arrays, the overlay array points to either
        For iItem = 1 To m_CountItems
            CopyMemory strLen, bOverlay(lPointer), 4&        ' key, tag?
            keyLen = (strLen And &HFFFF&)                   ' key length (max is 512 or 256 characters)
            If keyLen Then
                Key = Space$(keyLen \ 2)                    ' build string to hold key
                CopyMemory ByVal StrPtr(Key), bOverlay(lPointer + 4), keyLen  ' copy string into the key
                If m_CountKey = 0& Then                     ' handle 1st instance, else AddKey will reset array
                    m_CountKey = 1&                         ' update key count
                    m_Keys(m_CountKey) = Key                ' copy the key
                    m_IndexedKey(m_CountKey) = iItem        ' identify item it is keyed to
                    m_Items(iItem).Key = StrPtr(m_Keys(m_CountKey)) ' cross-reference
                Else
                    AddKey Key, 0&, iItem                   ' call function to sort key & add references
                End If
            End If
            strLen = (strLen \ &H10000)                     ' tag length is in HiWord of value
            If strLen Then                                  ' has a tag property?
                m_Items(iItem).Tag = Space$(strLen \ 2)     ' create string for the tag
                CopyMemory ByVal StrPtr(m_Items(iItem).Tag), bOverlay(lPointer + keyLen + 4), strLen ' copy string into the tag
            End If
            lPointer = lPointer + keyLen + strLen + 4
        Next
        ' done, remove the overlay
        iparseOverlayHost_Byte bOverlay(), 0, 0, 0, 0, 0
        
        If Not cCompressor Is Nothing Then
            ' keys/tags were uncompressed; free array & update current compression options
            bCompressMethod = (bCompressMethod And Not 4)
            Erase bKeys
        End If
        lPointer = ValidationOffset + lBoundOffset ' pointer to xtraBytes data
        ' Version 0/1: 2 bytes
            CopyMemory m_DelayLoad, bImport(lPointer + 4), 2&
        ' Version 2: 1 byte
            ' processed above; contains image/key/tag compression if used
        ' Version 3+: process those bytes
        '  ...
        lPointer = lPointer + xtraBytes + 4     ' move to beginning of image data
        
        ' one subversion of version 2 compresses only pixel data. The other compresses entire bitmap with header
        ' When entire bitmap is compressed, we can allow DelayLoad to be active else we can't
        If bCompressMethod = 2 Then
            If cCompressor Is Nothing Then Set cCompressor = New cHuffman
            If cCompressor.DecodeByte(bImport(), lPointer, bOverlay(), Nothing) = False Then
                Erase m_Items()
                Erase m_Keys()
                m_CountItems = 0&
                m_CountKey = 0&
                Exit Function
            End If
            Set cCompressor = Nothing
        End If
        
        ' create the DIB class and import the image, unless delay load is set
        If m_DelayLoad = True And bCompressMethod <> 1 Then
            ' delay load does not create the 32bpp imagelist until it is first used
            ' Resize the array so it only holds the image data & format header
            Set cImages = New c32bppDIB
            If bCompressMethod = vbDefault Then
                ReDim Preserve bOverlay(1 To UBound(bImport) - lPointer + 1)
                CopyMemory bOverlay(1), bImport(lPointer), UBound(bOverlay)
            End If
            ' have class maintian the array until needed
            cImages.SetOriginalFormat bOverlay
            ImportImageList = True
            
        ElseIf bCompressMethod = 1 Then
            ' compressed image bytes only. When uncompressed, done directly to DIB
            If cCompressor Is Nothing Then Set cCompressor = New cHuffman
            Set cImages = New c32bppDIB
            ImportImageList = cCompressor.DecodeByte(bImport(), lPointer, bImport(), cImages)
            Set cCompressor = Nothing
            
        Else
            ' no delay load, create DIB now
            Set cImages = New c32bppDIB
            If bCompressMethod = 2 Then
                ImportImageList = cImages.LoadPicture_Stream(bOverlay(), , , 1, UBound(bOverlay))
            Else
                ImportImageList = cImages.LoadPicture_Stream(bImport(), , , lPointer, UBound(bImport) - lPointer + 1)
            End If
        End If
        
        ' when compiled, add this efficiency enhancement. Can crash if turned on in IDE
        Debug.Assert IsUnCompiled(bInIDE)
        cImages.KeepGDIplusActive = Not bInIDE
    End If
    
End Function

Private Function IsUnCompiled(bInIDE As Boolean) As Boolean
    bInIDE = True
    IsUnCompiled = True
End Function

Private Function GetMaxImageCount(imgCx As Long, imgCy As Long)
    ' returns maximum number of images for a specific image list
    If imgCx > 0 Then
        If imgCy > 0 Then GetMaxImageCount = (1024000 \ (imgCx * imgCy))
    End If
End Function

Private Function ExtractIconLocal(srcList As c32bppDIB, ByVal iIndex As Long, pBitDepth As Long, ByVal imgWidth As Long, ByVal imgHeight As Long) As Long

    Dim bUnique As Boolean, bPerPixelAlpha As Boolean
    Dim X As Long, y As Long
    Dim bIcon() As Byte, Pow2(0 To 8) As Integer
    Dim bSource() As Byte, oSA As SafeArray
    Dim lSource() As Long, oSALong As SafeArray                             ' source overlay use (arrays never initialized)
    Dim cPalette(1 To 256) As Long, ColorCount As Long, palColor As Byte    ' palettized icon use
    Dim scanWidth As Long, pScanWidth As Long, mScanWidth As Long           ' source, palette, mask scan widths
    Dim pIndex As Long, aPtr As Long                                        ' array pointers
    Dim pOffset As Long, mOffset As Long                                    ' palette & mask array offsets
    Dim maskByte As Byte, maskPos As Long                                   ' bit shifting use
    Dim resLength As Long                                                   ' icon bit depth & overall size
    
    Pow2(0) = 1 ' create a power of two lookup table
    For X = 1 To 8: Pow2(X) = Pow2(X - 1) * 2: Next
    
    scanWidth = imgWidth * 4&                            ' scanwidth of source image
    mScanWidth = iparseByteAlignOnWord(1, imgWidth)      ' scanwidth for dest mask bytes
    iIndex = iIndex - 1
    
    ' presize the array to max expected size; will be significantly smaller if image can be paletted
    mOffset = scanWidth * imgHeight + 39 + 1024          ' identify where in array the maskbytes will begin
    ReDim bIcon(0 To mOffset + mScanWidth * imgHeight)   ' oversize array vs re-dimming later
    
    ' create a byte overlay array so we can look at the alpha bytes only
    iparseOverlayHost_Byte bSource(), VarPtr(oSA), 2, srcList.Height, srcList.scanWidth, srcList.BitsPointer
    ' add Long array overlay so we can easily look at 4 byte colors
    iparseOverlayHost_Long lSource(), VarPtr(oSALong), 2, srcList.Height, srcList.Width, srcList.BitsPointer
    
    
    ' create mask and determine if paletting is possible
    aPtr = iIndex * scanWidth + 3                       ' position for first alpha byte in source
    For y = 0 To imgHeight - 1
        maskPos = 7                                     ' position for 1st mask bit
        pIndex = mOffset + (mScanWidth * y)             ' position in array for this row of source pixels
        For X = aPtr To aPtr + scanWidth - 4 Step 4&
            If bSource(X, y) = 0 Then                   ' set bit to 1 if pixel is transparent
                maskByte = maskByte Or Pow2(maskPos)
            Else                                        ' determine if per-pixel alphablending is in use
                If bSource(X, y) < 255 Then bPerPixelAlpha = True   ' will be converted to 32bpp icon
            End If
            If maskPos = 0 Then                         ' mask shifted out
                bIcon(pIndex) = maskByte                ' append mask byte
                pIndex = pIndex + 1                     ' move to next mask index
                maskPos = 7                             ' reset mask bit pos
                maskByte = 0                            ' reset byte
            Else
                maskPos = maskPos - 1                   ' shift bit mask
            End If
        Next
        If maskByte Then                                ' widths not evenly divisible by 4 will have unfilled bits
            bIcon(pIndex) = maskByte                    ' write the final byte for this row
            maskByte = 0
        End If
    Next
    
    ' user requests paletted icon, see if it is doable; not doable if bSimpleAlpha=False
    aPtr = iIndex * imgWidth                            ' position of of first source pixel
    If Not bPerPixelAlpha Then
        ' here we are going to count colors, if we exceed 256 then we abort and cannot palettize
        For y = 0 To imgHeight - 1
            For X = aPtr To aPtr + imgWidth - 1         ' log each color
                pIndex = FindColor(cPalette, lSource(X, y), ColorCount, bUnique)
                If bUnique Then                         ' new color for the palette
                    If ColorCount = 256 Then            ' exceeded 256
                        ColorCount = 0&                 ' flag indicating can't palettize
                        y = imgHeight                   ' aborts outer loop
                        Exit For
                    Else
                        ColorCount = ColorCount + 1     ' increment palette count
                        If pIndex < ColorCount Then     ' keep palette sorted for searching
                            CopyMemory cPalette(pIndex + 1), cPalette(pIndex), (ColorCount - pIndex) * 4&
                        End If
                        cPalette(pIndex) = lSource(X, y) ' add color to palette
                    End If
                End If
            Next
        Next
    End If
    
    ' now build the color portion of the icon
    If ColorCount = 0& Then
    
        pIndex = 40                                     ' position in icon array where color data starts
        If bPerPixelAlpha Then                          ' can create true color icon (supported by Win2K+)

' 32bpp icons
            pBitDepth = 32                              ' icon bit depth & size of color data needed
            pScanWidth = scanWidth
            aPtr = iIndex * scanWidth                   ' position of of first source pixel (Byte)
            ' this should be the easiest, but we need to remove pre-multiplication
            For y = 0& To imgHeight - 1&
                For X = aPtr To aPtr + scanWidth - 1 Step 4
                    Select Case bSource(X + 3&, y)
                    Case 255
                        CopyMemory bIcon(pIndex), bSource(X, y), 4&
                    Case 0 ' do nothing
                    Case Else
                        palColor = bSource(X + 3&, y)
                        bIcon(pIndex) = (255& * bSource(X, y) \ palColor)
                        bIcon(pIndex + 1&) = (255& * bSource(X + 1&, y) \ palColor)
                        bIcon(pIndex + 2&) = (255& * bSource(X + 2&, y) \ palColor)
                        bIcon(pIndex + 3&) = palColor
                    End Select
                    pIndex = pIndex + 4&
                Next
            Next
            aPtr = iIndex * imgWidth                    ' position of of first source pixel (Long)
        Else

' 24bpp Icons
            pBitDepth = 24                              ' icon bit depth & calc size of color data needed
            pScanWidth = iparseByteAlignOnWord(pBitDepth, imgWidth)
            For y = 0 To imgHeight - 1
                pIndex = 40 + y * pScanWidth            ' move to position for current row
                For X = aPtr To aPtr + imgWidth - 1
                    CopyMemory bIcon(pIndex), lSource(X, y), 3& ' copy 3 color bytes & increment pointer
                    pIndex = pIndex + 3
                Next
            Next
        End If
        resLength = pScanWidth * imgHeight + mScanWidth * imgHeight + 40  ' calc actual size of icon data to be used
        ' shift the mask bytes to their proper array position
        CopyMemory bIcon(pScanWidth * imgHeight + 40), bIcon(mOffset), mScanWidth * imgHeight
        
    Else    ' we can palettize, create the palette indexes
        
        If ColorCount > 16 Then
' 8bpp icons
             pBitDepth = 8              ' use 256 color palette (2^8)
        ElseIf ColorCount > 2 Then
' 4bpp icons
            pBitDepth = 4               ' use 16 color palette (2^4)
        Else
' 1bpp icons
            pBitDepth = 1               ' use 2 color palette  (2^1)
        End If
' 2bpp icons can be created, but Windows doesn't seem to support them. pBitDepth=2
        
        pScanWidth = iparseByteAlignOnWord(pBitDepth, imgWidth)      ' palette indexes scanwidth
        pOffset = Pow2(pBitDepth) * 4 + 40                           ' start of pixel data (after palette & bitmapinfo header)
        pIndex = pScanWidth * imgHeight + pOffset                    ' start of mask data
        CopyMemory bIcon(pIndex), bIcon(mOffset), mScanWidth * imgHeight ' shift mask bytes to their final position
        resLength = pIndex + mScanWidth * imgHeight                  ' calc actual size of icon data to be used
        
        ' build the palette indexes
        For y = 0 To imgHeight - 1
            pIndex = pOffset + (pScanWidth * y)                     ' position for next row of palette indexes
            ' handle 8bpp separate, less code needed
            If pBitDepth = 8 Then
                For X = aPtr To aPtr + imgWidth - 1
                    bIcon(pIndex) = FindColor(cPalette, lSource(X, y), ColorCount, False) - 1
                    pIndex = pIndex + 1
                Next
            Else            ' 1 or 4 bit icons
                maskPos = 8 - pBitDepth                             ' position in palette's bit mask
                For X = aPtr To aPtr + imgWidth - 1
                    ' add to the palette index byte
                    palColor = FindColor(cPalette, lSource(X, y), ColorCount, False) - 1
                    maskByte = maskByte Or (Pow2(maskPos) * palColor)
                    If maskPos = 0 Then                             ' palette shifted out
                        bIcon(pIndex) = maskByte                    ' append palette index
                        maskPos = 8 - pBitDepth                     ' reset palette bit pos
                        maskByte = 0                                ' reset palette byte
                        pIndex = pIndex + 1                         ' move to next palette index
                    Else
                        maskPos = maskPos - pBitDepth               ' shift bit mask
                    End If
                Next
                If maskByte Then                                    ' write remainder if any
                    bIcon(pIndex) = maskByte
                    maskByte = 0
                End If
            End If
        Next
        
        CopyMemory bIcon(40), cPalette(1), Pow2(pBitDepth) * 4      ' transfer palette to icon array
        ' below is the bitmapinfo header's biClrUsed & biClrImportant members
        ' but browsing icon formats, this does not seem to be needed
        CopyMemory bIcon(32), Pow2(pBitDepth), 2&           ' colors used
        CopyMemory bIcon(36), ColorCount, 4&                ' colors important
    
    End If
    iparseOverlayHost_Byte bSource(), 0, 0, 0, 0, 0         ' remove overlays
    iparseOverlayHost_Long lSource, 0, 0, 0, 0, 0
    
    ' build the bitmapinfo header
    CopyMemory bIcon(0), 40&, 4&                            ' header size
    CopyMemory bIcon(4), imgWidth, 4&                       ' width
    CopyMemory bIcon(8), CLng(imgHeight * 2), 4&            ' height * 2
    CopyMemory bIcon(12), 1, 2&                             ' planes
    CopyMemory bIcon(14), pBitDepth, 2&                     ' bitcount
    CopyMemory bIcon(20), CLng(resLength - 40), 4&          ' color bits + mask bits + palette
        
    ' return result
    ExtractIconLocal = CreateIconFromResourceEx(bIcon(0), resLength, 1, &H30000, 0, 0, 0&)


End Function

Private Function FindColor(ByRef PaletteItems() As Long, Color As Long, Count As Long, isNew As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between 1 and 20 iterations
    
    ' [in] PaletteItems(). Long Array to search within. Array must be 1-bound
    ' [in] Color. A value to search for. Order is always ascending
    ' [in] Count. Number of items in PaletteItems() to compare against
    ' [out] isNew. If Color not found, isNew is True else False
    ' [out] Return value: The Index where Color was found or where the new Color should be inserted

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If Count = 0& Then
        FindColor = 1&
        isNew = True
        Exit Function
    End If
    
    UB = Count
    LB = 1&
    
    Do Until LB > UB
        newIndex = LB + ((UB - LB) \ 2&)
        If PaletteItems(newIndex) = Color Then
            Exit Do
        ElseIf PaletteItems(newIndex) > Color Then ' new color is lower in sort order
            UB = newIndex - 1&
        Else ' new color is higher in sort order
            LB = newIndex + 1&
        End If
    Loop

    If LB > UB Then  ' color was not found
        If Color > PaletteItems(newIndex) Then newIndex = newIndex + 1&
        isNew = True
    Else
        isNew = False
    End If
    
    FindColor = newIndex

End Function

