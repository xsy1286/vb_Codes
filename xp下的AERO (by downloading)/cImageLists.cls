VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' For updates: visit http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=69621&lngWId=1
' Review Usage_ImageList.Rtf for available methods and properties

' This class-only version of the imagelist usercontrol can only be used during run-time (obviously)
' The classes are identical to the usercontrol except some functions have been removed that pertain
' only to the property page used with that usercontrol.

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' IMAGE LIST RESTRICTIONS/ABILITIES:

' These classes will not be compatible with VB controls that expect the
' VB ImageList control. Use these classes for custom-drawn usercontrols and objects
' like buttons, grids, etc.  If the control has an hDC, then these classes
' can render/draw to it.

' You should call ManageGDItoken from within your project before any image list is created or imported.
' This will setup a GDI+ safe environment that can prevent GDI+ crashes when VB IDE is ENDed abruptly.
' Additionally, when compiled and GDI+ is available, additional optimiziations within the classes are
' "turned on" to allow better performance.  See KeepGDIplusActive property in c32bppDIB class and calls
' to it from within this class and/or cImageList class.

' See the Usage_ImageList.Rtf file provided for an overview of capabilities and methods/properties
' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' Extraneous Routines
' -------------------
' The following routines can be removed from the c32bppDIB class if you like. The reason
' they are not removed here is because you may have these classes within a project that is
' already using the c32bppDIB suite.

' If you are using pre-existing c32bppDIB suite in your project and will include these
' classes also, you MUST REPLACE your c32bppDIB, cGDIplus classes and modParsers bas
' files with the ones in this project. They have been slightly modifiied to work with
' the cImageList & cImageLists classes.  Otherwise, routines that can be removed are:
' - From c32bppDIB:
'       BlendToColor , CreateDropShadow, CreateRegion, GetAlphaMask
'       MakeGrayScale , MakeImageInverse, MirrorImage, PngPropertyGet
'       PngPropertySet , RenderDropShadow_JIT, SaveToFile, SaveToFile_PNG
'       SetDIBbits , SetPixel, TrimImage
' - From modParsers:
'       iparseCreateShapedRegion , iparseDeleteFile, local_CreatePartialRegion, local_CreateWin98Region
' - From cGDIplus
'       SaveToJPG

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long


' ***** THUNK APIs and VARIABLES (IDE GDI+ Usage, See ManageGDIToken)
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const GWL_WNDPROC As Long = -4
Private Const GW_OWNER As Long = 4
Private Const WS_CHILD As Long = &H40000000

'Public Enum eGrayScaleFormulas
'    gsclNTSCPAL = 0             ' R=R*.299, G=G*.587, B=B*.114 - Default
'    gsclCCIR709 = 1             ' R=R*.213, G=G*.715, B=B*.072
'    gsclSimpleAvg = 2           ' R,G, and B = (R+G+B)/3
'    gsclRedMask = 3             ' uses only the Red sample value: RGB = Red / 3
'    gsclGreenMask = 4           ' uses only the Green sample value: RGB = Green / 3
'    gsclBlueMask = 5            ' uses only the Blue sample value: RGB = Blue / 3
'    gsclRedGreenMask = 6        ' uses Red & Green sample value: RGB = (Red+Green) / 2
'    gsclBlueGreenMask = 7       ' uses Blue & Green sample value: RGB = (Blue+Green) / 2
'    gsclNone = -1
'End Enum
'
'Public Enum eConstants          ' See ImageSizesFromStream
'    HIGH_COLOR = &HFFFF00
'    TRUE_COLOR = &HFF000000
'    TRUE_COLOR_ALPHA = &HFFFFFFFF
'End Enum
'
Private Type ImageListKey
    Key As String               ' imagelist key (required)
    Index As Long               ' provides key to index cross-reference
End Type

Private Const m_lvMagicNr As Long = &H4C56494C ' LVIL: LaVolpe ImageList

Private m_Items() As cImageList         ' collection of imagelists
Private m_Keys() As ImageListKey        ' collection of keys used
Private m_CountItems As Long            ' number of items
Private m_GDItoken As Long              ' GDI+ token usable for life of usercontrol
Private m_CompressOpt As Boolean        ' option to compress image (1) key/tag information (2)

Private Const m_Version As Long = 2
'^^ when new version implemented, must include new version code in cImageList.ExportImageList

Public Property Get ImageList(Index As Variant) As cImageList

    ' Property returns an Image List object which can be used to add/delete images and render images
    ' Parameters:
    '   Index :: Key or Index that identifies the image list. All imagelists have keys. Keys can be numeric
        
    Dim iIndex As Long
    iIndex = KeyToIndex(Index, 0&)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.ImageList Property", "Invalid Key"
        Exit Property
    End If
    Set ImageList = m_Items(iIndex)
    
End Property

Public Function CreateImageList(ByVal ImageWidth As Long, ByVal ImageHeight As Long, ByVal Key As String, _
                                Optional InitImgCount As Long = 1&, _
                                Optional DelayLoaded As Boolean = False) As cImageList
    
    ' Function Creates a new imagelist containing image sizes of ImageWidth x ImageHeight
    ' Parameters:
    '   ImageWidth :: size of individual images within the list
    '   ImageHeight:: size of individual images within the list
    '   Key :: Key to reference this image list
    '   DelayLoaded :: no effect for ImageLists created during runtime. However, during runtime
    '       if an ImageList is created and later stored to disk (Exported) and eventually
    '       Imported, then DelayLoaded will be honored
    
    Dim kIndex As Long
    Dim bIsNew As Boolean
    Dim newList As cImageList
    
    ' first see if a passed key is already registered
    kIndex = FindKey(Key, bIsNew, m_CountItems)
    If bIsNew = False Then
        Err.Raise 1962, "ImageList.CreateImageList", "Key is already associated with another image list in the collection"
        Exit Function    ' duplicated key
    End If
        
    ' ensure imagelist does not exceed max values
    Set newList = New cImageList
    If InitImgCount < 1 Then InitImgCount = 1
    If newList.InitializeImageList(InitImgCount, ImageWidth, ImageHeight) = False Then
        Err.Raise 1968, "ImageList.CreateImageList", "Image list cannot support images of that size"
        Exit Function
    End If
        
    ' resize arrays if needed, keep small buffer
    m_CountItems = m_CountItems + 1
    If m_CountItems = 1& Then
        ReDim m_Items(1 To m_CountItems + 3)
        ReDim m_Keys(1 To m_CountItems + 3)
    ElseIf m_CountItems > UBound(m_Items) Then
        ReDim Preserve m_Items(1 To m_CountItems + 3)
        ReDim Preserve m_Keys(1 To m_CountItems + 3)
    End If
    
    ' add the key to the collection
    newList.ImageClass.gdiToken = m_GDItoken
    newList.DelayLoad = DelayLoaded
    Set m_Items(m_CountItems) = newList
    AddKey Key, kIndex, m_CountItems
    
    
    Set CreateImageList = newList

End Function

Public Function DestroyImageList(ByVal Index As Variant) As Boolean
    
    ' Function removes an image form the image list
    ' Parameters:
    '   Index :: Key or Index that identifies the image list. All imagelists have keys. Keys can be numeric

    Dim iIndex As Long
    Dim kIndex As Long
    
    iIndex = KeyToIndex(Index, kIndex)  ' Validate passed Key
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.DestroyImageList", "Invalid Key or Index"
        Exit Function
    End If
   
    If m_CountItems = 1& Then   ' deleting last one, purge completely
        Call DestroyAll
    Else
        If kIndex = 0& Then kIndex = IndexToKey(iIndex)
        Call RemoveKey(kIndex)  ' remove the key from the listing
        Set m_Items(iIndex) = Nothing   ' free memory as needed
        If iIndex < m_CountItems Then
            CopyMemory ByVal VarPtr(m_Items(iIndex)), ByVal VarPtr(m_Items(iIndex + 1)), (m_CountItems - iIndex) * 4&
            CopyMemory ByVal VarPtr(m_Items(m_CountItems)), 0&, 4&  ' reset the object pointer
        End If
        m_CountItems = m_CountItems - 1
        For kIndex = 1 To m_CountItems
            If m_Keys(kIndex).Index > iIndex Then m_Keys(kIndex).Index = m_Keys(kIndex).Index - 1
        Next
    
        ' array size management; don't let it remain too big after many items were deleted
        If UBound(m_Items) - m_CountItems > 5 Then
            ReDim Preserve m_Items(1 To m_CountItems + 3)
            ReDim Preserve m_Keys(1 To m_CountItems + 3)
        End If
    End If
    DestroyImageList = True

End Function

Public Function ImportImageList(inData() As Byte, Optional Key As String) As cImageList

    ' Function parses the inData() to create an image list
    ' Parameters:
    '   inData() :: an array of bytes that comprise a formatted imagelist
    '       When an image list is exported to file, simply read the file bytes and pass them here
    '   Key [out] :: the key assigned to the imported image list
    
    Dim ValidationOffset As Long, ImageOffset As Long
    Dim vVersion As Long, vData As Long, lPointer As Long
    Dim lBoundOffset As Long
    
    ' validation first
    If iparseArrayProps(VarPtrArray(inData), vData) = 0& Then Exit Function
    If vData > 1& Then Exit Function
    
    ' validate the array is our imagelist
    lPointer = LBound(inData)
    lBoundOffset = lPointer - 1
    If Abs(UBound(inData) - LBound(inData)) < 20 Then Exit Function   ' minimal length need to validate
    
    CopyMemory vData, inData(lPointer), 4&                  ' get magic number which is LVIL
    If vData <> m_lvMagicNr Then Exit Function
        
    CopyMemory ValidationOffset, inData(lPointer + 8), 4&   ' location to IL's xtraData offset
    CopyMemory ImageOffset, inData(lPointer + 12), 4&       ' location to IL's image offset
    If UBound(inData) < (ImageOffset - lPointer + 1) Then Exit Function ' invalid data
    
    ' get the number of xtraBytes used by the IL version & add it to the Validation offset
    CopyMemory vData, inData(ValidationOffset + lBoundOffset), 4&
    ' if the ImageOffset does not equal the sum, then invalid data
    If ImageOffset <> (vData + ValidationOffset) Then Exit Function
    
    Dim keyLen As Long, baseKey As String
    Dim kIndex As Long, bOK As Boolean, fIndex As Long
    
    ' appears to be valid data. Get its version
    CopyMemory vVersion, inData(lPointer + 4), 4&
    lPointer = lPointer + 16                                ' move to imagelist name
    If vVersion > m_Version Then                            ' version > our version, so we can't import future version data
        Err.Raise 1977, "ImageList>ImportImageList", "Cannot import image list. It has been formatted for a newer version of this control."
        Exit Function
    End If
    
    m_CountItems = m_CountItems + 1                         ' create a new structure for the imagelist
    If iparseIsArrayEmpty(VarPtrArray(m_Items)) = 0 Then
        ReDim Preserve m_Items(1 To m_CountItems)
        ReDim Preserve m_Keys(1 To m_CountItems)
    ElseIf m_CountItems > UBound(m_Items) Then
        ReDim Preserve m_Items(1 To m_CountItems + 1)
        ReDim Preserve m_Keys(1 To m_CountItems + 1)
    End If
    
    CopyMemory keyLen, inData(lPointer), 4&                     ' length of imagelist key
    Key = Space$(keyLen \ 2)                                    ' create string neeeded
    CopyMemory ByVal StrPtr(Key), inData(lPointer + 4), keyLen  ' copy the key
    Set m_Items(m_CountItems) = New cImageList                  ' create a new imagelist
                    
    If m_CountItems = 1 Then                                    ' handle 1st occurence specifically
        m_Keys(m_CountItems).Key = Key                          ' set the key & cross-reference
        m_Keys(m_CountItems).Index = m_CountItems
    Else
        kIndex = FindKey(Key, bOK, m_CountItems - 1)          ' find insert position of key in array
        If Not bOK Then
            baseKey = Key
            Do Until bOK
                fIndex = fIndex + 1
                Key = baseKey & "(" & fIndex & ")"
                kIndex = FindKey(Key, bOK, m_CountItems - 1)  ' find insert position of key in array
            Loop
        End If
        baseKey = Key                                         ' baseKey will be zeroized on AddKey call
        AddKey baseKey, kIndex, m_CountItems                  ' insert the key into the array
    End If
    
    On Error Resume Next
    ' import the image data into the imagelist & calculate the header size
    bOK = False
    bOK = m_Items(m_CountItems).ImportImageList(inData(), lPointer + keyLen + 3 - lBoundOffset, ValidationOffset, vVersion)
    If bOK Then
        m_Items(m_CountItems).ImageClass.gdiToken = m_GDItoken
        Set ImportImageList = m_Items(m_CountItems)
    End If
    If Not bOK Then
        ' failed to import, rollback
        m_Keys(m_CountItems).Key = vbNullString
        m_Keys(m_CountItems).Index = 0&
        Set m_Items(m_CountItems) = Nothing
        m_CountItems = m_CountItems - 1
        If Err Then
            Err.Raise Err.Number, "ImageList.ImportImageList", Err.Description
            Err.Clear
        Else
            Err.Raise 1981, "ImageList.ImportImageList", "Data could not be imported."
        End If
    End If

End Function

Public Function ExportImageList(outData() As Byte, ImageList As Variant) As Boolean

    ' Function exports an imagelist, its key, images, images' keys/tags to an array
    ' Parameters:
    '   outData() :: an array which will be resized on return
    '   ImageList :: the ImageList instance you want exported or its ordinal index or its Key
    ' Return Value: If false, the outData array should be considered invalid
    
    Dim sKey As String, kIndex As Long
    Dim lPointer As Long, buffLen As Long
    Dim ValidationOffset As Long, ImageOffset As Long
    Dim ilExport As cImageList
    Dim bOK As Boolean
    
    If TypeOf ImageList Is cImageList Then
        kIndex = PointerToIndex(ObjPtr(ImageList))
        sKey = Me.Key(kIndex)
        Set ilExport = ImageList
    Else
        lPointer = KeyToIndex(ImageList, kIndex)
        If lPointer = 0& Then
            Err.Raise 1961, "ImageList.ExportImageList", "Invalid Key or Index"
            Exit Function
        End If
        If kIndex = 0& Then kIndex = IndexToKey(lPointer)
        sKey = m_Keys(kIndex).Key
        Set ilExport = m_Items(lPointer)
    End If
        
    buffLen = Len(sKey) * 2 + 20
    ' call function to export image list data
    
    On Error Resume Next
    bOK = ilExport.ExportImageList(outData, buffLen, ValidationOffset, ImageOffset, m_CompressOpt)
    If bOK = True Then
        lPointer = LBound(outData)
        CopyMemory outData(lPointer), m_lvMagicNr, 4&
        ' next 4 bytes are for version. 0 indicates version 1
        CopyMemory outData(lPointer + 4), m_Version, 4&
        CopyMemory outData(lPointer + 8), ValidationOffset, 4&  ' used to validate array for import
        CopyMemory outData(lPointer + 12), ImageOffset, 4&      ' position of image in array
        buffLen = buffLen - 20                                  ' length of key
        CopyMemory outData(lPointer + 16), buffLen, 4&          ' store the length & the key
        CopyMemory outData(lPointer + 20), ByVal StrPtr(sKey), buffLen
        ExportImageList = True
    Else
        If Err Then
            Err.Raise Err.Number, "ImageList.ExportImageList", Err.Description
        Else
            Err.Raise 1980, "ImageList.ExportImageList", "Error exporting image list. Ensure adequate permissions and file space exists."
        End If
    End If

End Function

Public Property Get Key(Index As Variant) As String
    
    ' Returns a Key associated with an imagelist instance
    Dim iIndex As Long, kIndex As Long
    iIndex = KeyToIndex(Index, kIndex)
    If iIndex = 0& Then ' error
        Err.Raise 1961, "ImageList.Key", "Invalid Key or Index"
        Exit Property
    End If
    
    If kIndex = 0 Then kIndex = IndexToKey(iIndex)
    Key = m_Keys(kIndex).Key
    
End Property

Public Property Let Key(Index As Variant, ByVal newKey As String)
    
    ' Changes the key for an image list
    ' Parameters:
    '   OldKey :: Key that currently references this image list
    '   NewKey :: Key that will reference this image list
        
    Dim kIndex As Long, bIsNew As Boolean
    Dim iIndex As Long, newIndex As Long
    
    If newKey = vbNullString Then
        Err.Raise 1964, "ImageList.Key", "New Key cannot be blank"
        Exit Property
    End If
    
    iIndex = KeyToIndex(Index, kIndex)
    If iIndex = 0& Then
        Err.Raise 1961, "ImageList.Key", "Invalid Key"
        Exit Property
    End If
    
    If kIndex = 0& Then kIndex = IndexToKey(iIndex)
    newIndex = FindKey(newKey, bIsNew, m_CountItems)
    
    ' validate new key not duplicated
    If bIsNew = False Then      ' key exists
        ' error if not for the same key index
        If newIndex = kIndex Then
            m_Keys(kIndex).Key = newKey ' change it per user request
        Else
            Err.Raise 1962, "ImageList.Key", "Key is already associated with another image list in the collection"
            Exit Property
        End If
    
    ElseIf (newIndex = kIndex) Or (newIndex - kIndex = 1) Then
        ' new key is unique & will be placed in same position as oldkey
        m_Keys(kIndex).Key = newKey
    
    Else
        ' new key is unique and will be placed elsewhere
        RemoveKey kIndex    ' remove old key if any
        AddKey newKey, 0&, iIndex
    End If
    
End Property

Public Property Get ListCount() As Long
    ListCount = m_CountItems    ' returns number of imagelists managed by the control
End Property
Public Property Let ListCount(newCount As Long)
    ' dummy property only, can be called to force an update during design time
    ' the property page uses this to force updates
End Property
Public Property Get TotalImages() As Long
    ' for informational purposes only. Number of images for all imagelists combined
    Dim Index As Long
    Dim nrImages As Long
    For Index = 1 To m_CountItems
        nrImages = nrImages + m_Items(Index).ImageCount
    Next
    TotalImages = nrImages
End Property
Public Property Let TotalImages(newCount As Long)
    ' dummy property only so it is displayed during IDE
End Property

Public Property Get CompressImageList() As Boolean
    ' option to compress imagelist using huffmann compression.
    ' If zLIB or GDI+ are installed, option is ignored and imagelists are always
    ' compressed in PNG format. When not installed, this option will use huffmann
    ' compression which can reduce imagelist to 50% or better vs uncompressed
    ' The size difference will show in the form's frx file. Huffmann can be a
    ' bit slow when control is uncompiled.
    CompressImageList = m_CompressOpt
End Property
Public Property Let CompressImageList(bCompress As Boolean)
    If m_CompressOpt <> bCompress Then
        m_CompressOpt = bCompress
    End If
End Property

Public Property Get CompressionNeeded() As Boolean
    ' returns if image list can be saved to PNG format or not
    ' If it can then property returns False, otherwise,
    ' property returns True. See CompressImageList
    Dim c32bpp As New c32bppDIB
    If c32bpp.isGDIplusEnabled = False Then
        CompressionNeeded = Not c32bpp.isZlibEnabled
    End If
End Property
Public Function MaxImagesForImageList(ByVal ImageWidth As Long, ByVal ImageHeight As Long) As Long
    ' return how many images can fit on an image list where each image
    ' will be of the passed ImageWidth and ImageHeight.
    ' Note: The cImageList class has same function (MaxImages) but its return value
    ' only applies to its own dimensions
    If ImageWidth > 0& Then
        If ImageHeight > 0& Then MaxImagesForImageList = (1024000 \ (ImageWidth * ImageHeight))
    End If
End Function

Public Sub DestroyAll()
    ' destroys all imagelists
    Erase m_Items()
    Erase m_Keys()
    m_CountItems = 0&
End Sub

Public Function IsKeyUnique(inKey As String) As Boolean
    ' Checks if a key is unique among image lists
    ' Parameter
    '   Key :: Key that identifies the image list. All imagelists have keys. Keys can be numeric
    Call FindKey(inKey, IsKeyUnique, m_CountItems)
End Function

Public Function ImageSizesFromStream(inStream() As Byte, arraySizes() As Long) As Long

    ' Function will return the number of images, bit depths and color count contained
    ' within the passed array.  For non multi-resource files, the return value would be one.
    ' If the passed array contains no recognized image, then the return value is zero

    ' arraySizes() will be populated as follows
    ' 1st dimension of the array
    '   0 element: image/icon width
    '   1 element: image/icon height
    '   The next two elements only apply if array contains multiple images/icons
    '   2 element: image/icon bit depth: 1,2,4,8,16,24,32
    '   3 element: image/icon colors: 2,4,16,256,HIGH_COLOR,TRUE_COLOR, TRUE_COLOR_ALPHA
    '   -- HIGH_COLOR, TRUE_COLOR,& TRUE_COLOR_ALPHA are public variables
    ' 2nd dimension of the array is 0 to number of icons - 1
    
    Dim tImage As New c32bppDIB, nrImages As Long
    If tImage.LoadPicture_Stream(inStream, , , , , True) = True Then
        nrImages = tImage.SourceIconSizes(arraySizes)
        If nrImages = 0& Then
            nrImages = 1
            ReDim arraySizes(0 To 3, 0 To 0)
            arraySizes(0, 0) = tImage.Width
            arraySizes(1, 0) = tImage.Height
        End If
        Set tImage = Nothing
    Else
        ReDim arraySizes(0 To 3, -1 To -1)
    End If
    ImageSizesFromStream = nrImages

End Function

' *********************************************************
'       INTERNAL AND PROPERTY-PAGE RELATED ROUTINES
' *********************************************************

Private Sub AddKey(Key As String, KeyIndex As Long, ItemIndex As Long)

    ' Helper function. Adds a key.
    ' FYI: The passed key will return as vbNullString
    If KeyIndex = 0& Then KeyIndex = FindKey(Key, False, m_CountItems - 1)
    
    If KeyIndex < m_CountItems Then ' need to shift array to maintain ASC sort order
        CopyMemory ByVal VarPtr(m_Keys(KeyIndex + 1)), ByVal VarPtr(m_Keys(KeyIndex)), (m_CountItems - KeyIndex) * 8&
    End If
    CopyMemory ByVal VarPtr(m_Keys(KeyIndex).Key), StrPtr(Key), 4& ' move the string vs copying it
    CopyMemory ByVal VarPtr(Key), 0&, 4&    ' release strptr on passed string
    m_Keys(KeyIndex).Index = ItemIndex

End Sub

Private Function KeyToIndex(Key As Variant, KeyIndex As Long) As Long

    ' Helper function to retrieve imagelist item from key/index
    Dim bIsNew As Boolean
    Select Case VarType(Key)
    Case vbString                                       ' Strings are treated as keys
        If Not Key = vbNullString Then
            KeyIndex = FindKey(CStr(Key), bIsNew, m_CountItems)
            If bIsNew = False Then KeyToIndex = m_Keys(KeyIndex).Index
        End If
    Case vbLong, vbInteger, vbDouble, vbByte, vbSingle  ' Numerical varTypes are treated as Indexes
        On Error Resume Next
        KeyToIndex = Int(Key)
        If KeyToIndex < 1 Or KeyToIndex > m_CountItems Then KeyToIndex = vbDefault
        ' KeyIndex is returned as zero. If the item's key is required, the calling
        ' routine will call IndexToKey to obtain it.
        If Err Then Err.Clear
    End Select
    
End Function

Private Function IndexToKey(ItemIndex As Long) As Long
    ' Unlike the cImageList image items, keys for the image lists themselves
    ' are not cross-referenced. This collection should always be relatively
    ' small in number (less than 100) and iterating thru them is quick
    If ItemIndex Then
        Dim Index As Long
        For Index = 1 To m_CountItems
            If m_Keys(Index).Index = ItemIndex Then
                IndexToKey = Index
                Exit For
            End If
        Next
    End If

End Function

Private Function PointerToIndex(classPtr As Long) As Long
    ' When needed, we can get the item index from the object pointer
    ' ExportImageList can call this function
    Dim Index As Long
    For Index = 1 To m_CountItems
        If ObjPtr(m_Items(Index)) = classPtr Then
            PointerToIndex = Index
            Exit For
        End If
    Next

End Function

Private Sub RemoveKey(KeyIndex As Long)
    
    ' Helper function to remove a key
    m_Keys(KeyIndex).Key = vbNullString
    m_Keys(KeyIndex).Index = 0&
    
    If KeyIndex < m_CountItems Then
        CopyMemory ByVal VarPtr(m_Keys(KeyIndex)), ByVal VarPtr(m_Keys(KeyIndex + 1)), (m_CountItems - KeyIndex) * 8&
        ' zero out the final key's string pointer
        CopyMemory ByVal VarPtr(m_Keys(m_CountItems).Key), 0&, 4&
    End If

End Sub

Private Function FindKey(Key As String, ByRef isNew As Boolean, KeyCount As Long) As Long

    ' BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between just 1 and 20 iterations
    
    ' [in] Key. A value to search for. Order is always ascending
    ' [out] isNew. If Key not found, isNew is True else False
    ' [out] Return value: The Index where Key was found or where the new Key should be inserted

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If KeyCount = 0& Then
        newIndex = 1&
        isNew = True
    Else
    
        UB = KeyCount
        LB = 1&
        
        Do Until LB > UB
            newIndex = LB + ((UB - LB) \ 2&)
            Select Case StrComp(Key, m_Keys(newIndex).Key, vbTextCompare)
            Case 0 ' key found
                Exit Do
            Case Is < 0 ' key is < current key
                UB = newIndex - 1&
            Case Else ' key is > current key
                LB = newIndex + 1&
            End Select
        Loop
    
        If LB > UB Then  ' key was not found
            If StrComp(Key, m_Keys(newIndex).Key, vbTextCompare) > 0 Then newIndex = newIndex + 1
            isNew = True
        Else
            isNew = False
        End If
    
    End If
    
    FindKey = newIndex

End Function

Friend Sub ppgImageLists(ImageListArray() As cImageList)
    ' property page use only. Gets imagelist to be modified
    If m_CountItems = 0& Then
        ReDim ImageListArray(0)
    Else
        ImageListArray = m_Items
        ReDim Preserve ImageListArray(1 To m_CountItems)
    End If
End Sub

Public Function ManageGDIToken(hWnd As Long) As Long
        
    ' Return value is an hWnd to the GDI+ IDE-safe monitor.
    ' If return value is zero, then GDI+ is not installed on the O/S
        
    ' So this function will create (if needed) only one API window for the entire
    ' VB instance and remains as long as VB is opened. The API window is a child
    ' of a hidden VB top-level window and when that window closes, so will the
    ' API window which then will release GDI+. By keeping its own reference to
    ' the GDI+ library, there is no way for that library to be unloaded without
    ' the API window closing.
    ' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69434&lngWId=1
        
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    Dim bOK As Boolean
    Dim cGDIp As cGDIPlus
    
    ' find the hidden VB owner window. All VB apps have this (run-time & design-time).
    If hWnd = 0& Then Exit Function
    Do
        hwndGDIsafe = GetParent(hWnd)
        If Not hwndGDIsafe = 0& Then hWnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(hWnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then hWnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(hWnd, 0&, "Static", "GDI+Safe Patch")
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        m_GDItoken = True
        Exit Function                   ' can abort
    End If
    
    On Error Resume Next
    Set cGDIp = New cGDIPlus
    If cGDIp.isGDIplusOk(m_GDItoken, True) = False Then Exit Function
    If Err Then
        m_GDItoken = 0&
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, hWnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = zFnAddr("user32", "CallWindowProcA")                ' function pointer CallWindowProc
            z_Code(3) = zFnAddr("kernel32", "VirtualFree")                  ' function pointer VirtualFree
            z_Code(4) = zFnAddr("kernel32", "FreeLibrary")                  ' function pointer FreeLibrary
            z_Code(5) = m_GDItoken                                          ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = zFnAddr("user32", "SetWindowLongA")                 ' function pointer SetWindowLong
            z_Code(8) = zFnAddr("user32", "SetTimer")                       ' function pointer SetTimer
            z_Code(9) = zFnAddr("user32", "KillTimer")                      ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            
            bOK = True
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        
        Else
        
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
            'MsgBox "Failed to create monitoring window", vbExclamation + vbOKOnly, "Error"
            
        End If
    
    Else
        
      VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
      z_ScMem = 0&
      'MsgBox "Failed to initialize thunk memory", vbExclamation + vbOKOnly, "Error"
        
    End If
    
    If bOK = False Then
        Call cGDIp.InitializeGDIplus(m_GDItoken, True) ' release token if error occurred
        m_GDItoken = 0&
    End If
    
End Function

Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the address of the specified DLL/procedure

  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
  Debug.Assert zFnAddr                                     'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")

End Function

' ***************************************************
'        U S E R C O N T R O L   E V E N T S
' ***************************************************

Private Sub Class_Initialize()
    m_CompressOpt = True
End Sub

Private Sub Class_Terminate()
    Erase m_Items()                     ' destroys all images, frees keys & GDI+ references
    ' WARNING: Do not release the m_gdiToken. Do not add that code to this project
    ' The token is automatically released when VB IDE or compiled app closes.
    ' This is taken care of via the ASM thunks employed when ManageGDIToken was called
End Sub



