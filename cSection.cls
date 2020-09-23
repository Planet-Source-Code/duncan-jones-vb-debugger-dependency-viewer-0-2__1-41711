VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICollectionItem
Dim sKey As String

Public SectionName As String
Public PhysicalAddress As Long
Public VirtualAddress As Long
Public SizeOfRawData As Long
Public PointerToRawData As Long       '\\ This pts to actual offset in file
Public PointerToRelocations As Long
Public PointerToLinenumbers As Long
Public NumberOfRelocations As Integer
Public NumberOfLinenumbers As Integer
 
Private mlCharacteristics As Long

Public Enum ImageSectionCharacteristicsFlags
    IMAGE_SCN_CNT_CODE = &H20                         '\\ Section contains code.
    IMAGE_SCN_CNT_INITIALIZED_DATA = &H40             '\\ Section contains initialized data.
    IMAGE_SCN_CNT_UNINITIALIZED_DATA = &H80           '\\ Section contains uninitialized data.
    IMAGE_SCN_LNK_INFO = &H200                        '\\ Section contains comments or some other type of information.
    IMAGE_SCN_LNK_REMOVE = &H800                      '\\ Section contents will not become part of image.
    IMAGE_SCN_LNK_COMDAT = &H1000                     '\\ Section contents comdat.
    IMAGE_SCN_MEM_FARDATA = &H8000
    IMAGE_SCN_MEM_PURGEABLE = &H20000
    IMAGE_SCN_MEM_16BIT = &H20000
    IMAGE_SCN_MEM_LOCKED = &H40000
    IMAGE_SCN_MEM_PRELOAD = &H80000
    IMAGE_SCN_ALIGN_1BYTES = &H100000                 '\\
    IMAGE_SCN_ALIGN_2BYTES = &H200000                 '\\
    IMAGE_SCN_ALIGN_4BYTES = &H300000                 '\\
    IMAGE_SCN_ALIGN_8BYTES = &H400000                 '\\
    IMAGE_SCN_ALIGN_16BYTES = &H500000                '\\ Default alignment if no others are specified.
    IMAGE_SCN_ALIGN_32BYTES = &H600000                '\\
    IMAGE_SCN_ALIGN_64BYTES = &H700000                '\\
    IMAGE_SCN_LNK_NRELOC_OVFL = &H1000000             '\\ Section contains extended relocations.
    IMAGE_SCN_MEM_DISCARDABLE = &H2000000             '\\ Section can be discarded.
    IMAGE_SCN_MEM_NOT_CACHED = &H4000000              '\\ Section is not cachable.
    IMAGE_SCN_MEM_NOT_PAGED = &H8000000               '\\ Section is not pageable.
    IMAGE_SCN_MEM_SHARED = &H10000000                 '\\ Section is shareable.
    IMAGE_SCN_MEM_EXECUTE = &H20000000                '\\ Section is executable.
    IMAGE_SCN_MEM_READ = &H40000000 '\\ Section is readable.
    IMAGE_SCN_MEM_WRITE = &H80000000  '\\ Section is writeable.
End Enum

Public Property Get ByteAlignment() As Long

Select Case True
Case (mlCharacteristics And IMAGE_SCN_ALIGN_1BYTES)
    ByteAlignment = 1
Case (mlCharacteristics And IMAGE_SCN_ALIGN_2BYTES)
    ByteAlignment = 2
Case (mlCharacteristics And IMAGE_SCN_ALIGN_4BYTES)
    ByteAlignment = 4
Case (mlCharacteristics And IMAGE_SCN_ALIGN_8BYTES)
    ByteAlignment = 8
Case (mlCharacteristics And IMAGE_SCN_ALIGN_16BYTES)
    ByteAlignment = 16
Case (mlCharacteristics And IMAGE_SCN_ALIGN_64BYTES)
    ByteAlignment = 64
Case Else
    ByteAlignment = 32
End Select

End Property

Public Property Let Characteristics(ByVal newval As Long)

    mlCharacteristics = newval
    
End Property



Public Property Get ContainsCode() As Boolean

    ContainsCode = (mlCharacteristics And IMAGE_SCN_CNT_CODE)

End Property

Public Property Get ContainsInitialisedData() As Boolean

    ContainsInitialisedData = (mlCharacteristics And IMAGE_SCN_CNT_INITIALIZED_DATA)

End Property


Public Property Get ContainsLinkInfo() As Boolean

    ContainsLinkInfo = (mlCharacteristics And IMAGE_SCN_LNK_INFO)

End Property

Public Property Get ContainsUninitialisedData() As Boolean

    ContainsUninitialisedData = (mlCharacteristics And IMAGE_SCN_CNT_UNINITIALIZED_DATA)

End Property


Public Property Get Executable() As Boolean

    Executable = (mlCharacteristics And IMAGE_SCN_MEM_EXECUTE)
    
End Property

Public Property Get Readable() As Boolean

    Readable = (mlCharacteristics And IMAGE_SCN_MEM_READ)

End Property


Public Property Get Writeable() As Boolean

    Writeable = (mlCharacteristics And IMAGE_SCN_MEM_WRITE)

End Property


Public Property Get ICollectionItem_ClassName() As String

ICollectionItem_ClassName = "cSection"

End Property

Public Property Get ICollectionItem_Key() As String

If sKey = "" Then
    sKey = ICollectionItem_ClassName & DebugMain.MakeUniqueKey
End If
ICollectionItem_Key = sKey

End Property
