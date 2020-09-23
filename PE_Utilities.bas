Attribute VB_Name = "PE_Utilities"
Option Explicit

'\\ Portable Executable Format descriptions, from WinNT.h

Public Enum ImageSignatureTypes
    IMAGE_DOS_SIGNATURE = &H5A4D     ''\\ MZ
    IMAGE_OS2_SIGNATURE = &H454E     ''\\ NE
    IMAGE_OS2_SIGNATURE_LE = &H454C  ''\\ LE
    IMAGE_VXD_SIGNATURE = &H454C     ''\\ LE
    IMAGE_NT_SIGNATURE = &H4550      ''\\ PE00
End Enum

Private Type IMAGE_DOS_HEADER
    e_magic As Integer   ''\\ Magic number
    e_cblp As Integer    ''\\ Bytes on last page of file
    e_cp As Integer      ''\\ Pages in file
    e_crlc As Integer    ''\\ Relocations
    e_cparhdr As Integer ''\\ Size of header in paragraphs
    e_minalloc As Integer ''\\ Minimum extra paragraphs needed
    e_maxalloc As Integer ''\\ Maximum extra paragraphs needed
    e_ss As Integer    ''\\ Initial (relative) SS value
    e_sp As Integer    ''\\ Initial SP value
    e_csum As Integer  ''\\ Checksum
    e_ip As Integer  ''\\ Initial IP value
    e_cs As Integer  ''\\ Initial (relative) CS value
    e_lfarlc As Integer ''\\ File address of relocation table
    e_ovno As Integer ''\\ Overlay number
    e_res(0 To 3) As Integer ''\\ Reserved words
    e_oemid As Integer ''\\ OEM identifier (for e_oeminfo)
    e_oeminfo As Integer ''\\ OEM information; e_oemid specific
    e_res2(0 To 9) As Integer ''\\ Reserved words
    e_lfanew As Long ''\\ File address of new exe header
End Type

Public Enum ImageMachineTypes
    IMAGE_FILE_MACHINE_I386 = &H14C   ''\\ Intel 386.
    IMAGE_FILE_MACHINE_R3000 = &H162  ''\\ MIPS little-endian,= &H160 big-endian
    IMAGE_FILE_MACHINE_R4000 = &H166  ''\\ MIPS little-endian
    IMAGE_FILE_MACHINE_R10000 = &H168  ''\\ MIPS little-endian
    IMAGE_FILE_MACHINE_WCEMIPSV2 = &H169  ''\\ MIPS little-endian WCE v2
    IMAGE_FILE_MACHINE_ALPHA = &H184      ''\\ Alpha_AXP
    IMAGE_FILE_MACHINE_POWERPC = &H1F0    ''\\ IBM PowerPC Little-Endian
    IMAGE_FILE_MACHINE_SH3 = &H1A2   ''\\ SH3 little-endian
    IMAGE_FILE_MACHINE_SH3E = &H1A4  ''\\ SH3E little-endian
    IMAGE_FILE_MACHINE_SH4 = &H1A6   ''\\ SH4 little-endian
    IMAGE_FILE_MACHINE_ARM = &H1C0   ''\\ ARM Little-Endian
    IMAGE_FILE_MACHINE_IA64 = &H200  ''\\ Intel 64
End Enum

Private Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Private Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Private Type IMAGE_OPTIONAL_HEADER_NT
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

Public Enum ImageTypeFlags
    IMAGE_FILE_RELOCS_STRIPPED = &H1      ''\\ Relocation info stripped from file.
    IMAGE_FILE_EXECUTABLE_IMAGE = &H2     ''\\ File is executable  (i.e. no unresolved externel references).
    IMAGE_FILE_LINE_NUMS_STRIPPED = &H4   ''\\ Line nunbers stripped from file.
    IMAGE_FILE_LOCAL_SYMS_STRIPPED = &H8  ''\\ Local symbols stripped from file.
    IMAGE_FILE_AGGRESIVE_WS_TRIM = &H10   ''\\ Agressively trim working set
    IMAGE_FILE_LARGE_ADDRESS_AWARE = &H20 ''\\ App can handle >2gb addresses
    IMAGE_FILE_BYTES_REVERSED_LO = &H80   ''\\ Bytes of machine word are reversed.
    IMAGE_FILE_32BIT_MACHINE = &H100      ''\\ 32 bit word machine.
    IMAGE_FILE_DEBUG_STRIPPED = &H200     ''\\ Debugging info stripped from file in .DBG file
    IMAGE_FILE_REMOVABLE_RUN_FROM_SWAP = &H400  ''\\ If Image is on removable media, copy and run from the swap file.
    IMAGE_FILE_NET_RUN_FROM_SWAP = &H800  ''\\ If Image is on Net, copy and run from the swap file.
    IMAGE_FILE_SYSTEM = &H1000            ''\\ System File.
    IMAGE_FILE_DLL = &H2000               ''\\ File is a DLL.
    IMAGE_FILE_UP_SYSTEM_ONLY = &H4000    ''\\ File should only be run on a UP machine
    IMAGE_FILE_BYTES_REVERSED_HI = &H8000 ''\\ Bytes of machine word are reversed.
End Enum

Public Enum ImageDataDirectoryIndexes
    IMAGE_DIRECTORY_ENTRY_EXPORT = 0  ''\\ Export Directory
    IMAGE_DIRECTORY_ENTRY_IMPORT = 1  ''\\ Import Directory
    IMAGE_DIRECTORY_ENTRY_RESOURCE = 2 ''\\ Resource Directory
    IMAGE_DIRECTORY_ENTRY_EXCEPTION = 3   ''\\ Exception Directory
    IMAGE_DIRECTORY_ENTRY_SECURITY = 4   ''\\ Security Directory
    IMAGE_DIRECTORY_ENTRY_BASERELOC = 5  ''\\ Base Relocation Table
    IMAGE_DIRECTORY_ENTRY_DEBUG = 6   ''\\ Debug Directory
    IMAGE_DIRECTORY_ENTRY_ARCHITECTURE = 7   ''\\ Architecture Specific Data
    IMAGE_DIRECTORY_ENTRY_GLOBALPTR = 8  ''\\ RVA of GP
    IMAGE_DIRECTORY_ENTRY_TLS = 9  ''\\ TLS Directory
    IMAGE_DIRECTORY_ENTRY_LOAD_CONFIG = 10    ''\\ Load Configuration Directory
    IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT = 11   ''\\ Bound Import Directory in headers
    IMAGE_DIRECTORY_ENTRY_IAT = 12  ''\\ Import Address Table
    IMAGE_DIRECTORY_ENTRY_DELAY_IMPORT = 13   ''\\ Delay Load Import Descriptors
End Enum

'\\ -- Finding a module's imports information....
Private Type IMAGE_IMPORT_DESCRIPTOR
    lpImportByName As Long ''\\ 0 for terminating null import descriptor
    TimeDateStamp As Long  ''\\ 0 if not bound,
                           ''\\ -1 if bound, and real date\time stamp
                           ''\\ in IMAGE_DIRECTORY_ENTRY_BOUND_IMPORT (new BIND)
                           ''\\ O.W. date/time stamp of DLL bound to (Old BIND)
    ForwarderChain As Long ''\\ -1 if no forwarders
    lpName As Long
    lpFirstThunk As Long ''\\ RVA to IAT (if bound this IAT has actual addresses)
End Type

Private Type IMAGE_IMPORT_BY_NAME
    Ordinal As Integer
    Name As Byte '\\This is the null terminated ascii name of the function...
End Type

'\\ --Image sections.......
Private Type IMAGE_SECTION_HEADER
    ImageName(0 To 7) As Byte
    PhysicalAddress As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long       '\\ This pts to actual offset in file
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type

'\\ --- Exports directory..........................
Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    lpName As Long
    Base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    lpAddressOfFunctions As Long    '\\ Three parrallel arrays...(LONG)
    lpAddressOfNames As Long        '\\ (LONG)
    lpAddressOfNameOrdinals As Long '\\ (INTEGER)
End Type

'\\ -- RESOURCES Directory
Private Type IMAGE_RESOURCE_DIRECTORY
    Characteristics As Long '\\Seems to be always zero?
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    NumberOfNamedEntries As Integer
    NumberOfIdEntries As Integer
End Type

Private Type IMAGE_RESOURCE_DIRECTORY_ENTRY
    dwName As Long
    dwDataOffset As Long
End Type

'\\ --Reading and writing files....
Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Private Declare Function ReadFile Lib "kernel32" (ByVal hfile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function ReadFileLong Lib "kernel32" Alias "ReadFile" (ByVal hfile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal hfile As Long, lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long

Public Enum FileOffsetTypes
    FILE_BEGIN = 0&
    FILE_CURRENT = 1&
    FILE_END = 2&
End Enum
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hfile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As FileOffsetTypes) As Long


Declare Function ReadProcessMemoryLong Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Public Sub FillDLLModuleInforFromFile(ByVal hfile As Long, ByVal cmod As cModule)

Dim BytesRead As Long

'\\ Skip to the start of the PE file definition
Call SkipDOSStub(hfile)
Call SetFilePointer(hfile, 4, 0, FILE_CURRENT)

'\\ Now read in the image file header
Dim imhModule As IMAGE_FILE_HEADER
Call ReadFileLong(hfile, VarPtr(imhModule), Len(imhModule), BytesRead, ByVal 0&)
If Err.LastDllError Then
    Debug.Print LastSystemError
Else
    '\\ That read OK - now read the optional header
    Dim imhoModule As IMAGE_OPTIONAL_HEADER
    If imhModule.SizeOfOptionalHeader >= Len(imhoModule) Then
        Call ReadFileLong(hfile, VarPtr(imhoModule), Len(imhoModule), BytesRead, ByVal 0&)
        If Err.LastDllError Then
            Debug.Print LastSystemError
        Else
            With cmod.Image
                .Machine = imhModule.Machine
                .MajorLinkerVersion = imhoModule.MajorLinkerVersion
                .MinorLinkerVersion = imhoModule.MinorLinkerVersion
            End With
        End If
    End If
    Dim imhoModuleNT As IMAGE_OPTIONAL_HEADER_NT
    If imhModule.SizeOfOptionalHeader - Len(imhoModule) >= Len(imhoModuleNT) Then
        Call ReadFileLong(hfile, VarPtr(imhoModuleNT), Len(imhoModuleNT), BytesRead, ByVal 0&)
        If Err.LastDllError Then
            Debug.Print LastSystemError
        Else
            cmod.Image.BaseAddress = imhoModuleNT.ImageBase
            
            Call ProcessExportTable(imhoModuleNT.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT), cmod.Image)
            Call ProcessImportTable(imhoModuleNT.DataDirectory(IMAGE_DIRECTORY_ENTRY_IMPORT), cmod.Image)


            '\\ need to process the sections....
            Call ProcessSections(hfile, imhModule.NumberOfSections, cmod.Image)
        End If
    End If
End If

mdifrmDebuggermain.RefreshModulesList

End Sub

' see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dndebug/html/msdn_peeringpe.asp
Public Sub FillProcessInfoFromFile(ByVal hfile As Long, ByVal cProc As cProcess)

Dim BytesRead As Long

Dim filepointer As INT64

'\\ Skip to the start of the PE file definition
Call SkipDOSStub(hfile)
Call SetFilePointer(hfile, 4, 0, FILE_CURRENT)

'\\ Now read in the image file header
Dim imhProcess As IMAGE_FILE_HEADER
Call ReadFileLong(hfile, VarPtr(imhProcess), Len(imhProcess), BytesRead, ByVal 0&)
If Err.LastDllError Then
    Debug.Print LastSystemError
Else
    '\\ That read OK - now read the optional header
    Dim imhoProc As IMAGE_OPTIONAL_HEADER
    If imhProcess.SizeOfOptionalHeader >= Len(imhoProc) Then
        Call ReadFileLong(hfile, VarPtr(imhoProc), Len(imhoProc), BytesRead, ByVal 0&)
        If Err.LastDllError Then
            Debug.Print LastSystemError
        Else
            With cProc.Image
                .Machine = imhProcess.Machine
                .MajorLinkerVersion = imhoProc.MajorLinkerVersion
                .MinorLinkerVersion = imhoProc.MinorLinkerVersion
            End With
        End If
    End If
    Dim imhoProcNT As IMAGE_OPTIONAL_HEADER_NT
    If imhProcess.SizeOfOptionalHeader - Len(imhoProc) >= Len(imhoProcNT) Then
        Call ReadFileLong(hfile, VarPtr(imhoProcNT), Len(imhoProcNT), BytesRead, ByVal 0&)
        If Err.LastDllError Then
            Debug.Print LastSystemError
        Else
            '\\ Save the current file pointer
            filepointer = FileApiUtils.GetCurrentFilePointer(hfile)
            cProc.Image.BaseAddress = imhoProcNT.ImageBase
            
            Call ProcessExportTable(imhoProcNT.DataDirectory(IMAGE_DIRECTORY_ENTRY_EXPORT), cProc.Image)
            Call ProcessImportTable(imhoProcNT.DataDirectory(IMAGE_DIRECTORY_ENTRY_IMPORT), cProc.Image)
            
            '\\ Restore the file pointer to what it was before we processed the import table
            Call FileApiUtils.ApiFileSeek(hfile, filepointer, FILE_BEGIN)
            Call ProcessSections(hfile, imhProcess.NumberOfSections, cProc.Image)
        End If
    End If
End If

mdifrmDebuggermain.RefreshModulesList

End Sub




Private Sub ProcessExportTable(ExportDirectory As IMAGE_DATA_DIRECTORY, Image As cPortableExecutableImage)

Dim deThis As IMAGE_EXPORT_DIRECTORY
Dim lBytesWritten As Long
Dim lpAddress As Long

Dim nFunction As Long
Dim fExport As cFunction

On Local Error GoTo MadOverflow

If ExportDirectory.VirtualAddress > 0 And ExportDirectory.Size > 0 Then
    '\\ Get the true address from the RVA
    lpAddress = Image.AbsoluteAddress(ExportDirectory.VirtualAddress)
    '\\ Copy the image_export_directory structure...
    Call ReadProcessMemoryLong(DebugProcess.Handle, lpAddress, VarPtr(deThis), Len(deThis), lBytesWritten)
    With deThis
        If .lpName <> 0 Then
            Image.Name = StringFromOutOfProcessPointer(DebugProcess.Handle, Image.AbsoluteAddress(.lpName), 32, False)
        End If
        If .NumberOfFunctions > 0 Then
            For nFunction = 1 To .NumberOfFunctions
                Set fExport = New cFunction
                lpAddress = LongFromOutOfprocessPointer(DebugProcess.Handle, Image.AbsoluteAddress(.lpAddressOfNames) + ((nFunction - 1) * 4))
                fExport.Name = StringFromOutOfProcessPointer(DebugProcess.Handle, Image.AbsoluteAddress(lpAddress), 64, False)
                fExport.Ordinal = .Base + IntegerFromOutOfprocessPointer(DebugProcess.Handle, Image.AbsoluteAddress(.lpAddressOfNameOrdinals) + ((nFunction - 1) * 2))
                fExport.ProcAddress = LongFromOutOfprocessPointer(DebugProcess.Handle, Image.AbsoluteAddress(.lpAddressOfFunctions) + ((nFunction - 1) * 4))
                If fExport.Name <> "" And fExport.ProcAddress <> 0 Then
                    Image.ExportedFunctions.Add fExport, fExport.Name
                End If
            Next nFunction
        End If
    End With
End If

MadOverflow:
    Exit Sub
    
End Sub

Private Sub ProcessImportTable(ImportDirectory As IMAGE_DATA_DIRECTORY, Image As cPortableExecutableImage)

Dim lpAddress As Long
Dim diThis As IMAGE_IMPORT_DESCRIPTOR
Dim byteswritten As Long
Dim sName As String
Dim lpNextName As Long
Dim lpNextThunk As Long

Dim lImportEntryIndex As Long

Dim mFunction As cFunction
Dim impDep As cImportDependency
Dim nOrdinal As Integer
Dim lpFuncAddress As Long

On Local Error GoTo MadOverlow

'\\ If the image has an imports section...
If ImportDirectory.VirtualAddress > 0 And ImportDirectory.Size > 0 Then
    '\\ Get the true address from the RVA
    lpAddress = Image.AbsoluteAddress(ImportDirectory.VirtualAddress)
    Call ReadProcessMemoryLong(DebugProcess.Handle, lpAddress, VarPtr(diThis), Len(diThis), byteswritten)
    
    While diThis.lpName <> 0
        '\\ Process this import directory entry
        sName = StringFromOutOfProcessPointer(DebugProcess.Handle, Image.AbsoluteAddress(diThis.lpName), 32, False)
        Set impDep = New cImportDependency
        impDep.Name = sName
        '\\ Process the import file's functions list
        If diThis.lpImportByName <> 0 Then
            lpNextName = LongFromOutOfprocessPointer(DebugProcess.Handle, Image.AbsoluteAddress(diThis.lpImportByName))
            lpNextThunk = LongFromOutOfprocessPointer(DebugProcess.Handle, Image.AbsoluteAddress(diThis.lpFirstThunk))
            While (lpNextName <> 0) And (lpNextThunk <> 0)
                '\\ get the function address
                lpFuncAddress = LongFromOutOfprocessPointer(DebugProcess.Handle, lpNextThunk)
                nOrdinal = IntegerFromOutOfprocessPointer(DebugProcess.Handle, lpNextName)
                '\\ Skip the two-byte ordinal hint
                lpNextName = lpNextName + 2
                '\\ Get this function's name
                sName = StringFromOutOfProcessPointer(DebugProcess.Handle, Image.AbsoluteAddress(lpNextName), 64, False)
                If sName <> "" Then
                    Set mFunction = New cFunction
                    With mFunction
                        .Name = sName
                        .Ordinal = nOrdinal
                        .ProcAddress = lpNextThunk
                    End With
                    impDep.Functions.Add mFunction, mFunction.Name
                    '\\ Get the next imported function...
                    lImportEntryIndex = lImportEntryIndex + 1
                    lpNextName = LongFromOutOfprocessPointer(DebugProcess.Handle, Image.AbsoluteAddress(diThis.lpImportByName + (lImportEntryIndex * 4)))
                    lpNextThunk = LongFromOutOfprocessPointer(DebugProcess.Handle, Image.AbsoluteAddress(diThis.lpFirstThunk + (lImportEntryIndex * 4)))
                Else
                    lpNextName = 0
                End If
            Wend
        End If
        
        Image.ImportDependencies.Add impDep, impDep.Name
        
        '\\ And get the next one
        lpAddress = lpAddress + Len(diThis)
        Call ReadProcessMemoryLong(DebugProcess.Handle, lpAddress, VarPtr(diThis), Len(diThis), byteswritten)
    Wend

End If

MadOverlow:
    '\\ Eep - how do I read a process memory if the address > MAXLONG??
    Exit Sub
    
End Sub
Private Sub ProcessSections(ByVal hfile As Long, ByVal SectionCount As Long, ByVal Image As cPortableExecutableImage)

'\\ Assumes that the hFile pointer is at the start of the "sections" bit
Dim headers() As IMAGE_SECTION_HEADER
Dim lBytesRead As Long
Dim hdIndex As Long
Dim nByte As Long
Dim oSect As cSection

ReDim headers(0 To SectionCount - 1) As IMAGE_SECTION_HEADER
Call ReadFileLong(hfile, VarPtr(headers(0)), SectionCount * Len(headers(0)), lBytesRead, ByVal 0&)
If Err.LastDllError Then
    Debug.Print LastSystemError
Else
    For hdIndex = 0 To SectionCount - 1
        Set oSect = New cSection
        With oSect
            For nByte = 0 To 7
                If headers(hdIndex).ImageName(nByte) > 0 Then
                    .SectionName = .SectionName & Chr$(headers(hdIndex).ImageName(nByte))
                End If
            Next nByte
            .Characteristics = headers(hdIndex).Characteristics
            .NumberOfLinenumbers = headers(hdIndex).NumberOfLinenumbers
            .NumberOfRelocations = headers(hdIndex).NumberOfRelocations
            .PhysicalAddress = headers(hdIndex).PhysicalAddress
            .PointerToLinenumbers = headers(hdIndex).PointerToLinenumbers
            .PointerToRawData = headers(hdIndex).PointerToRawData
            .PointerToRelocations = headers(hdIndex).PointerToRelocations
            .SizeOfRawData = headers(hdIndex).SizeOfRawData
            .VirtualAddress = headers(hdIndex).VirtualAddress
        End With
        Image.Sections.Add oSect, oSect.ICollectionItem_Key
        
    Next hdIndex
End If

End Sub


Private Function SkipDOSStub(ByVal hfile As Long) As Long

Dim BytesRead As Long

'\\ Go to start of file...
Call SetFilePointer(hfile, 0, 0, FILE_BEGIN)
If Err.LastDllError Then
    Debug.Print LastSystemError
End If

Dim stub As IMAGE_DOS_HEADER
Call ReadFileLong(hfile, VarPtr(stub), Len(stub), BytesRead, ByVal 0&)
If Err.LastDllError Then
    Debug.Print LastSystemError
Else
    Call SetFilePointer(hfile, stub.e_lfanew, 0, FILE_BEGIN)
End If

SkipDOSStub = stub.e_lfanew

End Function

Public Sub PrintASample(ByVal hfile As Long, ByVal address As Long, ByVal bytecount As Long)

Dim ssample As String, lBytesRead As Long, buff() As Byte, nIndex As Long

Call SetFilePointer(hfile, address, 0, FILE_BEGIN)
ReDim buff(bytecount) As Byte
Call ReadFileLong(hfile, VarPtr(buff(0)), bytecount, lBytesRead, ByVal 0&)
For nIndex = 0 To lBytesRead
    ssample = ssample & Chr$(buff(nIndex))
Next nIndex

Debug.Print ssample

End Sub

