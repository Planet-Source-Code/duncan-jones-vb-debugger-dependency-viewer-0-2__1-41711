Attribute VB_Name = "DebugMain"
Option Explicit


Public Enum DebugEventTypes
    EXCEPTION_DEBUG_EVENT = 1&
    CREATE_THREAD_DEBUG_EVENT = 2&
    CREATE_PROCESS_DEBUG_EVENT = 3&
    EXIT_THREAD_DEBUG_EVENT = 4&
    EXIT_PROCESS_DEBUG_EVENT = 5&
    LOAD_DLL_DEBUG_EVENT = 6&
    UNLOAD_DLL_DEBUG_EVENT = 7&
    OUTPUT_DEBUG_STRING_EVENT = 8&
    RIP_EVENT = 9&
End Enum

Private Type DEBUG_EVENT_HEADER
    dwDebugEventCode As DebugEventTypes
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Enum ExceptionFlags
    EXCEPTION_CONTINUABLE = 0
    EXCEPTION_NONCONTINUABLE = 1   '\\ Noncontinuable exception
End Enum

Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15 '\\ maximum number of exception parameters

Public Enum ExceptionCodes
    EXCEPTION_GUARD_PAGE_VIOLATION = &H80000001
    EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
    EXCEPTION_BREAKPOINT = &H80000003
    EXCEPTION_SINGLE_STEP = &H80000004
    EXCEPTION_ACCESS_VIOLATION = &HC0000005
    EXCEPTION_IN_PAGE_ERROR = &HC0000006
    EXCEPTION_INVALID_HANDLE = &HC0000008
    EXCEPTION_NO_MEMORY = &HC0000017
    EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
    EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
    EXCEPTION_INVALID_DISPOSITION = &HC0000026
    EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
    EXCEPTION_FLOAT_DENORMAL_OPERAND = &HC000008D
    EXCEPTION_FLOAT_DIVIDE_BY_ZERO = &HC000008E
    EXCEPTION_FLOAT_INEXACT_RESULT = &HC000008F
    EXCEPTION_FLOAT_INVALID_OPERATION = &HC0000090
    EXCEPTION_FLOAT_OVERFLOW = &HC0000091
    EXCEPTION_FLOAT_STACK_CHECK = &HC0000092
    EXCEPTION_FLOAT_UNDERFLOW = &HC0000093
    EXCEPTION_INTEGER_DIVIDE_BY_ZERO = &HC0000094
    EXCEPTION_INTEGER_OVERFLOW = &HC0000095
    EXCEPTION_PRIVILEGED_INSTRUCTION = &HC0000096
    EXCEPTION_STACK_OVERFLOW = &HC00000FD
    EXCEPTION_CONTROL_C_EXIT = &HC000013A
End Enum

Private Type EXCEPTION_RECORD
    ExceptionCode                                        As ExceptionCodes
    ExceptionFlags                                       As ExceptionFlags
    pExceptionRecord                                     As Long
    ExceptionAddress                                     As Long
    NumberParameters                                     As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS)   As Long
End Type


Private Type DEBUG_EXCEPTION_DEBUG_INFO
    Header As DEBUG_EVENT_HEADER
    ExceptionRecord As EXCEPTION_RECORD
    dwFirstChance As Long
End Type

Private Type DEBUG_CREATE_THREAD_DEBUG_INFO
    Header As DEBUG_EVENT_HEADER
    hThread As Long
    lpThreadLocalBase As Long
    lpStartAddress As Long
End Type

Private Type DEBUG_CREATE_PROCESS_DEBUG_INFO
    Header As DEBUG_EVENT_HEADER
    hfile As Long
    hProcess As Long
    hThread As Long
    lpBaseOfImage As Long
    dwDebugInfoFileOffset As Long
    nDebugInfoSize As Long
    lpThreadLocalBase As Long
    lpStartAddress As Long
    lpImageName As Long
    fUnicode As Integer
End Type

Private Type DEBUG_EXIT_THREAD_DEBUG_INFO
    Header As DEBUG_EVENT_HEADER
    dwExitCode As Long
End Type

Private Type DEBUG_EXIT_PROCESS_DEBUG_INFO
    Header As DEBUG_EVENT_HEADER
    dwExitCode As Long
End Type

Private Type DEBUG_LOAD_DLL_DEBUG_INFO
    Header As DEBUG_EVENT_HEADER
    hfile As Long
    lpBaseOfDll As Long
    dwDebugInfoFileOffset As Long
    nDebugInfoSize As Long
    lpImageName As Long
    fUnicode As Integer
End Type

Private Type DEBUG_UNLOAD_DLL_DEBUG_INFO
    Header As DEBUG_EVENT_HEADER
    lpBaseOfDll As Long
End Type

Private Type DEBUG_OUTPUT_DEBUG_STRING_INFO
    Header As DEBUG_EVENT_HEADER
    lpDebugStringData As Long
    fUnicode As Integer
    nDebugStringLength As Integer
End Type

Private Type DEBUG_RIP_INFO
    Header As DEBUG_EVENT_HEADER
    dwError As Long
    dwType As Long
End Type

'\\ because debug_event is a union struct, need a biggest case buffer...
Private Type DEBUG_EVENT_BUFFER
    Header As DEBUG_EVENT_HEADER
    buffer(0 To 87) As Byte  '\\ Largest size of union is 100, size of header is 12...
End Type

'\\ Debug API calls
Private Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Function WaitForDebugEvent Lib "kernel32" (lpDebugEvent As DEBUG_EVENT_BUFFER, ByVal dwMilliseconds As Long) As Long
Private Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'\\local memory copying bumpf
Private Declare Sub CopyMemoryDebugEventHeader Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_EVENT_HEADER, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugExceptionInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_EXCEPTION_DEBUG_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugCreateThreadInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_CREATE_THREAD_DEBUG_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugCreateProcessInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_CREATE_PROCESS_DEBUG_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugExitThreadInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_EXIT_THREAD_DEBUG_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugExitProcessInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_EXIT_PROCESS_DEBUG_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugLoadDllInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_LOAD_DLL_DEBUG_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugUnloadDllInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_UNLOAD_DLL_DEBUG_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugOutputDebugstringInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_OUTPUT_DEBUG_STRING_INFO, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryDebugRIPInfo Lib "kernel32" Alias "RtlMoveMemory" (Destination As DEBUG_RIP_INFO, ByVal Source As Long, ByVal Length As Long)

Private Const INFINITE = &HFFFF      '  Infinite timeout
Private Const SHORTWAIT = 100

Public Enum DebugStates
    DBG_CONTINUE = &H10002
    DBG_TERMINATE_THREAD = &H40010003
    DBG_TERMINATE_PROCESS = &H40010004
    DBG_CONTROL_C = &H40010005
    DBG_CONTROL_BREAK = &H40010008
    DBG_EXCEPTION_NOT_HANDLED = &H80010001
End Enum

Public Enum Win32States
    STATUS_WAIT_0 = &H0
    STATUS_ABANDONED_WAIT_0 = &H80
    STATUS_USER_APC = &HC0
    STATUS_TIMEOUT = &H102
    STATUS_PENDING = &H103
    STATUS_SEGMENT_NOTIFICATION = &H40000005
    STATUS_GUARD_PAGE_VIOLATION = &H80000001
    STATUS_DATATYPE_MISALIGNMENT = &H80000002
    STATUS_BREAKPOINT = &H80000003
    STATUS_SINGLE_STEP = &H80000004
    STATUS_ACCESS_VIOLATION = &HC0000005
    STATUS_IN_PAGE_ERROR = &HC0000006
    STATUS_INVALID_HANDLE = &HC0000008
    STATUS_NO_MEMORY = &HC0000017
    STATUS_ILLEGAL_INSTRUCTION = &HC000001D
    STATUS_NONCONTINUABLE_EXCEPTION = &HC0000025
    STATUS_INVALID_DISPOSITION = &HC0000026
    STATUS_ARRAY_BOUNDS_EXCEEDED = &HC000008C
    STATUS_FLOAT_DENORMAL_OPERAND = &HC000008D
    STATUS_FLOAT_DIVIDE_BY_ZERO = &HC000008E
    STATUS_FLOAT_INEXACT_RESULT = &HC000008F
    STATUS_FLOAT_INVALID_OPERATION = &HC0000090
    STATUS_FLOAT_OVERFLOW = &HC0000091
    STATUS_FLOAT_STACK_CHECK = &HC0000092
    STATUS_FLOAT_UNDERFLOW = &HC0000093
    STATUS_INTEGER_DIVIDE_BY_ZERO = &HC0000094
    STATUS_INTEGER_OVERFLOW = &HC0000095
    STATUS_PRIVILEGED_INSTRUCTION = &HC0000096
    STATUS_STACK_OVERFLOW = &HC00000FD
    STATUS_CONTROL_C_EXIT = &HC000013A
End Enum

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long


    Const TH32CS_SNAPHEAPLIST = &H1
    Const TH32CS_SNAPPROCESS = &H2
    Const TH32CS_SNAPTHREAD = &H4
    Const TH32CS_SNAPMODULE = &H8
    Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
    Const TH32CS_INHERIT = &H80000000
    Const MAX_PATH As Integer = 260
    Private Type PROCESSENTRY32
        dwSize As Long
        cntUsage As Long
        th32ProcessID As Long
        th32DefaultHeapID As Long
        th32ModuleID As Long
        cntThreads As Long
        th32ParentProcessID As Long
        pcPriClassBase As Long
        dwFlags As Long
        szExeFile As String * MAX_PATH
    End Type

    Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
    Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
    Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

    Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
    Private Declare Function GetModuleBaseName Lib "psapi.dll" _
              Alias "GetModuleBaseNameA" _
                        (ByVal hProcess As Long, _
                         ByVal hModule As Long, _
                         ByVal lpBaseName As String, _
                         ByVal nSize As Long) As Long

    Private Declare Function EnumProcessModules Lib "psapi.dll" _
                        (ByVal hProcess As Long, _
                         lphModule As Long, _
                         ByVal cb As Long, _
                         lpcbNeeded As Long) As Long



'\ API Error decoding
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Enum ProcessAccessPriviledges
     PROCESS_TERMINATE = &H1
     PROCESS_CREATE_THREAD = &H2
     PROCESS_SET_SESSIONID = &H4
     PROCESS_VM_OPERATION = &H8
     PROCESS_VM_READ = &H10
     PROCESS_VM_WRITE = &H20
     PROCESS_DUP_HANDLE = &H40
     PROCESS_CREATE_PROCESS = &H80
     PROCESS_SET_QUOTA = &H100
     PROCESS_SET_INFORMATION = &H200
     PROCESS_QUERY_INFORMATION = &H400
     PROCESS_SYNCHRONISE = &H100000
     PROCESS_ALL_ACCESS = &H100FFF
End Enum
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As ProcessAccessPriviledges, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

'\\ Local memory manipulation routines
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long

'\\ Other process memory manipulation...
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemoryBytes Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Byte, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long


Private Type LDT_ENTRY
  LimitLow As Integer
  BaseLow As Integer
  BaseMid As Byte
  Flags1 As Byte
  Flasg2 As Byte
  BaseHi As Byte
End Type
Private Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long

Private Enum ThreadSelectorSegmentTypes
    Read_Only_Data = 0
    Read_Write_Data = 1
    Unused = 2
    Read_Write_Expand_Down = 3
    Execute_Only = 4
    Executeable_Readable_Code = 5
    Execute_Only_Conforming = 6
    Execute_Only_Readable_Conforming = 7
End Enum


'\\ Bits and byte manipulations .......
Public Enum BYTE_BITMASKS
   Bit_0 = &H1
   Bit_1 = &H2
   Bit_2 = &H4
   Bit_3 = &H8
   Bit_4 = &H10
   Bit_5 = &H20
   Bit_6 = &H40
   Bit_7 = &H80
End Enum
Private Declare Sub CopyMemoryByte Lib "kernel32" Alias "RtlMoveMemory" (Destination As Byte, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryWord Lib "kernel32" Alias "RtlMoveMemory" (Destination As Integer, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryFromByte Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Byte, ByVal Length As Long)
Private Declare Sub CopyMemoryFromWord Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Integer, ByVal Length As Long)

'\\ ShellExecuteEx stuff...
Public Enum ShellExecuteExFlags
    SEE_MASK_CLASSNAME = &H1
    SEE_MASK_CLASSKEY = &H3
    SEE_MASK_IDLIST = &H4
    SEE_MASK_INVOKEIDLIST = &HC
    SEE_MASK_ICON = &H10
    SEE_MASK_HOTKEY = &H20
    SEE_MASK_NOCLOSEPROCESS = &H40
    SEE_MASK_CONNECTNETDRV = &H80
    SEE_MASK_FLAG_DDEWAIT = &H100
    SEE_MASK_DOENVSUBST = &H200
    SEE_MASK_FLAG_NO_UI = &H400
    SEE_MASK_UNICODE = &H4000
    SEE_MASK_NO_CONSOLE = &H8000
    SEE_MASK_ASYNCOK = &H100000
    SEE_MASK_HMONITOR = &H200000
End Enum

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As ShellExecuteExFlags
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    '  Optional fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Declare Function ShellExecuteEx Lib "shell32.dll" (sei As SHELLEXECUTEINFO) As Long

'\\ Createprocess stuff
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Enum ProcessCreationFlags
    DEBUG_PROCESS = &H1
    DEBUG_ONLY_THIS_PROCESS = &H2
    CREATE_SUSPENDED = &H4
    DETACHED_PROCESS = &H8
    CREATE_NEW_CONSOLE = &H10
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
    REALTIME_PRIORITY_CLASS = &H100
    CREATE_NEW_PROCESS_GROUP = &H200
    CREATE_UNICODE_ENVIRONMENT = &H400
    CREATE_SEPARATE_WOW_VDM = &H800
    CREATE_SHARED_WOW_VDM = &H1000
    CREATE_FORCEDOS = &H2000
    CREATE_DEFAULT_ERROR_MODE = &H4000000
    CREATE_NO_WINDOW = &H8000000
End Enum

Private Const STARTF_USESHOWWINDOW = &H1

Private Declare Function CreateProcess Lib _
    "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
                                       ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
                                       ByVal dwCreationFlags As ProcessCreationFlags, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long


'\\ -- Public state variables....
Public bContinueOK As Boolean
Public bWaitDebugee As Boolean

Public DebugProcess As cProcess

Public AppSettings As cDebugAppSettings


Private Sub DLLLoaded(DebugLoadDllInfo As DEBUG_LOAD_DLL_DEBUG_INFO)

Dim cmod As cModule
Dim sImagename As String

With DebugLoadDllInfo
               
     Set cmod = New cModule
     
    If .hfile <> 0 Then
        Call PE_Utilities.FillDLLModuleInforFromFile(.hfile, cmod)
    End If
       
    '\\ Get the debug info items....
    If .dwDebugInfoFileOffset > 0 And .nDebugInfoSize > 0 Then
    
    End If
    '\\ Add this to the process module collection
    DebugProcess.Modules.Add cmod, cmod.ICollectionItem_Key
End With

End Sub

'\ --[LoByte]-----------------------------------------------------------------------------
'\ Returns the low byte component of an integer value
'\ Parameters:
'\ w - The integer of which we need the loWord
'\
'\ ----------------------------------------------------------------------------------------
Public Function LoByte(w As Integer) As Byte

Call CopyMemoryByte(LoByte, VarPtr(w), 1)

End Function

Public Sub LogEvent(ByVal EventCode As DebugEventTypes, ByVal ExtraData As String)

Dim sEventName As String

Select Case EventCode
Case EXCEPTION_DEBUG_EVENT
    sEventName = "Debug Event"
Case CREATE_THREAD_DEBUG_EVENT
    sEventName = "Thread Created"
Case CREATE_PROCESS_DEBUG_EVENT
    sEventName = "Process Created"
Case EXIT_THREAD_DEBUG_EVENT
    sEventName = "Thread Exit"
Case EXIT_PROCESS_DEBUG_EVENT
    sEventName = "Process Exit"
Case LOAD_DLL_DEBUG_EVENT
    sEventName = "Load DLL"
Case UNLOAD_DLL_DEBUG_EVENT
    sEventName = "Unload DLL"
Case OUTPUT_DEBUG_STRING_EVENT
    sEventName = "Debug Message"
Case RIP_EVENT
    sEventName = "RIP Event"
End Select

mdifrmDebuggermain.LogEvent sEventName, ExtraData

If AppSettings.Logfilehandle <> 0 Then
    Print #AppSettings.Logfilehandle, sEventName
    Print #AppSettings.Logfilehandle, ExtraData
End If

End Sub

Public Function LongFromOutOfprocessPointer(ByVal hProcess As Long, ByVal lpAddress As Long) As Long

Dim lRet As Long
Dim lBytesWritten As Long

Call ReadProcessMemory(hProcess, lpAddress, ByVal VarPtr(lRet), Len(lRet), lBytesWritten)
If lBytesWritten > 0 Then
    LongFromOutOfprocessPointer = lRet
End If

End Function

Public Function IntegerFromOutOfprocessPointer(ByVal hProcess As Long, ByVal lpAddress As Long) As Integer

Dim lRet As Integer
Dim lBytesWritten As Long

Call ReadProcessMemory(hProcess, lpAddress, ByVal VarPtr(lRet), Len(lRet), lBytesWritten)
If lBytesWritten > 0 Then
    IntegerFromOutOfprocessPointer = lRet
End If

End Function
'\ --[LoWord]-----------------------------------------------------------------------------
'\ Returns the low word component of a long value
'\ Parameters:
'\ dw - The long of which we need the LoWord
'\
'\ ----------------------------------------------------------------------------------------
Public Function LoWord(dw As Long) As Integer

Call CopyMemoryWord(LoWord, VarPtr(dw), 2)

End Function


'\ --[HiByte]-----------------------------------------------------------------------------
'\ Returns the high byte component of an integer
'\ Parameters:
'\ w - The integer of which we need the HiByte
'\
'\ ----------------------------------------------------------------------------------------
Public Function HiByte(ByVal w As Integer) As Byte

Call CopyMemoryByte(HiByte, VarPtr(w) + 1, 1)

End Function

'\ --[HiWord]-----------------------------------------------------------------------------
'\ Returns the high word component of a long value
'\ Parameters:
'\ dw - The long of which we need the HiWord
'\
'\ ----------------------------------------------------------------------------------------
Public Function HiWord(dw As Long) As Integer

Call CopyMemoryWord(HiWord, VarPtr(dw) + 2, 2)

End Function

Private Function GetSelectorAddress(SelectorEntry As LDT_ENTRY) As Long

'\\ Combine BaseLow + BaseMid + BaseHigh to make 32 bit address
With SelectorEntry
    GetSelectorAddress = MakeLong(.BaseLow, MakeWord(.BaseMid, .BaseHi))
End With

End Function

Private Function GetSelectorType(SelectorEntry As LDT_ENTRY) As ThreadSelectorSegmentTypes

'\\ This is held in bits 0..4 of selector entry flags1
GetSelectorType = SelectorEntry.Flags1 And (Bit_0 Or Bit_1 Or Bit_2 Or Bit_3 Or Bit_4)

End Function

Public Function MakeUniqueKey() As String

Static nextId As Long

nextId = nextId + 1
MakeUniqueKey = Format$(nextId, "0000000000")

End Function

Public Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer

Dim wOut As Integer
Call CopyMemoryFromByte(VarPtr(wOut), LoByte, 1)
Call CopyMemoryFromByte(VarPtr(wOut) + 1, HiByte, 1)
MakeWord = wOut

End Function


Public Function MakeLong(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long

Dim lOut As Long
Call CopyMemoryFromWord(VarPtr(lOut), LoWord, 2)
Call CopyMemoryFromWord(VarPtr(lOut) + 2, HiWord, 2)
MakeLong = lOut

End Function
Private Sub ProcessAttached(DebugCreateProcessInfo As DEBUG_CREATE_PROCESS_DEBUG_INFO)

With DebugCreateProcessInfo
    If .hfile <> 0 Then
        '\\ Read the "portable executable format" details from this handle...
        Call PE_Utilities.FillProcessInfoFromFile(.hfile, DebugProcess)
    End If
    If .dwDebugInfoFileOffset > 0 And .nDebugInfoSize > 0 Then
        '\\ Read the debug info..
        
    End If
End With

End Sub

Public Function StringFromOutOfProcessPointer(ByVal hProcess As Long, ByVal lpString As Long, ByVal Length As Long, ByVal Unicode As Boolean) As String

Dim buf() As Byte
Dim lRet As Long
Dim lBytesWritten As Long
Dim sTemp As String

ReDim buf(Length) As Byte

lRet = ReadProcessMemoryBytes(hProcess, lpString, buf(0), Length, lBytesWritten)
If lBytesWritten = 0 Then
    While lBytesWritten = 0 And Length > 0
        Length = Length - 1
        lRet = ReadProcessMemoryBytes(hProcess, lpString, buf(0), Length, lBytesWritten)
    Wend
End If
If lRet <> 0 Then
    If Unicode Then
        StringFromOutOfProcessPointer = StrConv(buf, vbFromUnicode)
    Else
        For lRet = 0 To lBytesWritten
            If buf(lRet) = 0 Then
                Exit For
            End If
            sTemp = sTemp & Chr$(buf(lRet))
        Next lRet
        StringFromOutOfProcessPointer = sTemp
    End If
Else
    If Err.LastDllError Then
        Debug.Print LastSystemError
    End If
End If

End Function

Public Function StringFromPointer(lpString As Long, lMaxLength As Long) As String

  Dim sRet As String
  Dim lRet As Long

  If lpString = 0 Then
    StringFromPointer = ""
    Exit Function
  End If

  If IsBadStringPtrByLong(lpString, lMaxLength) Then
    '\\ An error has occured - do not attempt to use this pointer
      StringFromPointer = ""
    Exit Function
  End If

  '\\ Pre-initialise the return string...
  sRet = Space$(lMaxLength)
  CopyMemory ByVal sRet, ByVal lpString, ByVal Len(sRet)
  If Err.LastDllError = 0 Then
    If InStr(sRet, Chr$(0)) > 0 Then
      sRet = Left$(sRet, InStr(sRet, Chr$(0)) - 1)
    End If
  End If

  StringFromPointer = sRet

End Function


 

'\ -- [ LastSystemError ]----------------------------------
'\ Returns the message from the system which describes the
'\ last dll error to occur, as
'\ held in Err.LastDllError. This function should be
'\ called as soon after the API call
'\ which might have errored, as this member can be reset
'\ to zero by subsequent API calls.
'\ --------------------------------------------------------
Public Function LastSystemError() As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Dim sError As String * 500 '\ Preinitilise a string buffer to put any error message into
Dim lErrNum As Long
Dim lErrMsg As Long

lErrNum = Err.LastDllError

lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrNum, 0, sError, Len(sError), 0)

LastSystemError = Trim(sError)

End Function

Public Sub DebugLoop(ByVal ProcessId As Long)

Dim deBuffer As DEBUG_EVENT_BUFFER

Dim DebugExceptionInfo As DEBUG_EXCEPTION_DEBUG_INFO
Dim DebugCreateThreadInfo As DEBUG_CREATE_THREAD_DEBUG_INFO
Dim DebugCreateProcessInfo As DEBUG_CREATE_PROCESS_DEBUG_INFO
Dim DebugExitThreadInfo As DEBUG_EXIT_THREAD_DEBUG_INFO
Dim DebugExitProcessInfo As DEBUG_EXIT_PROCESS_DEBUG_INFO
Dim DebugLoadDllInfo As DEBUG_LOAD_DLL_DEBUG_INFO
Dim DebugUnloadDllInfo As DEBUG_UNLOAD_DLL_DEBUG_INFO
Dim DebugOutputDebugstringInfo As DEBUG_OUTPUT_DEBUG_STRING_INFO
Dim DebugRIPInfo As DEBUG_RIP_INFO
Dim hProc As Long


Dim lRet As Long

bContinueOK = True

'\\ Whatever happens you must not debug yourself!!!!
If ProcessId = GetCurrentProcess() Then
    Exit Sub
Else
    DebugProcess.ProcessId = ProcessId
End If

If ProcessId = 0 Then
    Dim PINFO As PROCESS_INFORMATION
    Dim si As STARTUPINFO
    
    With si
        .cb = Len(si)
    End With
    
    If CreateProcess(AppSettings.DebugeeAppName, vbNullString, 0&, 0&, 0&, DEBUG_PROCESS + DEBUG_ONLY_THIS_PROCESS + NORMAL_PRIORITY_CLASS, 0&, vbNullString, si, PINFO) Then
        ProcessId = PINFO.dwProcessId
        hProc = PINFO.hProcess
    Else
        LogEvent deBuffer.Header.dwDebugEventCode, "CreateProcess failed with error: " & LastSystemError()
        Exit Sub
    End If
Else
    If hProc = 0 Then
        hProc = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessId)
    End If
    If hProc = 0 Then
        LogEvent deBuffer.Header.dwDebugEventCode, "OpenProcess failed with error: " & LastSystemError()
        Exit Sub
    End If
    
    lRet = DebugActiveProcess(ProcessId)
    If lRet = 0 Then
        LogEvent deBuffer.Header.dwDebugEventCode, "DebugActiveprocess failed with error: " & LastSystemError()
        Exit Sub
    End If
End If
DebugProcess.Handle = hProc
'\\ Set the process being debugged's name
DebugProcess.Name = AppSettings.DebugeeAppName
While InStr(DebugProcess.Name, "\")
    DebugProcess.Name = Mid$(DebugProcess.Name, InStr(DebugProcess.Name, "\") + 1)
Wend

While bContinueOK
    If WaitForDebugEvent(deBuffer, SHORTWAIT) Then
        '\\ Pause the debugee to allow it to be investigated...
        bWaitDebugee = True
        mdifrmDebuggermain.mnuContinue.Enabled = True
        
        Select Case deBuffer.Header.dwDebugEventCode
        Case EXCEPTION_DEBUG_EVENT
        '\\ Process the exception code. When handling
        '\\ exceptions, remember to set the continuation
        '\\ status parameter (dwContinueStatus). This value
        '\\ is used by the ContinueDebugEvent function.
            Call CopyMemoryDebugExceptionInfo(DebugExceptionInfo, VarPtr(deBuffer), Len(DebugExceptionInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, IIf(DebugExceptionInfo.dwFirstChance, "First pass", "Final pass")
            Select Case DebugExceptionInfo.ExceptionRecord.ExceptionCode
                Case STATUS_ACCESS_VIOLATION
                '\\ First chance: Pass this on to the system.
                '\\ Last chance: Display an appropriate error.
                LogEvent deBuffer.Header.dwDebugEventCode, "Access violation"
 
                Case STATUS_BREAKPOINT
                '\\ First chance: Display the current
                '\\ instruction and register values.
                LogEvent deBuffer.Header.dwDebugEventCode, "Breakpoint"
 
                Case STATUS_DATATYPE_MISALIGNMENT
                '\\ First chance: Pass this on to the system.
                '\\ Last chance: Display an appropriate error.
                LogEvent deBuffer.Header.dwDebugEventCode, "DataType Misalignment"
 
                Case STATUS_SINGLE_STEP
                '\\ First chance: Update the display of the
                '\\ current instruction and register values.
                LogEvent deBuffer.Header.dwDebugEventCode, "Single step"
 
                Case DBG_CONTROL_C
                '\\ First chance: Pass this on to the system.
                '\\ Last chance: Display an appropriate error.
                LogEvent deBuffer.Header.dwDebugEventCode, "Ctrl+C"
 
                Case Else
                '\\ Handle other exceptions.
                LogEvent deBuffer.Header.dwDebugEventCode, "Exception : " & DebugExceptionInfo.ExceptionRecord.ExceptionCode
                    
            End Select
            If Not AppSettings.BreakOnException Then
                bWaitDebugee = False
            End If
 
        Case CREATE_THREAD_DEBUG_EVENT
            '\\ As needed, examine or change the thread's registers
            '\\ with the GetThreadContext and SetThreadContext functions;
            '\\ and suspend and resume thread execution with the
            '\\ SuspendThread and ResumeThread functions.
            Call CopyMemoryDebugCreateThreadInfo(DebugCreateThreadInfo, VarPtr(deBuffer), Len(DebugCreateThreadInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, "Base address: " & DebugCreateThreadInfo.lpStartAddress
            If Not AppSettings.BreakOnCreateThread Then
                bWaitDebugee = False
            End If
            
        Case CREATE_PROCESS_DEBUG_EVENT
            '\\ As needed, examine or change the registers of the
            '\\ process's initial thread with the GetThreadContext and
            '\\ SetThreadContext functions; read from and write to the
            '\\ process's virtual memory with the ReadProcessMemory and
            '\\ WriteProcessMemory functions; and suspend and resume
            '\\ thread execution with the SuspendThread and ResumeThread
            '\\ functions. Be sure to close the handle to the process image
            '\\ file with CloseHandle.
            Call CopyMemoryDebugCreateProcessInfo(DebugCreateProcessInfo, VarPtr(deBuffer), Len(DebugCreateProcessInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, "Process handle : " & DebugCreateProcessInfo.hProcess
            Call ProcessAttached(DebugCreateProcessInfo)
            
            Call CloseHandle(DebugCreateProcessInfo.hfile)
            If Not AppSettings.BreakOnCreateprocess Then
                bWaitDebugee = False
            End If
        
        Case EXIT_THREAD_DEBUG_EVENT
            '\\ Display the thread's exit code.
            Call CopyMemoryDebugExitThreadInfo(DebugExitThreadInfo, VarPtr(deBuffer), Len(DebugExitThreadInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, "Exit code: " & DebugExitThreadInfo.dwExitCode
            If Not AppSettings.BreakOnCreateThread Then
                bWaitDebugee = False
            End If
            
        Case EXIT_PROCESS_DEBUG_EVENT
            '\\ Display the process's exit code.
            Call CopyMemoryDebugExitProcessInfo(DebugExitProcessInfo, VarPtr(deBuffer), Len(DebugExitProcessInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, "Exit code: " & DebugExitProcessInfo.dwExitCode
            If Not AppSettings.BreakOnExitProcess Then
                bWaitDebugee = False
            End If
            mdifrmDebuggermain.Debugging = False
            bContinueOK = False
            
        Case LOAD_DLL_DEBUG_EVENT
            '\\ Read the debugging information included in the newly
            '\\ loaded DLL. Be sure to close the handle to the loaded DLL
            '\\ with CloseHandle.
            Call CopyMemoryDebugLoadDllInfo(DebugLoadDllInfo, VarPtr(deBuffer), Len(DebugLoadDllInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, IIf(DebugLoadDllInfo.nDebugInfoSize, "Debug info present", "No debug info")
           
            '\\ Extra processing for details of this events....
            Call DLLLoaded(DebugLoadDllInfo)
            
            Call CloseHandle(DebugLoadDllInfo.hfile)
            If Not AppSettings.BreakOnLoadDll Then
                bWaitDebugee = False
            End If
            
        Case UNLOAD_DLL_DEBUG_EVENT
            '\\ Display a message that the DLL has been unloaded.
            Call CopyMemoryDebugUnloadDllInfo(DebugUnloadDllInfo, VarPtr(deBuffer), Len(DebugUnloadDllInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, "DLL Base: " & DebugUnloadDllInfo.lpBaseOfDll
            If Not AppSettings.BreakOnUnloadDll Then
                bWaitDebugee = False
            End If
            
        Case OUTPUT_DEBUG_STRING_EVENT
            '\\ Display the output debugging string.
            Call CopyMemoryDebugOutputDebugstringInfo(DebugOutputDebugstringInfo, VarPtr(deBuffer), Len(DebugOutputDebugstringInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, StringFromOutOfProcessPointer(hProc, DebugOutputDebugstringInfo.lpDebugStringData, DebugOutputDebugstringInfo.nDebugStringLength, DebugOutputDebugstringInfo.fUnicode)
            If Not AppSettings.BreakOnDebugString Then
                bWaitDebugee = False
            End If
            
        Case RIP_EVENT
            '\\ Get the process about why the RIP occured
            Call CopyMemoryDebugRIPInfo(DebugRIPInfo, VarPtr(deBuffer), Len(DebugRIPInfo))
            LogEvent deBuffer.Header.dwDebugEventCode, "Error code: " & DebugRIPInfo.dwError
            bContinueOK = False
            If Not AppSettings.BreakOnRipEvent Then
                bWaitDebugee = False
            End If
        End Select
 
        While bWaitDebugee
            DoEvents
        Wend
        mdifrmDebuggermain.mnuContinue.Enabled = False
        Call ContinueDebugEvent(deBuffer.Header.dwProcessId, deBuffer.Header.dwThreadId, DBG_CONTINUE)
        
    Else
        '\\ No debug event occured...allow interface to clean itself up and try again
        DoEvents
    End If
Wend

If hProc <> 0 Then
    CloseHandle hProc
End If


End Sub

Public Sub FillProcessList(ByVal lstIn As ListBox)

Dim hSnapShot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Boolean
Dim strProcName As String

On Error GoTo OldWindowsversion

'\\Takes a snapshot of the processes and the heaps, modules, and threads used by the processes
hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
'\\set the length of our ProcessEntry-type
uProcess.dwSize = Len(uProcess)
'\\Retrieve information about the first process encountered in our system snapshot
r = Process32First(hSnapShot, uProcess)
Do While r
    strProcName = UCase(Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0)))
    If Trim$(strProcName) <> "" Then
        lstIn.AddItem strProcName
        lstIn.ItemData(lstIn.NewIndex) = uProcess.th32ProcessID
    End If
    '\\Retrieve information about the next process recorded in our system snapshot
    r = Process32Next(hSnapShot, uProcess)
Loop
'close our snapshot handle
CloseHandle hSnapShot

Exit Sub

OldWindowsversion:
    Dim proclst() As Long
    Dim lBytesRequired As Long, lItem As Long
    ReDim proclst(256) As Long
    Dim sName As String
    
    r = EnumProcesses(proclst(0), UBound(proclst) * 4, lBytesRequired)
    If lBytesRequired > (UBound(proclst) * 4) Then
        ReDim proclst(lBytesRequired / 4) As Long
        r = EnumProcesses(proclst(0), UBound(proclst) * 4, lBytesRequired)
    End If
    For lItem = 0 To UBound(proclst)
        If proclst(lItem) <> 0 Then
            sName = GetProcessNameFromId(proclst(lItem))
            If sName <> "" And InStr(sName, "?") = 0 Then
                lstIn.AddItem sName
                lstIn.ItemData(lstIn.NewIndex) = proclst(lItem)
            End If
        End If
    Next lItem

End Sub

Private Function GetProcessNameFromId(ByVal ProcId As Long) As String

Dim lngModuleHandle As Long
Dim lngReturnValue As Long
Dim strProcessName As String
Dim hProc As Long
Dim lngNumberOfBytesReceived As Long

hProc = OpenProcess(PROCESS_ALL_ACCESS, False, ProcId)

        lngModuleHandle = 0
        lngReturnValue = EnumProcessModules(hProc, lngModuleHandle, 4&, lngNumberOfBytesReceived)
        If Err.LastDllError Then
            Debug.Print LastSystemError
        End If
        
        ' Get the name of the module
        strProcessName = String$(256, 0)
        lngReturnValue = GetModuleBaseName(hProc, lngModuleHandle, strProcessName, Len(strProcessName))
        If Err.LastDllError Then
            Debug.Print LastSystemError
        End If
        
CloseHandle hProc

GetProcessNameFromId = Trim$(strProcessName)

End Function


Public Sub Main()

Set AppSettings = New cDebugAppSettings
Set DebugProcess = New cProcess

'\\ Show the debugger interface
mdifrmDebuggermain.Show
While Not mdifrmDebuggermain.Terminated
    DoEvents
Wend
Unload mdifrmDebuggermain
Set mdifrmDebuggermain = Nothing

Set DebugProcess = Nothing
Set AppSettings = Nothing

End Sub

