VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDebugAppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'\\ Class that defines the current session setting

Private msDebugeeAppName As String
Private msLogFile As String
Private mhFile As Integer

Public BreakOnException As Boolean     'EXCEPTION_DEBUG_EVENT
Public BreakOnCreateThread As Boolean  'CREATE_THREAD_DEBUG_EVENT
Public BreakOnCreateprocess As Boolean 'CREATE_PROCESS_DEBUG_EVENT     EXIT_THREAD_DEBUG_EVENT = 4&
Public BreakOnExitProcess As Boolean   'EXIT_PROCESS_DEBUG_EVENT
Public BreakOnLoadDll As Boolean       'LOAD_DLL_DEBUG_EVENT
Public BreakOnUnloadDll As Boolean     'UNLOAD_DLL_DEBUG_EVENT
Public BreakOnDebugString As Boolean   'OUTPUT_DEBUG_STRING_EVENT
Public BreakOnRipEvent As Boolean      'RIP_EVENT

Public ProcessId As Long
Public RunNew As Boolean

Public Event DebugeeNameChanged()

Public Property Get DebugeeAppName() As String

    DebugeeAppName = msDebugeeAppName
    
End Property

Public Property Let DebugeeAppName(ByVal newname As String)

If msDebugeeAppName <> newname Then
    msDebugeeAppName = newname
    RaiseEvent DebugeeNameChanged
End If

End Property

Public Property Get LogFile() As String

    LogFile = msLogFile

End Property

Public Property Let LogFile(ByVal newname As String)

If newname <> msLogFile Then
    If msLogFile <> "" Then
        Close #mhFile
    End If
    msLogFile = newname
    If msLogFile <> "" Then
       mhFile = FreeFile
       Open msLogFile For Output As #mhFile
    End If
End If

End Property


Public Property Get Logfilehandle() As Integer
    
    Logfilehandle = mhFile
    
End Property

'\\ --[SetupFromCommandLine]--------------------------------------
'\\ Sets the app settings from the command line....
'\\  -d"Filename" - debugee filename
'\\  -o"Filename" - log file filename
'\\  -bxx Break on exception
'\\  -bct Break on create thread
'\\  -bcp Break on create process
'\\  -bxp Break on exit process
'\\  -bld Break on load dll
'\\  -bud Break on unload dll
'\\  -bds Break on debug string
'\\  -brip Break on RIP event
'\\ e.g. -d"C:\Program Files\DevStudio\VB\Debugger\Tester\DebugTest.exe" -bds
'\\ --------------------------------------------------------------
Public Sub SetupFromCommandLine(ByVal CommandLine As String)

Dim nLHS As Long

If InStr(CommandLine, "-d") Then
    nLHS = InStr(CommandLine, "-d") + 3
    If InStr(Mid$(CommandLine, nLHS), Chr$(34)) Then
        DebugeeAppName = Mid$(CommandLine, nLHS, InStr(Mid$(CommandLine, nLHS), Chr$(34)) - 1)
    End If
End If
    
If InStr(CommandLine, "-o") Then
    nLHS = InStr(CommandLine, "-o") + 3
        If InStr(Mid$(CommandLine, nLHS), Chr$(34)) Then
        LogFile = Mid$(CommandLine, nLHS, InStr(Mid$(CommandLine, nLHS), Chr$(34)) - 1)
    End If
End If

BreakOnException = InStr(CommandLine, "-bxc")
BreakOnCreateThread = InStr(CommandLine, "-bct")
BreakOnCreateprocess = InStr(CommandLine, "-bcp")
BreakOnExitProcess = InStr(CommandLine, "-bxp")
BreakOnLoadDll = InStr(CommandLine, "-bld")
BreakOnUnloadDll = InStr(CommandLine, "-bud")
BreakOnDebugString = InStr(CommandLine, "-bds")
BreakOnRipEvent = InStr(CommandLine, "-brip")

End Sub

Private Sub Class_Terminate()

If mhFile <> 0 Then
    Close #mhFile
End If

End Sub


