VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPortableExecutableImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMachine As ImageMachineTypes


Private mMajorLinkerVersion As Byte
Private mMinorLinkerVersion As Byte


Private mBaseAddress As Long

Private mColSections As Collection
Private mColImportDependencies As Collection
Private mcolExportFunctions As Collection

Private msName As String

Public Function AbsoluteAddress(ByVal RelativeAddress As Long) As Long

AbsoluteAddress = mBaseAddress + RelativeAddress

End Function

Friend Property Let BaseAddress(ByVal newval As Long)

mBaseAddress = newval

End Property


Public Property Get ExportedFunctions() As Collection

    Set ExportedFunctions = mcolExportFunctions

End Property

Public Property Get ImportDependencies() As Collection

    Set ImportDependencies = mColImportDependencies
    
End Property

Public Property Get LinkerVersion() As Single

If mMinorLinkerVersion > 0 Then
    LinkerVersion = mMajorLinkerVersion + (mMinorLinkerVersion / 100)
Else
    LinkerVersion = mMajorLinkerVersion
End If

End Property

Friend Property Let Machine(ByVal newval As ImageMachineTypes)

    mMachine = newval
    
End Property

Public Property Get Machine() As ImageMachineTypes

    Machine = mMachine

End Property
Friend Property Let MajorLinkerVersion(ByVal newval As Byte)

mMajorLinkerVersion = newval

End Property


Friend Property Let MinorLinkerVersion(ByVal newval As Byte)

mMinorLinkerVersion = newval

End Property


Friend Property Let Name(ByVal newname As String)

    msName = newname
    
End Property

Public Property Get Name() As String
    
    Name = msName
    
End Property

Public Property Get Sections() As Collection

    Set Sections = mColSections
    
End Property

Private Sub Class_Initialize()

Set mColSections = New Collection
Set mColImportDependencies = New Collection
Set mcolExportFunctions = New Collection

End Sub


Private Sub Class_Terminate()

Set mColSections = Nothing
Set mColImportDependencies = Nothing
Set mcolExportFunctions = Nothing

End Sub


