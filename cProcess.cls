VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mProcessId As Long
Private mProcesshandle As Long

Private mBaseAddress As Long

Private mcolModules As Collection
Private mImage As cPortableExecutableImage

Public Name As String

Implements ICollectionItem

Public Property Get BaseAddress() As Long

BaseAddress = mBaseAddress

End Property

Friend Property Let BaseAddress(ByVal newval As Long)

mBaseAddress = newval

End Property



Public Property Get Image() As cPortableExecutableImage

If mImage Is Nothing Then
    Set mImage = New cPortableExecutableImage
End If
Set Image = mImage

End Property





Public Property Get Modules() As Collection

    Set Modules = mcolModules
    
End Property

Friend Property Let ProcessId(ByVal newval As Long)

mProcessId = newval

End Property

Public Property Get ProcessId() As Long

    ProcessId = mProcessId
    
End Property

Private Sub Class_Initialize()

Set mcolModules = New Collection

End Sub


Private Sub Class_Terminate()

Set mcolModules = Nothing
Set mImage = Nothing

End Sub

Public Property Get ICollectionItem_ClassName() As String

ICollectionItem_ClassName = "cProcess"

End Property


Public Property Get ICollectionItem_Key() As String

ICollectionItem_Key = "PROCID:" & ProcessId

End Property


Public Property Get Handle() As Long

    Handle = mProcesshandle

End Property

Friend Property Let Handle(ByVal hProc As Long)

mProcesshandle = hProc

End Property
