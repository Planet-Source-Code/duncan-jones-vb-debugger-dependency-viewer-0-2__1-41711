VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mBaseAddress As Long


Private mImage As cPortableExecutableImage

Implements ICollectionItem

Private mKey As String
Friend Property Let BaseAddress(ByVal newval As Long)

mBaseAddress = newval

End Property

Public Property Get BaseAddress() As Long

BaseAddress = mBaseAddress

End Property
Public Property Get Image() As cPortableExecutableImage

If mImage Is Nothing Then
    Set mImage = New cPortableExecutableImage
End If
Set Image = mImage

End Property


Public Property Get Name() As String

Name = mImage.Name

End Property

Private Sub Class_Terminate()

Set mImage = Nothing

End Sub


Public Property Get ICollectionItem_ClassName() As String

ICollectionItem_ClassName = "cModule"

End Property


Public Property Get ICollectionItem_Key() As String

If mKey = "" Then
    mKey = ICollectionItem_ClassName & MakeUniqueKey
End If

ICollectionItem_Key = mKey

End Property


