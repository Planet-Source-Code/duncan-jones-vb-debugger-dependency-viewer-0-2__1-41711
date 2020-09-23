VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImportDependency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mName As String

Private mcolFunctions As Collection

Friend Property Let Name(ByVal newname As String)

mName = newname

End Property

Public Property Get Name() As String

    Name = mName
    
End Property

Private Sub Class_Initialize()

Set mcolFunctions = New Collection

End Sub

Private Sub Class_Terminate()

Set mcolFunctions = Nothing

End Sub

Public Property Get Functions() As Collection

    Set Functions = mcolFunctions
    
End Property
