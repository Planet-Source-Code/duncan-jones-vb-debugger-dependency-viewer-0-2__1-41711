VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICollectionItem
Dim sKey As String

Public SymbolName As String
Public Ordinal As Long
Public FunctionRelativeVirtualAddress As Long
Public Property Get ICollectionItem_ClassName() As String

ICollectionItem_ClassName = "cSymbol"

End Property

Public Property Get ICollectionItem_Key() As String

If sKey = "" Then
    sKey = ICollectionItem_ClassName & DebugMain.MakeUniqueKey()
End If

ICollectionItem_Key = sKey

End Property
