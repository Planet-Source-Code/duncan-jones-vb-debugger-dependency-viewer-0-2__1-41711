VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mName As String
Private mnOrdinal As Integer
Private lpProcAddress As Long

Friend Property Let Name(ByVal newname As String)

mName = newname

End Property

Public Property Get Name() As String

    Name = mName
    
End Property

Friend Property Let Ordinal(ByVal newval As Integer)
    
    mnOrdinal = newval
    
End Property
Public Property Get Ordinal() As Integer

    Ordinal = mnOrdinal
    
End Property

Friend Property Let ProcAddress(ByVal newval As Long)
    lpProcAddress = newval
End Property

Public Property Get ProcAddress() As Long

    ProcAddress = lpProcAddress
    
End Property
