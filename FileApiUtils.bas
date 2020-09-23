Attribute VB_Name = "FileApiUtils"
Option Explicit

Public Type INT64
    LoDWord As Long
    HiDword As Long
End Type

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hfile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As FileOffsetTypes) As Long

Public Function ApiFileSeek(ByVal hfile As Long, Distance As INT64, ByVal MoveMethod As FileOffsetTypes) As INT64

Dim dout As INT64

dout.LoDWord = Distance.LoDWord
dout.HiDword = Distance.HiDword

dout.LoDWord = SetFilePointer(hfile, dout.LoDWord, dout.HiDword, MoveMethod)

If Err.LastDllError Then
    dout.LoDWord = -1
    dout.HiDword = 0
End If

ApiFileSeek = dout

End Function

Public Function GetCurrentFilePointer(ByVal hfile As Long) As INT64

Dim dout As INT64
Dim din As INT64

dout = ApiFileSeek(hfile, din, FILE_CURRENT)
GetCurrentFilePointer = dout

End Function
