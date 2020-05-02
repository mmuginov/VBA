Option Explicit

Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef GUID As Byte) As Long

Function GenerateGUID() As String
Dim ID(0 To 15) As Byte
Dim N As Long
Dim GUID As String
Dim Res As Long
Res = CoCreateGuid(ID(0))
For N = 0 To 15
    GUID = GUID & IIf(ID(N) < 16, "0", "") & Hex$(ID(N))
    If Len(GUID) = 8 Or Len(GUID) = 13 Or Len(GUID) = 18 Or Len(GUID) = 23 Then
        GUID = GUID & "-"
    End If
Next N
GenerateGUID = GUID
End Function
