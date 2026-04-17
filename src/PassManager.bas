Attribute VB_Name = "PassManager"
'----------------------------------------------------------------------
' ---PassManager---
'----------------------------------------------------------------------
Option Explicit



'---work

'---èoóÕópUDF
Public Function UDF_Hash(Optional ByVal val As Variant = "") As String
    UDF_Hash = GetDJB2Hash(CStr(val))
End Function

Private Function GetDJB2Hash(ByVal text As String) As String
    Dim hasher As HashProvider
    Set hasher = New HashProvider
    GetDJB2Hash = hasher.DJB2(text)
End Function

