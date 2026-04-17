VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HashProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------
' HashProvider
' 簡易ハッシュ生成用クラス
'-----------------------------------------------
Option Explicit



'---VBAの言語仕様上64bit版のみで動作します
Public Function DJB2(ByVal text As String) As Long
    Dim Hash As LongLong

    Dim i As Long
    Hash = 5381

    For i = 1 To Len(text)
        Hash = ((Hash * 33) + Asc(Mid$(text, i, 1))) And &H7FFFFFFF
    Next i

    DJB2 = CLng(Hash)
End Function

