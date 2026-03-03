Attribute VB_Name = "Z_ColNumToLetter"
Option Explicit

Function ColumnNumberToLetter(ByVal colNum As Integer) As String
    Dim colLetter As String
    colLetter = ""
    
    Do While colNum > 0
        colNum = colNum - 1
        colLetter = Chr(65 + (colNum Mod 26)) & colLetter
        colNum = colNum \ 26
    Loop
    
    ColumnNumberToLetter = colLetter
End Function

