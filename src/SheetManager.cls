VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------
' ---SheetManager---
' ワークシートの情報を取得するユーティリティクラス
' 1シートで1つのインスタンスの生成となり、シート毎の行番号・列番号情報の保持、取得が可能
'------------------------------------------------------------------------
Option Explicit


Private Const CLASS_NAME As String = "[SheetManager]"

Private pSheet As Worksheet
Private pBook As Workbook

Private pStartRow As Long
Private pEndRow As Long
Private pStartCol As Long
Private pEndCol As Long

Private pRange As Range
Private pMatrix As Variant


'---インスタンス生成時の初期化処理
Private Sub Class_Initialize()
    Reset
End Sub


'---以下各プロパティの設定と取得


'---シートオブジェクトの設定・取得
Public Property Set sheet(ByRef ws As Worksheet)
    Set pSheet = ws
End Property

Public Property Get sheet() As Worksheet
    Set sheet = pSheet
End Property


'---ブックオブジェクトの設定・取得
Public Property Set book(ByRef wb As Workbook)
    Set pBook = wb
End Property

Public Property Get book() As Workbook
    Set book = pBook
End Property


'---指定した範囲をRange型変数へ代入・取得
Public Property Set rng(ByRef SelectRng As Range)
    Set pRange = SelectRng
End Property

Public Property Get rng() As Range
    Set rng = pRange
End Property


'---rngで取得したRangeオブジェクトをバリアント型配列へ代入・取得
Public Property Let matrix(ByVal rng As Variant)
    pMatrix = rng
End Property

Public Property Get matrix() As Variant
    matrix = pMatrix
End Property


'---開始行の設定・取得
Public Property Let startRow(ByVal sRow As Long)
    pStartRow = sRow
End Property

Public Property Get startRow() As Long
    startRow = pStartRow
End Property


'---終了行の設定・取得
Public Property Let endRow(ByVal eRow As Long)
    pEndRow = eRow
End Property

Public Property Get endRow() As Long
    endRow = pEndRow
End Property


'---開始列の設定・取得
Public Property Let startCol(ByVal sCol As Long)
    pStartCol = sCol
End Property

Public Property Get startCol() As Long
    startCol = pStartCol
End Property


'---終了列の設定・取得
Public Property Let endCol(ByVal eCol As Long)
    pEndCol = eCol
End Property

Public Property Get endCol() As Long
    endCol = pEndCol
End Property


'---データのある最終行の取得、引数を省略するとA列が指定される
Public Property Get lastRow(Optional ByVal tgtCol As Long = 1) As Long
    
    If pSheet Is Nothing Then
        lastRow = 0
    Else
        lastRow = pSheet.Cells(pSheet.Rows.Count, tgtCol).End(xlUp).Row
    End If

End Property


'---データのある最終列の取得、引数を省略すると1行目が指定される
Public Property Get lastCol(Optional ByVal tgtRow As Long = 1) As Long
    
    If pSheet Is Nothing Then
        lastCol = 0
    Else
        lastCol = pSheet.Cells(tgtRow, pSheet.Columns.Count).End(xlToLeft).Column
    End If

End Property


'---削除処理
Public Sub Clear()
    
    Set pSheet = Nothing
    Set pBook = Nothing
    Set pRange = Nothing
    pMatrix = Empty

    pStartRow = 0
    pEndRow = 0
    pStartCol = 0
    pEndCol = 0

End Sub


'---初期化処理
Public Sub Reset()

    Set pSheet = Nothing
    Set pBook = Nothing
    Set pRange = Nothing
    pMatrix = Empty

    pStartRow = 1
    pEndRow = 1
    pStartCol = 1
    pEndCol = 1

End Sub





'---設定したプロパティの確認用
Public Sub PrintStatus()

    Dim wbStatus As String, wsStatus As String, rngStatus As String
    
    If pBook Is Nothing Then
        wbStatus = "未設定"
    Else
        wbStatus = pBook.Name
    End If
    
    If pSheet Is Nothing Then
        wsStatus = "未設定"
    Else
        wsStatus = pSheet.Name
    End If
    
    If pRange Is Nothing Then
        rngStatus = "未設定"
    Else
        rngStatus = pRange.Address
    End If
    
    Debug.Print "===== クラス状態 ====="
    Debug.Print "■ Workbook: "; wbStatus
    Debug.Print "■ Worksheet: "; wsStatus
    Debug.Print "■ Range: "; rngStatus
    
    Debug.Print "■ Start Row: "; pStartRow
    Debug.Print "■ End Row: "; pEndRow
    Debug.Print "■ Start Col: "; pStartCol
    Debug.Print "■ End Col: "; pEndCol
    
    If IsEmpty(pMatrix) Then
        Debug.Print "■ Matrix: 未設定"
    ElseIf IsArray(pMatrix) Then
        On Error Resume Next
        Debug.Print "■ Matrix: 配列サイズ - " & _
                    LBound(pMatrix, 1) & " to " & UBound(pMatrix, 1) & " 行, " & _
                    LBound(pMatrix, 2) & " to " & UBound(pMatrix, 2) & " 列"
        If Err.Number <> 0 Then Debug.Print "■ Matrix: 配列 (多次元でないか、範囲不明)"
        On Error GoTo 0
    Else
        Debug.Print "■ Matrix: 単一値 - " & pMatrix
    End If
    Debug.Print "======================"
    
End Sub


