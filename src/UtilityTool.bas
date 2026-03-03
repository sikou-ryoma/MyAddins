Attribute VB_Name = "UtilityTool"
Option Explicit


'---マクロ起動時にアクティブなブックの全シートの選択セルをA1に戻す
'---アクティブシートは最終シートになる
Public Sub All_A1Cell()
Attribute All_A1Cell.VB_ProcData.VB_Invoke_Func = " \n14"

    Const PROC_NAME As String = "All_A1Cell"
    
    Dim wb As Workbook, ws As Worksheet, wsLoop As Worksheet
    
    Application.ScreenUpdating = False
    
    Set wb = ActiveWorkbook
    Set ws = wb.ActiveSheet
    
    For Each wsLoop In wb.Worksheets
        wsLoop.Activate
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        Range("A1").Select
    Next wsLoop
    
    ws.Select Replace:=True
    
    Application.ScreenUpdating = True
    
End Sub


'---アクティブセルに対する基本的な情報の詳細を一覧で表示
'---基本的にマクロの作成時の情報を得るために利用
Public Sub ActiveSheetInfo()
Attribute ActiveSheetInfo.VB_ProcData.VB_Invoke_Func = " \n14"

    Const PROC_NAME As String = "ActiveSheetInfo"
    
    Dim wb As Workbook
    Dim ws As New SheetManager
    Dim colNum As Long
    
    Set wb = ActiveWorkbook
    Set ws.sheet = wb.ActiveSheet
    colNum = ws.lastCol(ActiveCell.Row)
    
    '---アクティブブックの情報をメッセージボックスに表示
    MsgBox _
        "Workbook  :  " & wb.FullName & vbCrLf & _
        "Sheet  :  " & ws.sheet.Name & vbCrLf & _
        "ActiveCellAddress  :  " & ActiveCell.Address & vbCrLf & _
        "value  :  " & ActiveCell.Value & vbCrLf & _
        "SelectionAddress  :  " & Selection.Address & vbCrLf & _
        "EndRow  :  " & ws.lastRow(ActiveCell.Column) & vbCrLf & _
        "EndColumn  :  " & colNum & "  (" & ColumnNumberToLetter(colNum) & ")" & vbCrLf & _
        "Color  :  " & ActiveCell.Interior.Color & vbCrLf & _
        "Font Color  :  " & ActiveCell.Font.Color, vbInformation, _
        PROC_NAME
    
End Sub


'---選択されているセル範囲の再代入を行う
Public Sub UpdateCellValue()
    
    Const PROC_NAME As String = "UpdateCellValue"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range

    Set wb = ActiveWorkbook
    Set ws = wb.ActiveSheet
        
    If ws.ProtectContents = False Then
        Set rng = Selection
        rng.Value = rng.Value
        MsgBox "選択範囲の再代入を行いました。", vbInformation, PROC_NAME
    Else
        MsgBox "シートの保護を解除してください。", vbExclamation, PROC_NAME
    End If

End Sub
