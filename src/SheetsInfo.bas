Attribute VB_Name = "SheetsInfo"
Option Explicit

Private Const MODULE_NAME As String = "SheetsInfo"

Private Enum eCol
    colIndex = 1
    colName
    colVisible
    colProtect
End Enum


'---アクティブなブックのシート一覧表を新規ブックで作成する
Public Sub GetAllSheets()

    Dim srcWb As Workbook
    Dim listWb As Workbook
    Dim ws As Worksheet
    
    On Error GoTo Finally
    
    Application.ScreenUpdating = False
    
    Set srcWb = ActiveWorkbook
    Set listWb = FindSheetListBook(srcWb)
    
    If listWb Is Nothing Then
        '---未生成であれば新規作成
        Set listWb = Workbooks.Add
        Set ws = listWb.Sheets(1)
        
        listWb.Windows(1).Caption = "[SheetList] " & srcWb.Name
        
        Call GetAllSheets_SetIndex(ws, srcWb.FullName)
    Else
        '---既存であれば更新
        Set ws = listWb.Worksheets("シート一覧")
        ws.Rows("3:" & ws.Rows.Count).Clear
    End If
    
    Call OutputSheetList(srcWb, ws)
 
    ws.Columns("A:D").AutoFit
    Call FreezeAt(ws, "A3")
    
Finally:
    Application.ScreenUpdating = True
    
End Sub


'---対象ブックの一覧ブックがすでに開いているか探す
Private Function FindSheetListBook(ByVal srcWb As Workbook) As Workbook

    Dim wb As Workbook
    Dim ws As Worksheet

    For Each wb In Application.Workbooks
        On Error Resume Next
        Set ws = wb.Worksheets("シート一覧")
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            If ws.Range("B1").Value = srcWb.FullName Then
                Set FindSheetListBook = wb
                Exit Function
            End If
        End If
        
        Set ws = Nothing
    Next wb

End Function


'---出力用ブックにインデックスをつける
Private Sub GetAllSheets_SetIndex(ByRef ws As Worksheet, ByVal bkName As String)

    ws.Name = "シート一覧"
    ws.Range("A1").Value = "WorkBook : "
    ws.Range("B1").Value = bkName
    ws.Cells(2, eCol.colIndex).Value = "SheetNo."
    ws.Cells(2, eCol.colName).Value = "SheetName"
    ws.Cells(2, eCol.colVisible).Value = "Visible"
    ws.Cells(2, eCol.colProtect).Value = "Protect"
    
End Sub


'---出力処理
Private Sub OutputSheetList(ByVal srcWb As Workbook, ByVal ws As Worksheet)

    Dim i As Long, j As Long
    Dim sh As Object

    For i = 1 To srcWb.Sheets.Count
        j = i + 2
        Set sh = srcWb.Sheets(i)
        Call GetAllSheets_SheetsName(ws, i, j, srcWb.FullName, sh.Name)
        Call GetAllSheets_Visible(sh, ws, j)
        Call GetAllSheets_ProtectContents(sh, ws, j)
    Next i

End Sub


'---シート名を記入
Private Sub GetAllSheets_SheetsName _
    (ByRef ws As Worksheet, ByVal i As Long, ByVal j As Long, ByVal bkName As String, ByVal shName As String)

    ws.Cells(j, eCol.colIndex).Value = i
    ws.Cells(j, eCol.colName).NumberFormat = "@"   '---一旦文字列で返して標準に戻す
    ws.Cells(j, eCol.colName).Value = shName
    ws.Cells(j, eCol.colName).NumberFormat = "General"
    
    ws.Hyperlinks.Add _
        Anchor:=ws.Cells(j, eCol.colName), _
        Address:=bkName, _
        SubAddress:="'" & shName & "'!A1"  '---アクセスしやすいようにハイパーリンクを付けておく

End Sub


'---シートの表示状態を記入
Private Sub GetAllSheets_Visible(ByVal sh As Object, ByRef ws As Worksheet, ByVal j As Long)

    Select Case sh.Visible
        Case xlSheetVisible
            ws.Cells(j, eCol.colVisible).Value = "表示"
        Case xlSheetHidden
            ws.Cells(j, eCol.colVisible).Value = "非表示"
        Case xlSheetVeryHidden
            ws.Cells(j, eCol.colVisible).Value = "VeryHidden"
    End Select

End Sub


'---シートの保護状態を記入
Private Sub GetAllSheets_ProtectContents(ByVal sh As Object, ByRef ws As Worksheet, ByVal j As Long)
    
    Select Case sh.ProtectContents
        Case True
            ws.Cells(j, eCol.colProtect).Value = "保護中"
        Case False
            ws.Cells(j, eCol.colProtect).Value = "保護解除中"
    End Select
        
End Sub


'---ウインドウ枠の固定
Public Sub FreezeAt(ByVal ws As Worksheet, ByVal cellAddress As String)

    Dim prevWs As Worksheet
    
    Set prevWs = ActiveSheet
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        ws.Range(cellAddress).Select
        .FreezePanes = True
    End With

    prevWs.Activate

End Sub

