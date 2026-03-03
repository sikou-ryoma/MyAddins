Attribute VB_Name = "PDFExportManager"
Option Explicit

Private Const MODULE_NAME As String = "PDFExportManager"
Private FO As FileOjt


'---コントローラー
Public Sub ExportSelectedSheets_PDF()

    Dim targetSheets As Sheets
    
    Set FO = New FileOjt
    Set targetSheets = ActiveWindow.SelectedSheets
    Call ExportSheetsEachToPDF(targetSheets)

End Sub


'---選択中の複数シートからPDFファイルを出力する
'---1シート1ファイルで出力される
Private Sub ExportSheetsEachToPDF(ByVal targetSheets As Sheets)

    Dim ws As Worksheet
    Dim pdfFolder As Variant
    Dim pdfPath As String
    
    pdfFolder = GetPdfOutputFolder()
    
    If pdfFolder = False Then
        MsgBox "処理を中断します。", vbExclamation, MODULE_NAME
        Exit Sub
    End If
    
    For Each ws In targetSheets
        
        '---念のためにワークシートのみに絞る
        If TypeName(ws) <> "Worksheet" Then GoTo ContinueLoop
        
        ws.Select Replace:=True
        
        pdfPath = pdfFolder & "\" & ws.Name & ".pdf"
        
        ws.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            fileName:=pdfPath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        
ContinueLoop:
    Next ws
    
    MsgBox "出力が完了しました。", vbInformation, MODULE_NAME

End Sub


'---保存先フォルダを選択するダイアログを表示
Private Function GetPdfOutputFolder() As Variant

    Dim basePath As String
    Dim wbName As String
    
    wbName = GetWorkbookBaseName(ActiveWorkbook)
    
    FO.folderPath = FO.GetFolderPath
    
    If FO.folderPath = False Then
        GetPdfOutputFolder = False
        Exit Function
    End If
    
    basePath = FO.folderPath & "\" & Format(Now, "yyyymmdd_hhnnss") & wbName & "_PDF"
    
    If Dir(basePath, vbDirectory) = "" Then
        MkDir basePath
    End If
    
    GetPdfOutputFolder = basePath

End Function


'---拡張子を以外の名前を取り出す処理
Private Function GetWorkbookBaseName(wb As Workbook) As String

    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetWorkbookBaseName = fso.GetBaseName(wb.Name)
    
End Function

