Attribute VB_Name = "Config"
Option Explicit

Public Const ADDIN_NAME As String = "MyTools"
Public Const VERSION As String = "v1.2.0"


'---ショートカット登録
'---Workbook_Openで使用
Public Sub RegisterShortcuts()

    On Error GoTo ErrHandler

    Application.OnKey "^+A", "'" & ThisWorkbook.Name & "'!All_A1Cell"
    Application.OnKey "^+1", "'" & ThisWorkbook.Name & "'!SwitchProtectSetting"
    Application.OnKey "^+I", "'" & ThisWorkbook.Name & "'!ActiveSheetInfo"
    Application.OnKey "^+2", "'" & ThisWorkbook.Name & "'!UpdateCellValue"
    Application.OnKey "^+3", "'" & ThisWorkbook.Name & "'!GetAllSheets"
    Application.OnKey "^+4", "'" & ThisWorkbook.Name & "'!ExportSelectedSheets_PDF"

    Exit Sub
    
ErrHandler:

    MsgBox "起動時ショートカット登録中にエラーが発生しました: " & Err.Description, vbExclamation, ADDIN_NAME
    
End Sub


'---ショートカット解除
'---Workbook_BeforeCloseで使用
Public Sub UnregisterShortcuts()
Attribute UnregisterShortcuts.VB_ProcData.VB_Invoke_Func = " \n14"

    On Error Resume Next
    
    Application.OnKey "^+A"
    Application.OnKey "^+1"
    Application.OnKey "^+I"
    Application.OnKey "^+2"
    Application.OnKey "^+3"
    Application.OnKey "^+4"

End Sub
