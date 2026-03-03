Attribute VB_Name = "ProtectSwitcher"
Option Explicit

Private Const MODULE_NAME As String = "ProtectSwitcher"


'---シートの保護・解除のスイッチ
'---SheetProtectとSheetUnprotectを一か所で行うラッパー
Public Sub SwitchProtectSetting()

    If ActiveSheet.ProtectContents = False Then
        Call SheetProtect
    Else
        Call SheetUnprotect
    End If

End Sub


'---アクティブシートを保護する
Private Sub SheetProtect()
        
    If ActiveSheet.ProtectContents = True Then
        MsgBox "シートはロックされています。", vbExclamation, MODULE_NAME
    Else
        ActiveSheet.Protect
        MsgBox "シートをロックしました。", vbInformation, MODULE_NAME
    End If
    
End Sub


'---アクティブシートの保護を解除する
Private Sub SheetUnprotect()
    
    Dim userResponse As Integer
    userResponse = MsgBox("ロックを解除しますか？", vbYesNo + vbQuestion, MODULE_NAME)
    If userResponse = vbYes Then
        If ActiveSheet.ProtectContents = False Then
            MsgBox "シートはロックされていません。", vbExclamation, MODULE_NAME
        Else
            ActiveSheet.Unprotect
            MsgBox "ロックを解除しました。" & vbCrLf & "作業後は必ずシートを保護してください。", vbInformation, MODULE_NAME
        End If
    End If

End Sub

