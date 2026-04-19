Attribute VB_Name = "管理者操作"
Option Explicit

'================================================================================
' 管理者用 メイン操作モジュール
'
' ■ 使い方
'   ・シート上にボタンを2つ配置してこのモジュールのマクロを割り当てる:
'       ・マクロ_最新取得  → マスタからこのブックに全シートをコピー
'       ・マクロ_マスタへ反映 → このブックの全シートをマスタへ書き戻し
'
' ■ 同期対象シート
'   管理者設定用コード.bas の GetSyncSheetNames() で定義
'================================================================================


'================================================================================
' 【ボタン1】マスタから最新データを取得
'  マスタファイルを読み取り専用で開き、対象シートの内容をこのブックに転写する
'================================================================================
Public Sub マクロ_最新取得()
    Dim masterPath As String
    Dim wbMaster As Workbook
    Dim sheetNames As Variant
    Dim i As Long
    Dim successCount As Long, skipCount As Long
    Dim logText As String

    masterPath = GetMasterPath()
    If Dir(masterPath) = "" Then
        MsgBox "マスタファイルにアクセスできません。" & vbCrLf & _
               "パス: " & masterPath & vbCrLf & vbCrLf & _
               "ネットワーク(Z:)の接続を確認してください。", vbCritical, "最新取得エラー"
        Exit Sub
    End If

    If MsgBox("マスタから最新データを取得します。" & vbCrLf & vbCrLf & _
              "このブックの対象シートは上書きされます。" & vbCrLf & _
              "よろしいですか？", vbYesNo + vbQuestion, "最新取得の確認") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Cleanup

    Set wbMaster = Workbooks.Open(fileName:=masterPath, ReadOnly:=True, UpdateLinks:=0)

    sheetNames = GetSyncSheetNames()

    For i = LBound(sheetNames) To UBound(sheetNames)
        Dim sheetName As String: sheetName = CStr(sheetNames(i))

        If Not SheetExistsIn(wbMaster, sheetName) Then
            logText = logText & " - " & sheetName & " (マスタに存在せずスキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If
        If Not SheetExistsIn(ThisWorkbook, sheetName) Then
            logText = logText & " - " & sheetName & " (ローカルに存在せずスキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If

        Dim srcSh As Worksheet, dstSh As Worksheet
        Set srcSh = wbMaster.Sheets(sheetName)
        Set dstSh = ThisWorkbook.Sheets(sheetName)

        ' 保護解除・フィルタクリア
        Call SafeUnprotect(dstSh)
        Call ClearAllFilters(dstSh)

        ' 既存内容をクリア → マスタの使用範囲をそのままコピー
        dstSh.Cells.Clear
        If srcSh.UsedRange.Cells.CountLarge > 0 Then
            srcSh.UsedRange.Copy
            dstSh.Range(srcSh.UsedRange.Address).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            dstSh.Range(srcSh.UsedRange.Address).PasteSpecial Paste:=xlPasteColumnWidths
        End If
        Application.CutCopyMode = False

        successCount = successCount + 1
        logText = logText & " - " & sheetName & " (取得成功)" & vbCrLf
NextSheet:
    Next i

    wbMaster.Close SaveChanges:=False
    Set wbMaster = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "最新取得が完了しました。" & vbCrLf & vbCrLf & _
           "成功: " & successCount & " シート" & vbCrLf & _
           "スキップ: " & skipCount & " シート" & vbCrLf & vbCrLf & _
           "詳細:" & vbCrLf & logText, vbInformation, "最新取得完了"
    Exit Sub

Cleanup:
    If Not wbMaster Is Nothing Then wbMaster.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "最新取得中にエラーが発生しました。" & vbCrLf & _
           Err.Description, vbCritical, "最新取得エラー"
End Sub


'================================================================================
' 【ボタン2】このブックの内容をマスタへ反映
'  マスタファイルを書き込み可能で開き、対象シートの内容をマスタに転写する
'================================================================================
Public Sub マクロ_マスタへ反映()
    Dim masterPath As String
    Dim wbMaster As Workbook
    Dim sheetNames As Variant
    Dim i As Long
    Dim successCount As Long, skipCount As Long
    Dim logText As String

    masterPath = GetMasterPath()
    If Dir(masterPath) = "" Then
        MsgBox "マスタファイルにアクセスできません。" & vbCrLf & _
               "パス: " & masterPath & vbCrLf & vbCrLf & _
               "ネットワーク(Z:)の接続を確認してください。", vbCritical, "マスタ反映エラー"
        Exit Sub
    End If

    If MsgBox("このブックの内容をマスタに反映（上書き）します。" & vbCrLf & vbCrLf & _
              "【注意】" & vbCrLf & _
              " ・マスタの同名シートは内容が全置換されます" & vbCrLf & _
              " ・他のユーザーがマスタを開いていると失敗します" & vbCrLf & _
              " ・反映前に「最新取得」で同期しておくと安全です" & vbCrLf & vbCrLf & _
              "本当に反映しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "マスタ反映の確認") = vbNo Then Exit Sub

    ' 他のExcelでマスタが既に開かれていないかチェック
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase(wb.FullName) = LCase(masterPath) Then
            MsgBox "マスタファイルが既にこのExcelで開かれています。" & vbCrLf & _
                   "閉じてから再実行してください。", vbCritical
            Exit Sub
        End If
    Next wb

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Cleanup

    Set wbMaster = Workbooks.Open(fileName:=masterPath, ReadOnly:=False, UpdateLinks:=0)

    If wbMaster.ReadOnly Then
        MsgBox "マスタファイルを書き込み可能で開けませんでした。" & vbCrLf & _
               "他のユーザーが編集中の可能性があります。", vbCritical
        GoTo Cleanup
    End If

    sheetNames = GetSyncSheetNames()

    For i = LBound(sheetNames) To UBound(sheetNames)
        Dim sheetName As String: sheetName = CStr(sheetNames(i))

        If Not SheetExistsIn(wbMaster, sheetName) Then
            logText = logText & " - " & sheetName & " (マスタに存在せずスキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If
        If Not SheetExistsIn(ThisWorkbook, sheetName) Then
            logText = logText & " - " & sheetName & " (ローカルに存在せずスキップ)" & vbCrLf
            skipCount = skipCount + 1
            GoTo NextSheet
        End If

        Dim srcSh As Worksheet, dstSh As Worksheet
        Set srcSh = ThisWorkbook.Sheets(sheetName)
        Set dstSh = wbMaster.Sheets(sheetName)

        ' マスタシートの保護解除・フィルタクリア
        Call SafeUnprotect(dstSh)
        Call ClearAllFilters(dstSh)

        ' 既存内容をクリア → ローカルの使用範囲をコピー
        dstSh.Cells.Clear
        If srcSh.UsedRange.Cells.CountLarge > 0 Then
            srcSh.UsedRange.Copy
            dstSh.Range(srcSh.UsedRange.Address).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
            dstSh.Range(srcSh.UsedRange.Address).PasteSpecial Paste:=xlPasteColumnWidths
        End If
        Application.CutCopyMode = False

        ' マスタシートを再保護（元々の保護設定に従う・必要なシートのみ）
        Select Case sheetName
            Case "工事番号一覧", "依頼履歴", "その他マスタ"
                Call SafeProtect(dstSh)
            Case Else
                ' その他のシートは保護せず（必要なら個別に追加）
        End Select

        successCount = successCount + 1
        logText = logText & " - " & sheetName & " (反映成功)" & vbCrLf
NextSheet:
    Next i

    wbMaster.Save
    wbMaster.Close SaveChanges:=False
    Set wbMaster = Nothing
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "マスタへの反映が完了しました。" & vbCrLf & vbCrLf & _
           "成功: " & successCount & " シート" & vbCrLf & _
           "スキップ: " & skipCount & " シート" & vbCrLf & vbCrLf & _
           "詳細:" & vbCrLf & logText, vbInformation, "マスタ反映完了"
    Exit Sub

Cleanup:
    If Not wbMaster Is Nothing Then wbMaster.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "マスタ反映中にエラーが発生しました。" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "【対処】" & vbCrLf & _
           " ・他ユーザーがマスタを開いていないか確認" & vbCrLf & _
           " ・Z:が接続されているか確認" & vbCrLf & _
           " ・問題解決後に再実行してください", vbCritical, "マスタ反映エラー"
End Sub
