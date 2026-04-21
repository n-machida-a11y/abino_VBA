Attribute VB_Name = "保存"
Option Explicit

Private Const CELL_INVOICE_FILE_NAME As String = "AK1"
Private Const CELL_ACCOUNTING_NO As String = "AF1"
Private Const CELL_MONTH_COL1 As String = "AE3"
Private Const CELL_MONTH_COL2 As String = "AB3"

'================================================================================
' 主要処理：保存＆印刷
'================================================================================
Public Sub 保存_印刷作業()
    Dim activeSht As Worksheet
    Set activeSht = ActiveSheet
    Dim initialScreen As Boolean
    initialScreen = Application.ScreenUpdating

    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    Call SyncInvoiceNoToMaster

    Dim baseFolder As String, invoiceFolder As String
    Dim fileName As String, monthYear As String

    fileName = CStr(activeSht.Range(CELL_INVOICE_FILE_NAME).Value)
    monthYear = CStr(activeSht.Range(CELL_MONTH_COL1).Value) & CStr(activeSht.Range(CELL_MONTH_COL2).Value)

    ' ファイル名に使えない文字(\/:*?<>|)を除去
    fileName = SanitizeFileName(fileName)
    monthYear = SanitizeFileName(monthYear)

    If Trim(fileName) = "" Then
        MsgBox "ファイル名が取得できません。" & vbCrLf & _
               "請求書シートのAK1セルにファイル名が設定されているか確認してください。", vbCritical
        Exit Sub
    End If

    baseFolder = Environ("USERPROFILE") & "\Downloads"
    invoiceFolder = baseFolder & "\" & monthYear & "請求書"

    On Error Resume Next
    Dim dirExists As String: dirExists = Dir(invoiceFolder, vbDirectory)
    On Error GoTo Cleanup
    If dirExists = "" Then MkDir invoiceFolder

    Dim newWb As Workbook
    Set newWb = Workbooks.Add
    activeSht.Copy Before:=newWb.Sheets(1)
    newWb.Sheets(1).Delete

    Dim filePath As String
    filePath = invoiceFolder & "\" & fileName & ".xlsx"
    newWb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close

    On Error Resume Next
    activeSht.Delete
    On Error GoTo 0

    Application.ScreenUpdating = initialScreen
    MsgBox "ファイルを保存しました。" & vbCrLf & "保存先：" & filePath, vbInformation
    Exit Sub

Cleanup:
    Application.ScreenUpdating = initialScreen
    Dim diag As String
    diag = "エラーが発生しました。" & vbCrLf & vbCrLf & _
           "【エラー詳細】" & vbCrLf & _
           " 番号: " & Err.Number & vbCrLf & _
           " 内容: " & Err.Description & vbCrLf & vbCrLf & _
           "【処理中の値】" & vbCrLf & _
           " ファイル名(AK1): [" & fileName & "]" & vbCrLf & _
           " 月年(AE3&AB3): [" & monthYear & "]" & vbCrLf & _
           " 保存フォルダ: [" & invoiceFolder & "]" & vbCrLf & _
           " 保存パス: [" & filePath & "]"
    MsgBox diag, vbCritical
End Sub

'================================================================================
' マスタ同期処理：請求書番号をマスタ依頼履歴に書き込む
' 【修正】複数の早期 Exit を廃止し、Cleanup ラベルで一元管理する。
'================================================================================
Private Sub SyncInvoiceNoToMaster()
    Dim mPath As String
    Dim wsSearch As Worksheet
    Dim wbMaster As Workbook
    Dim iraiNo As String
    Dim wsRireki As Worksheet
    Dim found As Range
    Dim targetRow As Long
    Dim accountingNo As String

    Set wsSearch = ThisWorkbook.Sheets(SHEET_IRAI_SEARCH)

    ' 【変更】Application.Run("GetMasterPath") → GetMasterPath() の直接呼び出し
    mPath = GetMasterPath()

    On Error GoTo Cleanup

    ' 空パス対策：Dir("") はエラー52(ファイル名が不正)を投げる
    If Trim(mPath) = "" Then
        MsgBox "マスタファイルのパスが設定されていません。" & vbCrLf & _
               "依頼履歴シートのG1セルにパスを設定してください。", vbCritical
        GoTo Cleanup
    End If
    Dim dirChk As String
    On Error Resume Next
    dirChk = Dir(mPath)
    On Error GoTo Cleanup
    If dirChk = "" Then
        MsgBox "マスタファイルが見つかりません。" & vbCrLf & "パス: " & mPath, vbCritical
        GoTo Cleanup
    End If

    Set wbMaster = Workbooks.Open(mPath)
    If wbMaster Is Nothing Then
        MsgBox "マスタファイルを開けません。", vbCritical
        GoTo Cleanup
    End If

    iraiNo = Trim(wsSearch.Range("A2").Value)
    If iraiNo = "" Then
        MsgBox "依頼NOが指定されていません。", vbCritical
        GoTo Cleanup
    End If

    Set wsRireki = wbMaster.Sheets(SHEET_IRAI_RIREKI)
    Set found = wsRireki.Columns("A").Find(What:=iraiNo, LookAt:=xlWhole)
    If found Is Nothing Then
        MsgBox "マスタ内に依頼NO [" & iraiNo & "] が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    targetRow = found.Row
    accountingNo = ActiveSheet.Range(CELL_ACCOUNTING_NO).Value
    wsRireki.Cells(targetRow, 2).Value = accountingNo

    wbMaster.Save

Cleanup:
    ' 正常時もエラー時も必ずここを通る
    If Not wbMaster Is Nothing Then wbMaster.Close False
    If Err.Number <> 0 Then
        MsgBox "マスタ同期中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    End If
End Sub

'================================================================================
' ファイル名に使えない文字を安全な記号で置換
'================================================================================
Private Function SanitizeFileName(ByVal name As String) As String
    Dim bad As String, i As Long, ch As String
    bad = "\/:*?""<>|"
    Dim result As String: result = name
    For i = 1 To Len(bad)
        ch = Mid(bad, i, 1)
        result = Replace(result, ch, "_")
    Next i
    SanitizeFileName = Trim(result)
End Function
