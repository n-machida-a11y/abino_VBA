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

    fileName = activeSht.Range(CELL_INVOICE_FILE_NAME).Value
    monthYear = activeSht.Range(CELL_MONTH_COL1).Value & activeSht.Range(CELL_MONTH_COL2).Value

    baseFolder = Environ("USERPROFILE") & "\Downloads"
    invoiceFolder = baseFolder & "\" & monthYear & "請求書"

    If Dir(invoiceFolder, vbDirectory) = "" Then MkDir invoiceFolder

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
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical
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

    If Dir(mPath) = "" Then
        MsgBox "マスタファイルが見つかりません。", vbCritical
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

