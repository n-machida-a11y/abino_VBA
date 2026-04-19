Attribute VB_Name = "管理者設定用コード"
Option Explicit

'================================================================================
' 管理者用 共通設定モジュール
'
' 【このファイルの役割】
' 管理者用Excelが参照する共通の定数を一元管理する。
' シート名やマスタパスの変更があればここだけ修正すればよい。
'================================================================================

' ===== テストモード =====
' True: テスト用マスタパス(TEST_MASTER_PATH)を使う
' False: 本番用マスタパス(PROD_MASTER_PATH)を使う
Public Const IS_TEST_MODE As Boolean = True

Public Const TEST_MASTER_PATH As String = "Z:\Users\n-machida\Desktop\工事番号管理表.xlsm"
Public Const PROD_MASTER_PATH As String = "Z:\全社共有\建築事業部\30_事務\工事番号管理\工事番号管理表.xlsm"

' ===== シート保護パスワード =====
Public Const SHEET_PASSWORD As String = "3555"

' ===== 同期対象のシート名一覧 =====
' ここに列挙したシートが「取得」「反映」の対象になる。
' マスタ側とこのブック側で同じ名前のシートが必要。
Public Function GetSyncSheetNames() As Variant
    GetSyncSheetNames = Array( _
        "工事番号一覧", _
        "管理マスタ", _
        "その他マスタ", _
        "依頼書セル設定", _
        "2025.6～個人番号", _
        "2025.6～採番ルール", _
        "依頼履歴" _
    )
End Function

'================================================================================
' マスタファイルのパスを返す共通関数
'================================================================================
Public Function GetMasterPath() As String
    If IS_TEST_MODE Then
        GetMasterPath = TEST_MASTER_PATH
    Else
        GetMasterPath = PROD_MASTER_PATH
    End If
End Function

'================================================================================
' 指定ブック内に特定名のシートが存在するか
'================================================================================
Public Function SheetExistsIn(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExistsIn = Not (ws Is Nothing)
End Function

'================================================================================
' シート保護ヘルパー
'================================================================================
Public Sub SafeUnprotect(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub

Public Sub SafeProtect(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=SHEET_PASSWORD, _
               UserInterfaceOnly:=False, _
               AllowFiltering:=True, _
               AllowSorting:=True, _
               DrawingObjects:=True, _
               Contents:=True, _
               Scenarios:=True
    On Error GoTo 0
End Sub

' フィルタをクリアする
Public Sub ClearAllFilters(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
End Sub
