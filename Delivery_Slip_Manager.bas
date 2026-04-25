Attribute VB_Name = "Delivery_Slip_Manager"
Option Explicit

' ============================================================
' 定数定義（セル範囲・色などを一箇所で管理）
' ============================================================
Private Const CHECK_SHEET_NAME As String = "Sheet1"   ' チェック対象シート名
Private Const CHECK_START_ROW   As Integer = 2         ' チェック開始行
Private Const CHECK_END_ROW     As Integer = 6         ' チェック終了行
Private Const CHECK_COLUMN      As Integer = 2         ' チェック対象列（B列）


' ============================================================
' エントリポイント
' ============================================================
Public Sub CheckDeliverySlip()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CHECK_SHEET_NAME)

    ' 判定を行い、結果リストを取得
    Dim emptyRows() As Integer
    emptyRows = GetEmptyRows(ws)

    ' 判定結果に基づいてセルを色付け
    ApplyCellColors ws, emptyRows

    ' 処理結果をユーザーに通知
    NotifyResult emptyRows

End Sub


' ============================================================
' 判定ロジック：空欄の行番号を配列で返す
' 空欄がなければ要素数0の配列を返す
' ============================================================
Private Function GetEmptyRows(ws As Worksheet) As Integer()

    Dim result()  As Integer
    Dim count     As Integer
    count = 0

    ' 一時バッファ（最大行数分確保）
    Dim buf(CHECK_START_ROW To CHECK_END_ROW) As Integer
    Dim r As Integer

    For r = CHECK_START_ROW To CHECK_END_ROW
        If Trim(CStr(ws.Cells(r, CHECK_COLUMN).Value)) = "" Then
            buf(count) = r
            count = count + 1
        End If
    Next r

    ' 実際に見つかった分だけコピーして返す
    If count = 0 Then
        ReDim result(0 To -1)   ' 空配列
    Else
        ReDim result(0 To count - 1)
        Dim i As Integer
        For i = 0 To count - 1
            result(i) = buf(i)
        Next i
    End If

    GetEmptyRows = result

End Function


' ============================================================
' View操作：判定結果に従いセルの背景色を設定する
'   空欄行 → vbYellow、入力済み行 → xlNone（色なし）
' ============================================================
Private Sub ApplyCellColors(ws As Worksheet, emptyRows() As Integer)

    Dim r As Integer

    ' まず全行を「入力済み」扱いでリセット
    For r = CHECK_START_ROW To CHECK_END_ROW
        ws.Cells(r, CHECK_COLUMN).Interior.ColorIndex = xlNone
    Next r

    ' 空欄だった行だけ黄色に塗る
    Dim i As Integer
    For i = 0 To UBound(emptyRows)
        ws.Cells(emptyRows(i), CHECK_COLUMN).Interior.Color = vbYellow
    Next i

End Sub


' ============================================================
' 通知：処理結果をMsgBoxで表示する
' ============================================================
Private Sub NotifyResult(emptyRows() As Integer)

    Dim msg As String

    If UBound(emptyRows) < 0 Then
        ' 空欄なし
        msg = "すべての必須項目が入力されています。" & vbCrLf & _
              "（B" & CHECK_START_ROW & "〜B" & CHECK_END_ROW & " の色をクリアしました）"
        MsgBox msg, vbInformation, "チェック完了"
    Else
        ' 空欄あり：行番号をリストアップ
        Dim rowList As String
        rowList = ""
        Dim i As Integer
        For i = 0 To UBound(emptyRows)
            rowList = rowList & "  ・B" & emptyRows(i) & vbCrLf
        Next i

        msg = "以下の行が未入力です（黄色でハイライトしました）：" & vbCrLf & rowList
        MsgBox msg, vbExclamation, "入力チェック結果"
    End If

End Sub
