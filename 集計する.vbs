' ==============================================================================
' ? ベーカリーカフェ 売上自動集計スクリプト（VBScript）
' 
' 【注意・免責事項】
' ※本ファイルに含まれるデータは完全な架空のダミーデータです。
' ※安心のため、実行前にメモ帳で中身（コード）をご確認の上、自己責任でご利用ください。
' ==============================================================================
Option Explicit

Dim fso, currentDir, outFile
Set fso = CreateObject("Scripting.FileSystemObject")

' スクリプト自身の配置フォルダパスを取得
currentDir = fso.GetParentFolderName(WScript.ScriptFullName)
outFile = currentDir & "\完成_売上集計表.xlsx"

' 1. フォルダ内のCSVファイルを1つ自動検出（名前固定を解除）
Dim file, csvFile
csvFile = ""
For Each file In fso.GetFolder(currentDir).Files
    If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
        csvFile = file.Path
        Exit For
    End If
Next

If csvFile = "" Then
    MsgBox "【お知らせ】集計対象のCSVファイルが見つかりません。" & vbCrLf & vbCrLf & _
           "同じフォルダに集計したいCSVファイルを1つ置いてから、" & vbCrLf & _
           "再度ダブルクリックしてください。", vbExclamation, "CSVファイルなし"
    WScript.Quit
End If

' 2. 既存の完成ファイルが開かれてロックされていないか確認（存在すれば削除）
If fso.FileExists(outFile) Then
    On Error Resume Next
    fso.DeleteFile outFile, True
    If Err.Number <> 0 Then
        MsgBox "【お知らせ】「完成_売上集計表.xlsx」がExcelで開かれています。" & vbCrLf & _
               "ファイルを閉じてから、再度ダブルクリックしてください。", vbExclamation, "ファイル使用中"
        WScript.Quit
    End If
    On Error GoTo 0
End If

' 3. Excelアプリケーションを裏側で非表示起動
Dim xlApp, wbCsv, wbOut, wsSummary, wsDetail
Set xlApp = Nothing
Set wbCsv = Nothing
Set wbOut = Nothing

Dim hasError, errDetail
hasError = False
errDetail = ""

On Error Resume Next
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Or xlApp Is Nothing Then
    MsgBox "Excelの起動に失敗しました: " & Err.Description, vbCritical, "起動エラー"
    WScript.Quit
End If

xlApp.Visible = False
xlApp.DisplayAlerts = False
xlApp.ScreenUpdating = False

' 4. CSVを開く
Set wbCsv = xlApp.Workbooks.Open(csvFile)
If Err.Number <> 0 Then
    hasError = True
    errDetail = "CSVの読み込みに失敗しました: " & Err.Description
End If

If Not hasError Then
    ' 新規集計ブックを作成
    Set wbOut = xlApp.Workbooks.Add
    Set wsSummary = wbOut.Sheets(1)
    wsSummary.Name = "売上サマリー"

    ' 先に明細シートをコピーして同一ブック内に配置
    wbCsv.Sheets(1).Copy , wsSummary
    Set wsDetail = wbOut.Sheets(2)
    wsDetail.Name = "明細データ"

    ' CSV元ブックは閉じる（原本は一切上書きしない）
    wbCsv.Close False
    Set wbCsv = Nothing

    ' 5. サマリーシートの見出し・装飾
    wsSummary.Cells(2, 2).Value = "ベーカリーカフェ 本日売上集計表"
    wsSummary.Cells(2, 2).Font.Size = 16
    wsSummary.Cells(2, 2).Font.Bold = True
    wsSummary.Cells(2, 2).Font.Color = RGB(15, 23, 42)

    wsSummary.Cells(3, 2).Value = "集計対象: " & fso.GetFileName(csvFile) & " | 実行日時: " & Now()
    wsSummary.Cells(3, 2).Font.Size = 10
    wsSummary.Cells(3, 2).Font.Color = RGB(100, 116, 139)

    ' テーブルヘッダー
    wsSummary.Cells(5, 2).Value = "商品カテゴリ"
    wsSummary.Cells(5, 3).Value = "販売数量"
    wsSummary.Cells(5, 4).Value = "売上金額"
    wsSummary.Cells(5, 5).Value = "売上シェア"

    With wsSummary.Range("B5:E5")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(5, 150, 105) ' エメラルドグリーン
        .HorizontalAlignment = -4108 ' xlCenter
        .RowHeight = 26
        .VerticalAlignment = -4108
    End With

    ' カテゴリ集計（※本デモは簡易サンプルのため固定4カテゴリですが、実務では動的集計可能）
    Dim categories, i, rowIdx
    categories = Array("ベーカリー", "カフェドリンク", "サンドイッチ", "焼き菓子")

    For i = 0 To UBound(categories)
        rowIdx = 6 + i
        wsSummary.Cells(rowIdx, 2).Value = categories(i)
        
        wsSummary.Cells(rowIdx, 3).Formula = "=SUMIF(明細データ!$C:$C, B" & rowIdx & ", 明細データ!$F:$F)"
        wsSummary.Cells(rowIdx, 4).Formula = "=SUMIF(明細データ!$C:$C, B" & rowIdx & ", 明細データ!$G:$G)"
        wsSummary.Cells(rowIdx, 5).Formula = "=D" & rowIdx & "/$D$10"
        
        wsSummary.Cells(rowIdx, 3).NumberFormatLocal = "#,##0""点"""
        wsSummary.Cells(rowIdx, 4).NumberFormatLocal = """\""#,##0"
        wsSummary.Cells(rowIdx, 5).NumberFormatLocal = "0.0%"
        wsSummary.Rows(rowIdx).RowHeight = 22
        wsSummary.Rows(rowIdx).VerticalAlignment = -4108
    Next

    ' 合計行
    wsSummary.Cells(10, 2).Value = "総合計"
    wsSummary.Cells(10, 3).Formula = "=SUM(C6:C9)"
    wsSummary.Cells(10, 4).Formula = "=SUM(D6:D9)"
    wsSummary.Cells(10, 5).Formula = "=SUM(E6:E9)"

    With wsSummary.Range("B10:E10")
        .Font.Bold = True
        .Interior.Color = RGB(241, 245, 249)
        .RowHeight = 24
        .VerticalAlignment = -4108
    End With
    wsSummary.Cells(10, 3).NumberFormatLocal = "#,##0""点"""
    wsSummary.Cells(10, 4).NumberFormatLocal = """\""#,##0"
    wsSummary.Cells(10, 5).NumberFormatLocal = "0.0%"

    ' 罫線設定
    wsSummary.Range("B5:E10").Borders.LineStyle = 1
    wsSummary.Range("B5:E10").Borders.Color = RGB(203, 213, 225)
    wsSummary.Range("B10:E10").Borders(8).LineStyle = 1 ' 上線
    wsSummary.Range("B10:E10").Borders(9).LineStyle = 9 ' 下二重線

    ' 明細データ装飾
    With wsDetail.Range("A1:G1")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(15, 23, 42)
        .HorizontalAlignment = -4108
    End With
    wsDetail.Columns("E:G").NumberFormatLocal = "#,##0"

    ' 列幅調整と表示設定
    wsSummary.Columns("B:E").AutoFit
    wsDetail.Columns("A:G").AutoFit
    wsSummary.Activate
    xlApp.ActiveWindow.DisplayGridlines = True
    xlApp.Calculate

    ' 保存
    wbOut.SaveAs outFile, 51
    If Err.Number <> 0 Then
        hasError = True
        errDetail = "集計表の保存に失敗しました: " & Err.Description
    End If
End If

' 6. 【最重要】安全なクリーンアップ（エラー時でも絶対にプロセスを残さない）
If Not wbCsv Is Nothing Then
    wbCsv.Close False
    Set wbCsv = Nothing
End If

If Not wbOut Is Nothing Then
    wbOut.Close False
    Set wbOut = Nothing
End If

If Not xlApp Is Nothing Then
    xlApp.Quit
    Set xlApp = Nothing
End If

On Error GoTo 0

' 7. 結果通知
If hasError Then
    MsgBox "【エラー】集計処理中にエラーが発生しました。" & vbCrLf & _
           errDetail & vbCrLf & vbCrLf & _
           "※裏のExcelプロセスは安全に終了しました。", vbCritical, "集計中断"
Else
    MsgBox "【集計が完了しました！】" & vbCrLf & vbCrLf & _
           "Excel画面を一切開かずに、裏側で自動集計を行いました。" & vbCrLf & _
           "「完成_売上集計表.xlsx」を作成しました！" & vbCrLf & vbCrLf & _
           "ファイルを開いて内容をご確認ください。", vbInformation, "自動集計システム（VBScript）"
End If