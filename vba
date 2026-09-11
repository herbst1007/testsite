Option Explicit

Sub ExportPDF_ColumnH_Only()
    Dim ws As Worksheet
    Dim i As Long
    Dim facilityName As String
    Dim targetFlag As String ' H列の判定用変数
    Dim safeFileName As String
    Dim saveFolder As String
    Dim fd As FileDialog
   
    ' 1. 出力先フォルダをダイアログで選択
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "PDFを保存するフォルダを選択してください"
    If fd.Show = -1 Then
        saveFolder = fd.SelectedItems(1) & "\"
    Else
        MsgBox "処理をキャンセルしました。", vbExclamation, "キャンセル"
        Exit Sub
    End If
   
    ' 画面の更新を一時停止
    Application.ScreenUpdating = False
   
    ' 対象シートをセット
    Set ws = ActiveSheet
   
    ' 既にフィルターがかかっている場合は一度解除する
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
   
    ' 2. 4行目から1841行目までを1件ずつループ処理
    For i = 4 To 1841
        ' C列（3列目）の施設名称を取得
        facilityName = ws.Cells(i, 3).Value
       
        ' H列（8列目）のフラグを取得（★ここを7から8に変更しています）
        targetFlag = ws.Cells(i, 8).Value
       
        ' 施設名が空欄ではなく、かつH列が「○」の場合のみ処理を実行
        ' ※全角の「○」と半角の「〇(漢数字のゼロ)」の混在を許容
        If facilityName <> "" And (targetFlag = "○" Or targetFlag = "〇") Then
           
            ' 3. 3行目を見出し行として、C列(Field:=3)を現在の施設名でフィルター
            ws.Range("A3:M1841").AutoFilter Field:=3, Criteria1:=facilityName
           
            ' 4. ファイル名のサニタイズ（禁則文字や改行を除外）
            safeFileName = Trim(facilityName)
            safeFileName = Replace(safeFileName, vbCr, "") ' 改行を除去
            safeFileName = Replace(safeFileName, vbLf, "")
            safeFileName = Replace(safeFileName, "\", "＿")
            safeFileName = Replace(safeFileName, "/", "＿")
            safeFileName = Replace(safeFileName, ":", "＿")
            safeFileName = Replace(safeFileName, "*", "＿")
            safeFileName = Replace(safeFileName, "?", "＿")
            safeFileName = Replace(safeFileName, """", "＿")
            safeFileName = Replace(safeFileName, "<", "＿")
            safeFileName = Replace(safeFileName, ">", "＿")
            safeFileName = Replace(safeFileName, "|", "＿")
           
            ' 5. PDFとして出力
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                                   Filename:=saveFolder & safeFileName & ".pdf", _
                                   Quality:=xlQualityStandard, _
                                   IncludeDocProperties:=True, _
                                   IgnorePrintAreas:=False, _
                                   OpenAfterPublish:=False
        End If
    Next i
   
    ' 最後にフィルターを解除して元の状態に戻す
    ws.AutoFilterMode = False
   
    ' 画面の更新を再開
    Application.ScreenUpdating = True
   
    MsgBox "H列が「○」の施設のPDF出力が完了しました！", vbInformation, "出力完了"

End Sub
