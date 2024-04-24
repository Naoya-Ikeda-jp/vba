Attribute VB_Name = "部署名部分マッチングと右に出力"
Sub ReplaceWithMatchingValues()

    ' ワークシート取得
    Dim dataSheet As Worksheet
    Dim refSheet As Worksheet
    Set dataSheet = ThisWorkbook.Worksheets("データ")
    Set refSheet = ThisWorkbook.Worksheets("参照")

    ' 変数初期化
    Dim dataRange As Range
    Dim refRange As Range
    Dim dataCell As Range
    Dim refCell As Range
    Dim replaceValue As String
    Dim matchedValue As String
    Dim startIndex As Long
    Dim endIndex As Long
    Dim str As String

    ' データシート範囲設定
    Set dataRange = dataSheet.Range("D38:D399")

    ' 参照シート範囲設定
    Set refRange = refSheet.Range("A4:A63")

    ' データシートループ
    For Each dataCell In dataRange
        ' マッチングフラグ初期化s
            matchedValue = ""

        ' データセット(マッチング用)
        str = dataCell.Value
        ' カウント
        match_cnt = 0
        
        ' 参照シートループ
        For Each refCell In refRange
            ' データセル値に参照セル値が含まれているか確認
            If InStr(str, refCell.Value) > 0 Then
                ' マッチングフラグ設定
                matchedValue = refCell.Value
                ' マッチングしたものを右に出力(H以降)
                dataCell.Offset(0, 4 + match_cnt).Value = matchedValue
                ' カウント増加
                match_cnt = match_cnt + 1

                ' 置換対象文字列取得
                replaceValue = InStr(str, refCell.Value)
                startIndex = replaceValue
                endIndex = startIndex + Len(refCell.Value) - 1

                ' データ保持(マッチング用)
                str = Replace(str, Mid(str, startIndex, endIndex), refCell.Offset(0, 2).Value)
            End If
        Next refCell
                ' マッチング結果をG列に出力
                dataCell.Offset(0, 3).Value = str

        ' マッチング結果のカウントをマッチング結果の右に出力
        dataCell.Offset(0, 4 + match_cnt).Value = "マッチング数:" & match_cnt
    Next dataCell

End Sub
