Attribute VB_Name = "特定の色を含む行を抜き出し"
Sub ExtractGreenRows()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim destRow As Integer
    Dim rng As Range
    Dim cell As Range

    ' 新しいシートを作成して結果を保存（先頭に追加）
    Set destWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    destWs.Name = "GreenRows"
    destRow = 1

    ' 各シートからデータを抽出
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> destWs.Name Then
            Set rng = ws.UsedRange
            For Each cell In rng
                If cell.Interior.Color = RGB(146, 208, 80) Then ' みどり色の場合
                    cell.EntireRow.Copy Destination:=destWs.Rows(destRow)
                    destRow = destRow + 1
                    Exit For ' 行全体が条件に合う場合は次の行へ
                End If
            Next cell
        End If
    Next ws

    MsgBox "緑色の抽出完了！", vbInformation
End Sub

Sub ExtractBlueRows()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim destRow As Integer
    Dim rng As Range
    Dim cell As Range

    ' 新しいシートを作成して結果を保存（先頭に追加）
    Set destWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    destWs.Name = "BlueRows"
    destRow = 1

    ' 各シートからデータを抽出
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> destWs.Name Then
            Set rng = ws.UsedRange
            For Each cell In rng
                If cell.Interior.Color = RGB(0, 176, 240) Then ' あお色の場合
                    cell.EntireRow.Copy Destination:=destWs.Rows(destRow)
                    destRow = destRow + 1
                    Exit For ' 行全体が条件に合う場合は次の行へ
                End If
            Next cell
        End If
    Next ws

    MsgBox "青色の抽出完了！", vbInformation
End Sub

Sub ExtractBothColors()
    Call ExtractGreenRows
    Call ExtractBlueRows
End Sub

