Attribute VB_Name = "Diff_color"
Sub SetColorSameKey()
    Dim iColCount                   '// 選択セル範囲の列数
    Dim iRowCount                   '// 選択セル範囲の行数
    Dim iRow                        '// 行ループカウンタ
    Dim iCol                        '// 列ループカウンタ
    Dim rSelect     As Range        '// 選択セル範囲
    Dim r           As Range        '// セル選択範囲の一番左の列の現在行Rangeオブジェクト
    Dim sLastKey                    '// 前回行の各列の連結文字列（同一判定キー）
    Dim sNowKey                     '// 今回行の各列の連結文字列（同一判定キー）
    Dim iColor                      '// 背景色
    Dim iColorFirst                 '// １色目の背景色
    Dim iColorSecond                '// ２色名の背景色
    Dim iLeftExpand                 '// 選択セル範囲より左側で背景色を設定したい列数
    Dim iRightExpand                '// 選択セル範囲より右側で背景色を設定したい列数
    
    '// 初期値設定
    iColorFirst = 65535
    iColorSecond = 5296274
    iLeftExpand = 1             '// 選択範囲より１列左側も背景色を設定
    iRightExpand = 2            '// 選択範囲より２列右側も背景色を設定
    
    '// 選択セル範囲をRangeオブジェクトに設定
    Set rSelect = Selection
    
    '// 選択セル範囲の行数と列数を取得
    iRowCount = rSelect.Rows.Count
    iColCount = rSelect.Columns.Count
    
    MaxRow = Range("A1").End(xlDown).Row
    MaxCol = Range("A1").End(xlToRight).Column
    
    DiffCol = "L"
    'DiffCol2 = "I"
    
    '// 選択行数ループ
    For iRow = 0 To MaxRow - 1
        
        '// 前回キー更新
        sLastKey = sNowKey
        
        '// 今回キー設定用に初期化
        sNowKey = Range(DiffCol & iRow + 1).Value
        'sNowKey = sNowKey & Range(DiffCol2 & iRow + 1).Value
        
        '// 前回行と今回行のセル値が異なる場合
        If sLastKey <> sNowKey Then
            '// 設定背景色が１色目の場合
            If iColor = iColorFirst Then
                '// ２色目を設定
                iColor = iColorSecond
            Else
                '// １色目を設定
                iColor = iColorFirst
            End If
        End If
        
        '// 選択セル範囲の現在行の背景色を設定
        Range(Cells(iRow + 1, 1), Cells(iRow + 1, MaxCol)).Interior.Color = iColor
    Next
End Sub
