Attribute VB_Name = "チェックボックス"
Sub 複数のチェックボックスを作成する()

    Dim objCell As Object
    
    '選択した範囲分のオブジェクトを取得します
    For Each objCell In Selection
    
        With objCell
        '選択したセルの位置へチェックボックスを配置し、セルのサイズに調整します。
        ActiveSheet.CheckBoxes.Add(.Left, .Top, .Width, .Height).Select
            With Selection
                .Caption = ""
                .Value = xlOff
                .LinkedCell = objCell.Offset(0, 1).Address
                .Display3DShading = False
'                .Top = ActiveCell.Top
                .Left = ActiveCell.Left + (.Width - .Height) / 2
            End With
        
        End With
        
    Next

End Sub

Sub 選択範囲のチェックボックスオン()

Dim cb As CheckBox

    '選択範囲のチェックボックスをループ
    For Each cb In ActiveSheet.CheckBoxes
        If Not Application.Intersect(cb.TopLeftCell, Selection) Is Nothing Then
            cb.Value = True
        End If
    Next cb

End Sub

Sub 選択範囲のチェックボックスオフ()

Dim cb As CheckBox

    '選択範囲のチェックボックスをループ
    For Each cb In ActiveSheet.CheckBoxes
        If Not Application.Intersect(cb.TopLeftCell, Selection) Is Nothing Then
            cb.Value = False
        End If
    Next cb

End Sub

