Attribute VB_Name = "HyperlinkReplace_Excel"
Sub HyperlinkReplace_Excel()
    Dim ws As Worksheet
    Dim hLink As Hyperlink
    Dim oldAddressPart As String
    Dim newAddressPart As String

    ' 検索文字列の入力を求める
    oldAddressPart = InputBox("検索文字列を入力してください（例：http://）", "検索文字列")
    If oldAddressPart = "" Then Exit Sub

    ' 置換文字列の入力を求める
    newAddressPart = InputBox("置換文字列を入力してください（例：https://）", "置換文字列")

    ' ワークシートごとにループ
    For Each ws In ActiveWorkbook.Worksheets
        ' ハイパーリンクごとにループ
        For Each hLink In ws.Hyperlinks
            ' ハイパーリンクのアドレスを置換
            If InStr(1, hLink.Address, oldAddressPart) > 0 Then
                hLink.Address = Replace(hLink.Address, oldAddressPart, newAddressPart)
            End If
        Next hLink
    Next ws

    MsgBox "ハイパーリンクのリンク先アドレスの置換が完了しました。", vbInformation, "置換完了"
End Sub
