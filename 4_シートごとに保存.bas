Attribute VB_Name = "シートごとに保存"
Sub sheets_save()
For Each シート In Worksheets
If (シート.name <> "macro" And シート.name <> "フォーマット" And シート.name <> "SS1データ") Then
シート.Copy
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & シート.name
ActiveWorkbook.Close
End If
Next シート
End Sub
