Attribute VB_Name = "�V�[�g���Ƃɕۑ�"
Sub sheets_save()
For Each �V�[�g In Worksheets
If (�V�[�g.name <> "macro" And �V�[�g.name <> "�t�H�[�}�b�g" And �V�[�g.name <> "SS1�f�[�^") Then
�V�[�g.Copy
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & �V�[�g.name
ActiveWorkbook.Close
End If
Next �V�[�g
End Sub
