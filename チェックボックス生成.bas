Attribute VB_Name = "�`�F�b�N�{�b�N�X"
Sub �����̃`�F�b�N�{�b�N�X���쐬����()

    Dim objCell As Object
    
    '�I�������͈͕��̃I�u�W�F�N�g���擾���܂�
    For Each objCell In Selection
    
        With objCell
        '�I�������Z���̈ʒu�փ`�F�b�N�{�b�N�X��z�u���A�Z���̃T�C�Y�ɒ������܂��B
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

Sub �I��͈͂̃`�F�b�N�{�b�N�X�I��()

Dim cb As CheckBox

    '�I��͈͂̃`�F�b�N�{�b�N�X�����[�v
    For Each cb In ActiveSheet.CheckBoxes
        If Not Application.Intersect(cb.TopLeftCell, Selection) Is Nothing Then
            cb.Value = True
        End If
    Next cb

End Sub

Sub �I��͈͂̃`�F�b�N�{�b�N�X�I�t()

Dim cb As CheckBox

    '�I��͈͂̃`�F�b�N�{�b�N�X�����[�v
    For Each cb In ActiveSheet.CheckBoxes
        If Not Application.Intersect(cb.TopLeftCell, Selection) Is Nothing Then
            cb.Value = False
        End If
    Next cb

End Sub

