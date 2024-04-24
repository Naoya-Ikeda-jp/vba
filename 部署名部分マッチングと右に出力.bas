Attribute VB_Name = "�����������}�b�`���O�ƉE�ɏo��"
Sub ReplaceWithMatchingValues()

    ' ���[�N�V�[�g�擾
    Dim dataSheet As Worksheet
    Dim refSheet As Worksheet
    Set dataSheet = ThisWorkbook.Worksheets("�f�[�^")
    Set refSheet = ThisWorkbook.Worksheets("�Q��")

    ' �ϐ�������
    Dim dataRange As Range
    Dim refRange As Range
    Dim dataCell As Range
    Dim refCell As Range
    Dim replaceValue As String
    Dim matchedValue As String
    Dim startIndex As Long
    Dim endIndex As Long
    Dim str As String

    ' �f�[�^�V�[�g�͈͐ݒ�
    Set dataRange = dataSheet.Range("D38:D399")

    ' �Q�ƃV�[�g�͈͐ݒ�
    Set refRange = refSheet.Range("A4:A63")

    ' �f�[�^�V�[�g���[�v
    For Each dataCell In dataRange
        ' �}�b�`���O�t���O������s
            matchedValue = ""

        ' �f�[�^�Z�b�g(�}�b�`���O�p)
        str = dataCell.Value
        ' �J�E���g
        match_cnt = 0
        
        ' �Q�ƃV�[�g���[�v
        For Each refCell In refRange
            ' �f�[�^�Z���l�ɎQ�ƃZ���l���܂܂�Ă��邩�m�F
            If InStr(str, refCell.Value) > 0 Then
                ' �}�b�`���O�t���O�ݒ�
                matchedValue = refCell.Value
                ' �}�b�`���O�������̂��E�ɏo��(H�ȍ~)
                dataCell.Offset(0, 4 + match_cnt).Value = matchedValue
                ' �J�E���g����
                match_cnt = match_cnt + 1

                ' �u���Ώە�����擾
                replaceValue = InStr(str, refCell.Value)
                startIndex = replaceValue
                endIndex = startIndex + Len(refCell.Value) - 1

                ' �f�[�^�ێ�(�}�b�`���O�p)
                str = Replace(str, Mid(str, startIndex, endIndex), refCell.Offset(0, 2).Value)
            End If
        Next refCell
                ' �}�b�`���O���ʂ�G��ɏo��
                dataCell.Offset(0, 3).Value = str

        ' �}�b�`���O���ʂ̃J�E���g���}�b�`���O���ʂ̉E�ɏo��
        dataCell.Offset(0, 4 + match_cnt).Value = "�}�b�`���O��:" & match_cnt
    Next dataCell

End Sub
