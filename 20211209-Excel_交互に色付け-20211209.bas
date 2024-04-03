Attribute VB_Name = "Excel_���݂ɐF�t��"

Sub SetColorSameKey()
    Dim iColCount                   '// �I���Z���͈̗͂�
    Dim iRowCount                   '// �I���Z���͈͂̍s��
    Dim iRow                        '// �s���[�v�J�E���^
    Dim iCol                        '// �񃋁[�v�J�E���^
    Dim rSelect     As Range        '// �I���Z���͈�
    Dim r           As Range        '// �Z���I��͈͂̈�ԍ��̗�̌��ݍsRange�I�u�W�F�N�g
    Dim sLastKey                    '// �O��s�̊e��̘A��������i���ꔻ��L�[�j
    Dim sNowKey                     '// ����s�̊e��̘A��������i���ꔻ��L�[�j
    Dim iColor                      '// �w�i�F
    Dim iColorFirst                 '// �P�F�ڂ̔w�i�F
    Dim iColorSecond                '// �Q�F���̔w�i�F
    Dim iLeftExpand                 '// �I���Z���͈͂�荶���Ŕw�i�F��ݒ肵������
    Dim iRightExpand                '// �I���Z���͈͂��E���Ŕw�i�F��ݒ肵������
    
    '// �����l�ݒ�
'    iColorFirst = RGB(255, 255, 204)
'    iColorSecond = RGB(255, 204, 255)
    iColorFirst = RGB(0, 255, 0)
    iColorSecond = RGB(255, 255, 0)
'    iLeftExpand = 1             '// �I��͈͂��P�񍶑����w�i�F��ݒ�
'    iRightExpand = 2            '// �I��͈͂��Q��E�����w�i�F��ݒ�
    iLeftExpand = InputBox("���������܂œh��Ԃ��Ώۂ�")             '// �I��͈͂��P�񍶑����w�i�F��ݒ�
    iRightExpand = InputBox("�E�������܂œh��Ԃ��Ώۂ�")            '// �I��͈͂��Q��E�����w�i�F��ݒ�
    
    '// �I���Z���͈͂�Range�I�u�W�F�N�g�ɐݒ�
    Set rSelect = Selection
    
    '// �I���Z���͈͂̍s���Ɨ񐔂��擾
    iRowCount = rSelect.Rows.Count
    iColCount = rSelect.Columns.Count
    
    '// �I���s�����[�v
    For iRow = 0 To iRowCount - 1
        '// ���ݍs�̈�ԍ��̃Z����Range�I�u�W�F�N�g�ɐݒ�
        '// �I���Z���͈͂̍��E�w�i�F�ݒ�̊�_�Z���Ƃ���
        Set r = rSelect.Cells(iRow + 1, 1)
        
        '// �O��L�[�X�V
        sLastKey = sNowKey
        
        '// ����L�[�ݒ�p�ɏ�����
        sNowKey = ""
        
        '// �I��񐔃��[�v
        For iCol = 0 To iColCount - 1
            '// �Z���l�𕶎���Ƃ��č���L�[�ɘA��
            sNowKey = sNowKey & CStr(rSelect.Cells(iRow + 1, iCol + 1).Value)
        Next
        
        '// �O��s�ƍ���s�̃Z���l���قȂ�ꍇ
        If sLastKey <> sNowKey Then
            '// �ݒ�w�i�F���P�F�ڂ̏ꍇ
            If iColor = iColorFirst Then
                '// �Q�F�ڂ�ݒ�
                iColor = iColorSecond
            Else
                '// �P�F�ڂ�ݒ�
                iColor = iColorFirst
            End If
        End If
        
        '// �����̊g���Z���ւ̔w�i�F��ݒ�
        Range(r.Offset(0, -1 * iLeftExpand), r.Offset(0, 0)).Interior.Color = iColor
        
        '// �I���Z���͈͂̌��ݍs�̔w�i�F��ݒ�
        rSelect.Range(Cells(iRow + 1, 1), Cells(iRow + 1, iColCount)).Interior.Color = iColor
        
        '// �E���̊g���Z���ւ̔w�i�F��ݒ�
        Range(r.Offset(0, 0), r.Offset(0, iColCount - 1 + iRightExpand)).Interior.Color = iColor
    Next
End Sub
