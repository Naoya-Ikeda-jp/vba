Attribute VB_Name = "Diff_color"
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
    iColorFirst = 65535
    iColorSecond = 5296274
    iLeftExpand = 1             '// �I��͈͂��P�񍶑����w�i�F��ݒ�
    iRightExpand = 2            '// �I��͈͂��Q��E�����w�i�F��ݒ�
    
    '// �I���Z���͈͂�Range�I�u�W�F�N�g�ɐݒ�
    Set rSelect = Selection
    
    '// �I���Z���͈͂̍s���Ɨ񐔂��擾
    iRowCount = rSelect.Rows.Count
    iColCount = rSelect.Columns.Count
    
    MaxRow = Range("A1").End(xlDown).Row
    MaxCol = Range("A1").End(xlToRight).Column
    
    DiffCol = "L"
    'DiffCol2 = "I"
    
    '// �I���s�����[�v
    For iRow = 0 To MaxRow - 1
        
        '// �O��L�[�X�V
        sLastKey = sNowKey
        
        '// ����L�[�ݒ�p�ɏ�����
        sNowKey = Range(DiffCol & iRow + 1).Value
        'sNowKey = sNowKey & Range(DiffCol2 & iRow + 1).Value
        
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
        
        '// �I���Z���͈͂̌��ݍs�̔w�i�F��ݒ�
        Range(Cells(iRow + 1, 1), Cells(iRow + 1, MaxCol)).Interior.Color = iColor
    Next
End Sub
