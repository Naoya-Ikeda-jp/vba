Attribute VB_Name = "����̐F���܂ލs�𔲂��o��"
Sub ActivateFirstSheet()
    Worksheets(1).Activate
End Sub
Sub ActivateLastSheet()
    ' ���[�N�u�b�N���̍Ō�̃V�[�g���A�N�e�B�u�ɂ���
    Worksheets(Worksheets.Count).Activate
End Sub
Sub ExtractGreenRows()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim destRow As Integer
    Dim rng As Range
    Dim cell As Range

    ' �V�����V�[�g���쐬���Č��ʂ�ۑ��i�擪�ɒǉ��j
    Set destWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    destWs.Name = "GreenRows"
    destRow = 2 ' �f�[�^��2�s�ڂ���J�n
    headerCopied = False
    
    ' �e�V�[�g����f�[�^�𒊏o
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> destWs.Name Then
            Set rng = ws.UsedRange
            For Each cell In rng
                If cell.Interior.Color = RGB(146, 208, 80) Then ' �݂ǂ�F�̏ꍇ
                ' �w�b�_�s���R�s�[�i�ŏ��̃V�[�g�̃w�b�_�̂݁j
                If Not headerCopied Then
                    ws.Rows(1).Copy Destination:=destWs.Rows(1)
                    headerCopied = True
                    lastCol = destWs.Cells(1, destWs.Columns.Count).End(xlToLeft).Column
                    End If
                    cell.EntireRow.Copy Destination:=destWs.Rows(destRow)
                    ' �󔒗�ƃV�[�g����ǉ�
                    destWs.Cells(destRow, lastCol + 2).Value = ws.Name
                    destWs.Cells(destRow, lastCol + 1).Value = ws.Index
                    destRow = destRow + 1
'                    Exit For ' �s�S�̂������ɍ����ꍇ�͎��̍s��
                End If
            Next cell
        End If
    Next ws

    MsgBox "�ΐF�̒��o�����I", vbInformation
End Sub

Sub ExtractBlueRows()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim destRow As Integer
    Dim rng As Range
    Dim cell As Range

    ' �V�����V�[�g���쐬���Č��ʂ�ۑ��i�擪�ɒǉ��j
    Set destWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    destWs.Name = "BlueRows"
    destRow = 2 ' �f�[�^��2�s�ڂ���J�n
    headerCopied = False
    
    ' �e�V�[�g����f�[�^�𒊏o
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> destWs.Name Then
            Set rng = ws.UsedRange
            For Each cell In rng
                If cell.Interior.Color = RGB(0, 176, 240) Then ' �����F�̏ꍇ
                ' �w�b�_�s���R�s�[�i�ŏ��̃V�[�g�̃w�b�_�̂݁j
                If Not headerCopied Then
                    ws.Rows(1).Copy Destination:=destWs.Rows(1)
                    headerCopied = True
                    lastCol = destWs.Cells(1, destWs.Columns.Count).End(xlToLeft).Column
                End If
                    cell.EntireRow.Copy Destination:=destWs.Rows(destRow)
                    ' �󔒗�ƃV�[�g����ǉ�
                    destWs.Cells(destRow, lastCol + 2).Value = ws.Name
                    destWs.Cells(destRow, lastCol + 1).Value = ws.Index
                    destRow = destRow + 1
'                    Exit For ' �s�S�̂������ɍ����ꍇ�͎��̍s��
                    End If
            Next cell
        End If
    Next ws

    MsgBox "�F�̒��o�����I", vbInformation
End Sub

Sub ExtractBothColors()
    Call ExtractGreenRows
    Call ExtractBlueRows
End Sub

