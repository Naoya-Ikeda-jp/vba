Attribute VB_Name = "����̐F���܂ލs�𔲂��o��"
Sub ExtractGreenRows()
    Dim ws As Worksheet
    Dim destWs As Worksheet
    Dim destRow As Integer
    Dim rng As Range
    Dim cell As Range

    ' �V�����V�[�g���쐬���Č��ʂ�ۑ��i�擪�ɒǉ��j
    Set destWs = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
    destWs.Name = "GreenRows"
    destRow = 1

    ' �e�V�[�g����f�[�^�𒊏o
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> destWs.Name Then
            Set rng = ws.UsedRange
            For Each cell In rng
                If cell.Interior.Color = RGB(146, 208, 80) Then ' �݂ǂ�F�̏ꍇ
                    cell.EntireRow.Copy Destination:=destWs.Rows(destRow)
                    destRow = destRow + 1
                    Exit For ' �s�S�̂������ɍ����ꍇ�͎��̍s��
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
    destRow = 1

    ' �e�V�[�g����f�[�^�𒊏o
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> destWs.Name Then
            Set rng = ws.UsedRange
            For Each cell In rng
                If cell.Interior.Color = RGB(0, 176, 240) Then ' �����F�̏ꍇ
                    cell.EntireRow.Copy Destination:=destWs.Rows(destRow)
                    destRow = destRow + 1
                    Exit For ' �s�S�̂������ɍ����ꍇ�͎��̍s��
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

