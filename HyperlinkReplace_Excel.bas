Attribute VB_Name = "HyperlinkReplace_Excel"
Sub HyperlinkReplace_Excel()
    Dim ws As Worksheet
    Dim hLink As Hyperlink
    Dim oldAddressPart As String
    Dim newAddressPart As String

    ' ����������̓��͂����߂�
    oldAddressPart = InputBox("�������������͂��Ă��������i��Fhttp://�j", "����������")
    If oldAddressPart = "" Then Exit Sub

    ' �u��������̓��͂����߂�
    newAddressPart = InputBox("�u�����������͂��Ă��������i��Fhttps://�j", "�u��������")

    ' ���[�N�V�[�g���ƂɃ��[�v
    For Each ws In ActiveWorkbook.Worksheets
        ' �n�C�p�[�����N���ƂɃ��[�v
        For Each hLink In ws.Hyperlinks
            ' �n�C�p�[�����N�̃A�h���X��u��
            If InStr(1, hLink.Address, oldAddressPart) > 0 Then
                hLink.Address = Replace(hLink.Address, oldAddressPart, newAddressPart)
            End If
        Next hLink
    Next ws

    MsgBox "�n�C�p�[�����N�̃����N��A�h���X�̒u�����������܂����B", vbInformation, "�u������"
End Sub
