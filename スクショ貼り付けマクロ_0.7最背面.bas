Attribute VB_Name = "�X�N�V���\��t���}�N��"
Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

Sub �X�N�V���\��t��()
    OpenClipboard
    EmptyClipboard
    CloseClipboard
    Dim CB As Variant
    Dim position As Integer: position = 33
    Dim size As Double: size = 1
    '�T�C�Y����
    size = 0.7
    Do While True
        CB = Application.ClipboardFormats
        If Not ActiveCell.Row = 1 Then
            If StrConv(ActiveCell.Offset(-1, 0).Value, vbUpperCase) = "EXIT" Then GoTo Quit
        End If
        On Error GoTo ErrorQuit
        For i = 1 To UBound(CB)
            If CB(i) = xlClipboardFormatBitmap Then
                ActiveSheet.Paste
                Set objShp = ActiveSheet.Shapes(Selection.Name)
                objShp.LockAspectRatio = msoTrue
                objShp.ScaleHeight size, msoTrue
                ActiveCell.Offset(position, 0).Select
                OpenClipboard
                EmptyClipboard
                CloseClipboard
                '�Ŕw��
                objShp.ZOrder msoSendToBack
            End If
        Next i
        DoEvents
    Loop

Quit:
    MsgBox "��~���܂����B", vbInformation
    ActiveCell.Offset(-1, 0).ClearContents
    GoTo ToEnd
ErrorQuit:
    MsgBox "�\�����ʓ���̂��ߒ�~���܂����B", vbInformation
ToEnd:
End Sub
