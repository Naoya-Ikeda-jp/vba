Attribute VB_Name = "スクショ貼り付けマクロ"
Declare Function OpenClipboard Lib "user32" (Optional ByVal hwnd As Long = 0) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long

Sub スクショ貼り付け()
    OpenClipboard
    EmptyClipboard
    CloseClipboard
    Dim CB As Variant
    Dim position As Integer: position = 33
    Dim size As Double: size = 1
    'サイズ調整
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
                '最背面
                objShp.ZOrder msoSendToBack
            End If
        Next i
        DoEvents
    Loop

Quit:
    MsgBox "停止しました。", vbInformation
    ActiveCell.Offset(-1, 0).ClearContents
    GoTo ToEnd
ErrorQuit:
    MsgBox "予期せぬ動作のため停止しました。", vbInformation
ToEnd:
End Sub
