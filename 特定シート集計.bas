Attribute VB_Name = "����V�[�g�W�v"
Sub ����V�[�g�W�v()
    Dim sFile As String
    Dim sWB As Workbook, dWB As Workbook
    Dim dSheetCount As Long
    Dim i As Long
    'Const SOURCE_DIR As String = "C:\work\000_temp\"
    Dim SOURCE_DIR As String
    SOURCE_DIR = CurDir
    Const DEST_FILE As String = "C:\work\000_temp\����V�[�g�W�v.xlsm"
    Const Sh_name As String = "2.���C�A�E�g��`"
    Const Save_name As String = "C:\work\000_temp\����V�[�g�W�v_���s��.xlsm"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '�w�肵���t�H���_���ɂ���u�b�N�̃t�@�C�������擾
    sFile = Dir(SOURCE_DIR & "*.xls")
    
    '�t�H���_���Ƀu�b�N���Ȃ���ΏI��
    If sFile = "" Then Exit Sub
    
    '�W��p�u�b�N���쐬
    Set dWB = Workbooks("����V�[�g�W�v.xlsm")
'    Set dWB = Workbooks.Add
    
    '�W��p�u�b�N�쐬���̃V�[�g�����擾
    dSheetCount = dWB.Worksheets.Count
    
    Do
        '�R�s�[���̃u�b�N���J��
        If SOURCE_DIR & sFile <> DEST_FILE Then
        Set sWB = Workbooks.Open(Filename:=SOURCE_DIR & sFile, ReadOnly:=False)
        
        '�R�s�[���̃V�[�g���W��p�u�b�N�ɃR�s�[
        sWB.Worksheets(Sh_name).Copy After:=dWB.Worksheets(dSheetCount)
        
        '�V�[�g�����t�@�C�����̈ꕔ+�V�[�g���̒l�ɕύX
        ActiveSheet.Name = Left(sFile, 8) + Sh_name
'        '�V�[�g�����Z��A1�̒l�ɕύX
'        ActiveSheet.Name = Range("A1").Value
                
        '�R�s�[���t�@�C�������
        sWB.Close
        End If
        
        '���̃u�b�N�̃t�@�C�������擾
        sFile = Dir()
    Loop While sFile <> ""
        
    '�W��p�u�b�N�쐬���ɂ������V�[�g���폜
    Application.DisplayAlerts = False
    For i = dSheetCount To 1 Step -1
        dWB.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True
      
    '�W��p�u�b�N��ۑ����ĕ���
    dWB.SaveAs Filename:=Save_name
'    dWB.SaveAs Filename:=DEST_FILE
    dWB.Close
    
    Application.ScreenUpdating = False
End Sub


