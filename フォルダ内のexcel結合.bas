Attribute VB_Name = "�t�H���_����excel����"
'�_�C�A���O�Ŏw�肳�ꂽ�t�H���_����Excel�V�[�g�����s�u�b�N��
'�܂Ƃ߂�v���O����
Sub collectSheet()
    Dim thisBook As Workbook: Set thisBook = ThisWorkbook
    Dim fd As FileDialog
    Dim folderName As String
    '�Ώۂ̃t�H���_��I��
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = False Then
        Exit Sub
    End If
    '�t�H���_�p�X���i�[
    folderName = fd.SelectedItems(1)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    'Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim f As Object
    'Dim f As file
    Application.ScreenUpdating = False
    '�t�H���_������������
    For Each f In fso.GetFolder(folderName).Files
        'Excel���ǂ����A���g�Ɠ����t�@�C�����łȂ����A
        '�ꎞ�t�@�C��(~$)�ł͂Ȃ����̃`�F�b�N
        If fso.GetExtensionName(f.Path) Like "xls*" And _
           f.Name <> thisBook.Name And _
           Left(f.Name, 2) <> "~$" Then
            Dim baseName As String: baseName = fso.GetBaseName(f.Path)
            Dim book As Workbook: Set book = Workbooks.Open(f.Path, , True)
            Dim sh As Worksheet
            '���[�N�V�[�g���W�߂�
            For Each sh In book.Worksheets
                With thisBook
                    sh.Copy after:=.Worksheets(.Worksheets.Count)
                   '�V�[�g���͍ő��31�����̐���������̂�Left�֐����g��
              '�V�����V�[�g���̓u�b�N��.�V�[�g���Ƃ���
                    ActiveSheet.Name = Left(baseName & "." & sh.Name, 31)
                End With
            Next sh
            book.Close False
        End If
    Next f
    Application.ScreenUpdating = True
End Sub


