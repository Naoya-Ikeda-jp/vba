Attribute VB_Name = "フォルダ内のexcel結合"
'ダイアログで指定されたフォルダ内のExcelシートを実行ブックに
'まとめるプログラム
Sub collectSheet()
    Dim thisBook As Workbook: Set thisBook = ThisWorkbook
    Dim fd As FileDialog
    Dim folderName As String
    '対象のフォルダを選択
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    If fd.Show = False Then
        Exit Sub
    End If
    'フォルダパスを格納
    folderName = fd.SelectedItems(1)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    'Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim f As Object
    'Dim f As file
    Application.ScreenUpdating = False
    'フォルダ内を検索する
    For Each f In fso.GetFolder(folderName).Files
        'Excelかどうか、自身と同じファイル名でないか、
        '一時ファイル(~$)ではないかのチェック
        If fso.GetExtensionName(f.Path) Like "xls*" And _
           f.Name <> thisBook.Name And _
           Left(f.Name, 2) <> "~$" Then
            Dim baseName As String: baseName = fso.GetBaseName(f.Path)
            Dim book As Workbook: Set book = Workbooks.Open(f.Path, , True)
            Dim sh As Worksheet
            'ワークシートを集める
            For Each sh In book.Worksheets
                With thisBook
                    sh.Copy after:=.Worksheets(.Worksheets.Count)
                   'シート名は最大で31文字の制限があるのでLeft関数を使う
              '新しいシート名はブック名.シート名とする
                    ActiveSheet.Name = Left(baseName & "." & sh.Name, 31)
                End With
            Next sh
            book.Close False
        End If
    Next f
    Application.ScreenUpdating = True
End Sub


