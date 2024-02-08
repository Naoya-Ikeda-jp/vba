Attribute VB_Name = "特定シート集計"
Sub 特定シート集計()
    Dim sFile As String
    Dim sWB As Workbook, dWB As Workbook
    Dim dSheetCount As Long
    Dim i As Long
    'Const SOURCE_DIR As String = "C:\work\000_temp\"
    Dim SOURCE_DIR As String
    SOURCE_DIR = CurDir
    Const DEST_FILE As String = "C:\work\000_temp\特定シート集計.xlsm"
    Const Sh_name As String = "2.レイアウト定義"
    Const Save_name As String = "C:\work\000_temp\特定シート集計_実行後.xlsm"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '指定したフォルダ内にあるブックのファイル名を取得
    sFile = Dir(SOURCE_DIR & "*.xls")
    
    'フォルダ内にブックがなければ終了
    If sFile = "" Then Exit Sub
    
    '集約用ブックを作成
    Set dWB = Workbooks("特定シート集計.xlsm")
'    Set dWB = Workbooks.Add
    
    '集約用ブック作成時のシート数を取得
    dSheetCount = dWB.Worksheets.Count
    
    Do
        'コピー元のブックを開く
        If SOURCE_DIR & sFile <> DEST_FILE Then
        Set sWB = Workbooks.Open(Filename:=SOURCE_DIR & sFile, ReadOnly:=False)
        
        'コピー元のシートを集約用ブックにコピー
        sWB.Worksheets(Sh_name).Copy After:=dWB.Worksheets(dSheetCount)
        
        'シート名をファイル名の一部+シート名の値に変更
        ActiveSheet.Name = Left(sFile, 8) + Sh_name
'        'シート名をセルA1の値に変更
'        ActiveSheet.Name = Range("A1").Value
                
        'コピー元ファイルを閉じる
        sWB.Close
        End If
        
        '次のブックのファイル名を取得
        sFile = Dir()
    Loop While sFile <> ""
        
    '集約用ブック作成時にあったシートを削除
    Application.DisplayAlerts = False
    For i = dSheetCount To 1 Step -1
        dWB.Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True
      
    '集約用ブックを保存して閉じる
    dWB.SaveAs Filename:=Save_name
'    dWB.SaveAs Filename:=DEST_FILE
    dWB.Close
    
    Application.ScreenUpdating = False
End Sub


