Attribute VB_Name = "ふりがな"
'手順************************
'・ADデータをコピー
'displayName surname givenName
'mail    department  officeLocation  companyName
'・PHONETICを入れる
'・フリガナをいれる
'・VBA実行
'範囲選択
'Alt F11
'Ctrl G
'「selection.setphonetic」を実行
'************************
Sub ふりがな()
Selection.SetPhonetic
End Sub
