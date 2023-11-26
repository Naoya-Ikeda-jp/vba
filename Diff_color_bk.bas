Attribute VB_Name = "Diff_color"
Sub Diff_color()
Attribute Diff_color.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    
MaxRow = Range("A1").End(xlDown).Row
MaxCol = Range("A1").End(xlToRight).Column
Diff_col = "L"
color_swith = 0

For i = 0 To MaxRow
    
    
If (Val(Diff_col & i + 1) = Val(Diff_col & i + 2)) Then
    Range("A" & i + 1, Cells(i + 2, MaxCol)).Select

    If (color_swith = 0) Then
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        color_swith = 1
    Else
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = 5296274
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        color_swith = 0
    End If

Else
    GoTo Continue


End If



Continue:
Next

End Sub
