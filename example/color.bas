Sub color()
'
' color Macro
'

'
    Range("A1:B5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub