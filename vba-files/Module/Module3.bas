Attribute VB_Name = "Module3"
Sub JOHN_MACRO()
'
' JOHN_MACRO Macro
' TBD
'

'
    Range("N14").Select
    ActiveCell.Formula2R1C1 = _
        "=IF(ISERROR(FIND(""="",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),1)),"""",SUMIF(C13,SUBSTITUTE(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),""="",""""),C12) + OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-2))"
    Range("N14").Select
    Selection.Copy
    Range("N15:N1000").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveWindow.ScrollRow = 14
End Sub


