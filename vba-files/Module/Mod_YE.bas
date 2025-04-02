Attribute VB_Name = "Mod_YE"
Sub updateMasterInfo()
    Dim WTBWkbk, MstrWkbk As Workbook
    Dim WTBWkst, MstrWkst As Worksheet
    Dim This_Import As Variant
    This_Import = "F:\1NET\Tammy Gorokhov\Miscellaneous\Entities Master List.xlsm"
    Application.EnableEvents = False
    Set WTBWkbk = ActiveWorkbook
    Set WTBWkst = ActiveSheet
    Set MstrWkbk = Workbooks.Open(This_Import)
    MstrWkbk.Activate
    Set MstrWkst = MstrWkbk.Sheets("ENTITIES")
    Dim lastRow As Long
    lastRow = MstrWkst.Cells(MstrWkst.Rows.Count, "F").End(xlUp).Row
    Dim index, tempRow As Long
    Dim tempCol As String
    For index = 5 To lastRow
        If WTBWkst.Range("B3").Value = MstrWkst.Range("F" & index).Value Then
            MstrWkst.Range("U" & index).Value = WTBWkst.Range("basis").Value
            MstrWkst.Range("V" & index).Value = WTBWkst.Range("qbVersion").Value
            MstrWkst.Range("W" & index).Value = WTBWkst.Range("officer").Value
            MstrWkst.Range("X" & index).Value = WTBWkst.Range("residentState").Value
            MstrWkst.Range("Y" & index).Value = WTBWkst.Range("pension").Value
        End If
    Next index
    MstrWkbk.Close SaveChanges:=True
    WTBWkbk.Activate
    WTBWkst.Activate
    Application.EnableEvents = True
    WTBWkst.Range("basis").Formula = "=IF(NOT(ISBLANK(B3)), IF(XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!C:C) <> 0, XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!C:C), """"), """")"
    WTBWkst.Range("qbVersion").Formula = "=IF(NOT(ISBLANK(B3)), IF(XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!D:D) <> 0, XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!D:D), """"), """")"
    WTBWkst.Range("officer").Formula = "=IF(NOT(ISBLANK(B3)), IF(XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!E:E) <> 0, XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!E:E), """"), """")"
    WTBWkst.Range("residentState").Formula = "=IF(NOT(ISBLANK(B3)), IF(XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!F:F) <> 0, XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!F:F), """"), """")"
    WTBWkst.Range("pension").Formula = "=IF(NOT(ISBLANK(B3)), IF(XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!G:G) <> 0, XLOOKUP(B3, 'ENTITY LIST'!B:B, 'ENTITY LIST'!G:G), """"), """")"
End Sub

