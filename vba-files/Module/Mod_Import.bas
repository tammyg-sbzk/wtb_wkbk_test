Attribute VB_Name = "Mod_Import"
'----------
' Import_Raw
' Import_PL
' Import_BS
' Import_TB
' Import_GL
' Import_MonthlyPL
'----------
Function Import_Raw(Raw_Sheet, Import_Type, Find_Hdr, OK_Cont)
Const VBA_Name As String = "Import_Raw"
Debug.Print "<" & VBA_Name & ">" & VBA_Name & "<"
Debug.Print "Disp_Sheet>" & Disp_Sheet & "<"
On Error GoTo ErrSub
Dim Src_WrkBk As Workbook
Dim Src_WrkSht, Target_Wrkbk As String
Dim This_Import As Variant
Dim Src_Range, Src_LastCol, This_Drive, Dash_Sheet, This_Path, This_Sheet, Tmp1_S As String
Dim Tmp_Str, Col_Cur, Col_Prev, Col_Calc As String
Dim Calc_DlrChg, Src_BegRow, Tmp1_I, Tmp_Row As Integer
Dim Src_LastRow As Long
Dim RngFound As Range

Const Tmp_Cell As String = "C1"
Const Dollar_Chg_Hdr As String = "$ Change Calculated"
Debug.Print "Raw_Sheet>" & Raw_Sheet & "<"
Tmp1_I = MsgBox("Do you wish to IMPORT a " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Current >" & Import_Type & "< download" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "This Will PURGE and Rebuild >" & Raw_Sheet & "<", vbQuestion + vbYesNo + vbDefaultButton2, "IMPORT " & Import_Type & " download")
If Tmp1_I = 7 Then
    OK_Cont = 0
    GoTo ExitRoutine
End If ' Tmp1_I = 7 = [Cancel]
If Tmp1_I = 6 Then
This_Sheet = ActiveSheet.Name
Dash_Sheet = This_Sheet
This_Path = Application.ActiveWorkbook.Path
This_Drive = Left(This_Path, 2)

ChDrive This_Drive
ChDir This_Path
This_Import = Application.GetOpenFilename(FileFilter:="MS-Excel Files (*.xlsx), *.xlsx", Title:="Browse .xlsx Files")
'Debug.Print This_Import
If This_Import = False Then
    Debug.Print "No Import"
    OK_Cont = 0
Else
 '   Debug.Print "Yes Import"
    Worksheets(Raw_Sheet).Visible = True
    FindColNumLtr Dash_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
    FindRow Dash_Sheet, "A", Tmp1_I, "<IMP_PATH>"
    Worksheets(Dash_Sheet).Range(Tmp1_S & Tmp1_I).Value = This_Import
    Target_Wrkbk = ActiveWorkbook.Name
    ' Purge Raw_Sheet
    FindLastRow Raw_Sheet, Tmp1_I
    If Tmp1_I > 0 Then Worksheets(Raw_Sheet).Range("A1:A" & Tmp1_I).EntireRow.Delete
    Set Src_WrkBk = Workbooks.Open(This_Import)
    Src_WrkSht = ActiveSheet.Name
    Src_BegRow = 0
    FindAnyWhere Src_WrkSht, Find_Hdr, Src_LastCol, Src_BegRow
Debug.Print "<!!-02-!!>" & Disp_Sheet & "<COL_01>"
    If Src_BegRow = 0 Then
        ' Build Raw Sheet for 'Cheap' QB
        Calc_DlrChg = 1 ' $ Change Column NOF on Raw Import, must build on RAW sheet
        FindLastColumn Src_WrkSht, Tmp1_I
        NumToLtr Tmp1_I, Src_LastCol
        For Tmp_Row = 15 To 1 Step -1
            If Len(Trim(Worksheets(Src_WrkSht).Cells(Tmp_Row, 1).Value)) < 1 Then
                Src_BegRow = Tmp_Row
                Exit For
            End If
        Next Tmp_Row
        FindLastRow Src_WrkSht, Src_LastRow
    Else
        ' Use variables from FindAnyWhere
        Calc_DlrChg = 0 ' $ Change Column in Download do NOT build on RAW sheet
        FindLastRow Src_WrkSht, Src_LastRow
    End If  ' Find_Hdr found or not
    Src_Range = "A" & Src_BegRow & ":" & Src_LastCol & Src_LastRow
    ActiveSheet.Range(Src_Range).Copy
    Workbooks(Target_Wrkbk).Worksheets(Raw_Sheet).Range("A1").PasteSpecial xlPasteValues
    Debug.Print "Target Raw Worksheet Activate>" & Raw_Sheet & "<"
    Workbooks(Target_Wrkbk).Worksheets(Raw_Sheet).Activate
    Application.CutCopyMode = False
    Src_WrkBk.Close SaveChanges:=False
    
    If Calc_DlrChg = 1 Then
        ' Build $ Change Column on Raw_Sheet
        FindLastColumn Raw_Sheet, Tmp1_I
        NumToLtr (Tmp1_I - 1), Col_Cur
        NumToLtr Tmp1_I, Col_Prev
        NumToLtr (Tmp1_I + 1), Col_Calc
        With Worksheets(Raw_Sheet).Range(Col_Calc & 1)
            .Value = Dollar_Chg_Hdr
            .Interior.Color = RGB(255, 0, 0)
            .Font.Color = RGB(255, 255, 0)
            .Font.Bold = True
        End With
        FindLastRow Raw_Sheet, Src_LastRow
        With Worksheets(Raw_Sheet)
            For Tmp_Row = 2 To Src_LastRow
                If IsEmpty(.Range(Col_Cur & Tmp_Row).Value) And IsEmpty(.Range(Col_Prev & Tmp_Row).Value) Then
                    ' Do nothing
                Else
                    Tmp_Str = "=.Range(" & Col_Cur & Tmp_Row & ")-.Range(" & Col_Prev & Tmp_Row & ")"
     '               Debug.Print "Tmp_Str >" & Tmp_Str & "<"
                    '.Range(Col_Calc & Tmp_Row).Formula = "=.Range(" & Col_Cur & Tmp_Row & ")-.Range(" & Col_Prev & Tmp_Row & ")"
                    .Range(Col_Calc & Tmp_Row).Value = .Range(Col_Cur & Tmp_Row).Value - .Range(Col_Prev & Tmp_Row).Value
                End If
            Next Tmp_Row
        End With
    Else
        ' $ Change Column included in download
    End If
    ' Break Point Test
    FindLastColumn Raw_Sheet, Tmp1_I
    NumToLtr Tmp1_I, Tmp1_S
    Tmp_Str = "A" & ":" & Tmp1_S
    Worksheets(Raw_Sheet).Columns(Tmp_Str).AutoFit
End If  ' This_Import = False
End If  ' Tmp1_I = No

ExitRoutine:
Worksheets(Raw_Sheet).Visible = False
Debug.Print "Complete>" & VBA_Name & "<"
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Function Import_PL()
Const VBA_Name As String = "Import_PL"
Debug.Print ">" & VBA_Name & "<"
On Error GoTo ErrSub

Dim Time_Beg, Time_End
Dim WkSheet As Worksheet
Dim This_Sheet, Raw_Sheet, Fmt_Sheet, Tmp_Col, Tmp1_S, Tmp2_S, Tmp3_S, Tmp4_S, Src_Str As String
Dim I, Tmp_Row, OK_Cont, Tmp1_I, Tmp2_I, Tmp3_I, Tmp4_I As Integer

Const Use_Raw As String = "Raw_PL"
Const Use_Fmt As String = "PL_01"
Const Desc_Sheet As String = "Profit & Loss"
Const Find_Hdr As String = "Change"
'Const Find_Hdr As String = "$ Change"

This_Sheet = ActiveSheet.Name
OK_Cont = 1
Raw_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_Raw Then Raw_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_Fmt Then Fmt_Sheet = WkSheet.Name
Next WkSheet
If Raw_Sheet = "NOF" Then
    Tmp1_I = MsgBox(VBA_Name & " Can NOT find a Raw Import Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "RAW Import Worksheet Not Found")
    GoTo ExitRoutine
End If
Debug.Print "Name>" & Raw_Sheet & "<"
FindColNumLtr This_Sheet, 1, Tmp_Row, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PL>"
Worksheets(This_Sheet).Range(Tmp_Col & (Tmp_Row - 1)).Value = Now()
Import_Raw Raw_Sheet, Desc_Sheet, Find_Hdr, OK_Cont
If OK_Cont = 0 Then
    Tmp1_I = MsgBox("Import has been Canceled", vbInformation, "IMPORT CANCELED")
    GoTo ExitRoutine
End If
FindColNumLtr Raw_Sheet, 1, I, Tmp_Col, Find_Hdr
Debug.Print VBA_Name & ">" & Raw_Sheet & "<Find>" & Find_Hdr & "<I>" & I & "<Tmp_Col>" & Tmp_Col & "<"
'Worksheets(Raw_Sheet).Cells(1, (I - 2)).Value = "<DESCRIPTION>"
FindAnyWhere Fmt_Sheet, Find_Hdr, Tmp4_S, Tmp4_I
Debug.Print "Change Found At>" & Tmp4_S & Tmp4_I & "<"
Worksheets(Fmt_Sheet).Columns(Tmp4_S).Delete
FindRow Fmt_Sheet, "A", Tmp_Row, "<HDR-1>"
Tmp_Row = Tmp_Row + 1
FindLastRow Fmt_Sheet, I
If I >= Tmp_Row Then Worksheets(Fmt_Sheet).Range("A" & Tmp_Row & ":" & "A" & I).EntireRow.Delete
FindAnyWhere Raw_Sheet, Find_Hdr, Tmp_Col, Tmp_Row
Src_Str = "A" & Tmp_Row & ":" & Tmp_Col
FindLastRow Raw_Sheet, Tmp_Row
Src_Str = Src_Str & Tmp_Row
FindColNumLtr Fmt_Sheet, 1, I, Tmp_Col, "<COL_01>"
FindRow Fmt_Sheet, "A", I, "<HDR-1>"
I = I + 1
Worksheets(Raw_Sheet).Range(Src_Str).Copy
Worksheets(Fmt_Sheet).Range(Tmp_Col & I).PasteSpecial xlPasteValues
Application.CutCopyMode = False
FindAnyWhere Fmt_Sheet, Find_Hdr, Tmp_Col, Tmp_Row
Src_Str = Tmp_Row & ":" & Tmp_Col & Tmp_Row
FindColNumLtr Fmt_Sheet, 1, I, Tmp_Col, "<COL_01>"
Src_Str = Tmp_Col & Src_Str
Debug.Print "Header>" & Src_Str & "<"
With Worksheets(Fmt_Sheet).Range(Src_Str)
    .Interior.Color = RGB(144, 161, 105)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .BorderAround xlContinuous
End With
If Trim(Worksheets(Raw_Sheet).Range("A1").Value) > "" Then
    ' ONLINE - Col(A) Contains Data - Do Nothing
Else
    ' LOCAL - Col(A) is Blank - Delete Col(A)
End If
' Format Headers & Footers
FindColNumLtr Fmt_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_01>"
FindAnyWhere Fmt_Sheet, Find_Hdr, Tmp_Col, Tmp_Row
Tmp4_S = Left(Tmp_Col, 1)
Tmp4_I = InStr("ABCDEFGHIJKLMNOP", Tmp4_S)
Tmp3_I = Tmp4_I - 2
NumToLtr Tmp3_I, Tmp3_S
Tmp2_I = Tmp4_I - 3
NumToLtr Tmp2_I, Tmp2_S
FindRow Fmt_Sheet, "A", Tmp_Row, "<HDR-1>"
Src_Str = "=" & Desc_Sheet & " - " & "Name_Client"
Worksheets(Fmt_Sheet).Range(Tmp4_S & Tmp_Row).Formula = Src_Str
Src_Str = "=TEXT(Yr_End," & Chr(34) & "mm/dd/yyyy" & Chr(34) & ")"
Worksheets(Fmt_Sheet).Range(Tmp4_S & (Tmp_Row)).Formula = Src_Str
With Worksheets(Fmt_Sheet).Range(Tmp4_S & (Tmp_Row - 1) & ":" & Tmp4_S & Tmp_Row)
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With
Tmp_Row = Tmp_Row + 2
Worksheets(Fmt_Sheet).Range(Tmp1_S & Tmp_Row & ":" & Tmp4_S & Tmp_Row).Font.Bold = True
FindLastRow Fmt_Sheet, OK_Cont
For Fmt_Row = OK_Cont To Tmp_Row Step -1
    For I = 1 To Tmp2_I
        If InStr(UCase(Worksheets(Fmt_Sheet).Cells(Fmt_Row, I).Value), "TOTAL") > 0 Then
            Worksheets(Fmt_Sheet).Range(Tmp1_S & Fmt_Row & ":" & Tmp4_S & Fmt_Row).Font.Bold = True
            Worksheets(Fmt_Sheet).Rows(Fmt_Row + 1).Insert Shift:=xlDown
            NumToLtr I, Tmp_Col
            Worksheets(Fmt_Sheet).Range(Tmp_Col & Fmt_Row & ":" & Tmp4_S & Fmt_Row).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            Exit For
        End If
    Next I
Next Fmt_Row
With Worksheets(Fmt_Sheet).Columns(Tmp1_S & ":" & Tmp2_S)
    .HorizontalAlignment = xlLeft
    .ColumnWidth = 5
End With
Worksheets(Fmt_Sheet).Columns(Tmp2_S).AutoFit
With Worksheets(Fmt_Sheet).Columns(Tmp3_S & ":" & Tmp4_S)
    .ColumnWidth = 25
    .HorizontalAlignment = xlRight
    .NumberFormat = "#,##0.00_);(#,##0.00);"
End With
FindAnyWhere Fmt_Sheet, "Net Income", Tmp_Col, Tmp_Row
Debug.Print "Net Income Last Row>" & Tmp_Col & Tmp_Row & "<"
With Worksheets(Fmt_Sheet).Range(Tmp1_S & Tmp_Row & ":" & Tmp4_S & Tmp_Row)
    .Font.Bold = True
    .Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
    .Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
End With

FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PL>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = Now()
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_02>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PATH>"
This_Path = Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_04>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PL>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = This_Path
Worksheets(This_Sheet).Activate

ExitRoutine:
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Function Import_BS()
Const VBA_Name As String = "Import_BS"
Debug.Print ">" & VBA_Name & "<"
On Error GoTo ErrSub

Dim Time_Beg, Time_End
Dim WkSheet As Worksheet
Dim This_Sheet, Raw_Sheet, Fmt_Sheet, Tmp_Col, Tmp1_S, Tmp2_S, Tmp3_S, Tmp4_S, Src_Str As String
Dim I, Tmp_Row, OK_Cont, Tmp1_I, Tmp2_I, Tmp3_I, Tmp4_I As Integer

Const Use_Raw As String = "Raw_BS"
Const Use_Fmt As String = "BS_01"
Const Desc_Sheet As String = "Balance Statement"
Const Find_Hdr As String = "Change"
'Const Find_Hdr As String = "$ Change"

This_Sheet = ActiveSheet.Name
OK_Cont = 1
Raw_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_Raw Then Raw_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_Fmt Then Fmt_Sheet = WkSheet.Name
Next WkSheet
If Raw_Sheet = "NOF" Then
    Tmp1_I = MsgBox(VBA_Name & " Can NOT find a Raw Import Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "RAW Import Worksheet Not Found")
    GoTo ExitRoutine
End If
Debug.Print "Name>" & Raw_Sheet & "<"
FindColNumLtr This_Sheet, 1, Tmp_Row, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_BS>"
Worksheets(This_Sheet).Range(Tmp_Col & (Tmp_Row - 1)).Value = Now()
Debug.Print "Call Import_Raw >" & OK_Cont & "<"
Import_Raw Raw_Sheet, Desc_Sheet, Find_Hdr, OK_Cont
Debug.Print "Return Import_Raw >" & OK_Cont & "<"
If OK_Cont = 0 Then
    Tmp1_I = MsgBox("Import has been Canceled", vbInformation, "IMPORT CANCELED")
    GoTo ExitRoutine
End If
FindColNumLtr Raw_Sheet, 1, I, Tmp_Col, Find_Hdr
Debug.Print VBA_Name & ">" & Raw_Sheet & "<Find>" & Find_Hdr & "<I>" & I & "<Tmp_Col>" & Tmp_Col & "<"
'Worksheets(Raw_Sheet).Cells(1, (I - 2)).Value = "<DESCRIPTION>"
FindAnyWhere Fmt_Sheet, Find_Hdr, Tmp4_S, Tmp4_I
Debug.Print "Change Found At>" & Tmp4_S & Tmp4_I & "<"
Worksheets(Fmt_Sheet).Columns(Tmp4_S).Delete
FindRow Fmt_Sheet, "A", Tmp_Row, "<HDR-1>"
Tmp_Row = Tmp_Row + 1
FindLastRow Fmt_Sheet, I
If I >= Tmp_Row Then Worksheets(Fmt_Sheet).Range("A" & Tmp_Row & ":" & "A" & I).EntireRow.Delete
FindAnyWhere Raw_Sheet, Find_Hdr, Tmp_Col, Tmp_Row
Src_Str = "A" & Tmp_Row & ":" & Tmp_Col
FindLastRow Raw_Sheet, Tmp_Row
Src_Str = Src_Str & Tmp_Row
FindColNumLtr Fmt_Sheet, 1, I, Tmp_Col, "<COL_01>"
FindRow Fmt_Sheet, "A", I, "<HDR-1>"
I = I + 1
Worksheets(Raw_Sheet).Range(Src_Str).Copy
Worksheets(Fmt_Sheet).Range(Tmp_Col & I).PasteSpecial xlPasteValues
Application.CutCopyMode = False
FindAnyWhere Fmt_Sheet, Find_Hdr, Tmp_Col, Tmp_Row
Src_Str = Tmp_Row & ":" & Tmp_Col & Tmp_Row
FindColNumLtr Fmt_Sheet, 1, I, Tmp_Col, "<COL_01>"
Src_Str = Tmp_Col & Src_Str
Debug.Print "Header>" & Src_Str & "<"
With Worksheets(Fmt_Sheet).Range(Src_Str)
    .Interior.Color = RGB(144, 161, 105)
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .BorderAround xlContinuous
End With
If Trim(Worksheets(Raw_Sheet).Range("A1").Value) > "" Then
    ' ONLINE - Col(A) Contains Data - Do Nothing
Else
    ' LOCAL - Col(A) is Blank - Delete Col(A)
End If
' Format Headers & Footers
FindColNumLtr Fmt_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_01>"
FindAnyWhere Fmt_Sheet, Find_Hdr, Tmp_Col, Tmp_Row
Tmp4_S = Left(Tmp_Col, 1)
Tmp4_I = InStr("ABCDEFGHIJKLMNOP", Tmp4_S)
Tmp3_I = Tmp4_I - 2
NumToLtr Tmp3_I, Tmp3_S
Tmp2_I = Tmp4_I - 3
NumToLtr Tmp2_I, Tmp2_S
FindRow Fmt_Sheet, "A", Tmp_Row, "<HDR-1>"
Src_Str = "=" & Desc_Sheet & " - " & "Name_Client"
Worksheets(Fmt_Sheet).Range(Tmp4_S & Tmp_Row).Formula = Src_Str
Src_Str = "=TEXT(Yr_End," & Chr(34) & "mm/dd/yyyy" & Chr(34) & ")"
Worksheets(Fmt_Sheet).Range(Tmp4_S & (Tmp_Row)).Formula = Src_Str
With Worksheets(Fmt_Sheet).Range(Tmp4_S & (Tmp_Row - 1) & ":" & Tmp4_S & Tmp_Row)
    .Font.Bold = True
    .HorizontalAlignment = xlRight
End With
Tmp_Row = Tmp_Row + 2
Worksheets(Fmt_Sheet).Range(Tmp1_S & Tmp_Row & ":" & Tmp4_S & Tmp_Row).Font.Bold = True
FindLastRow Fmt_Sheet, OK_Cont
For Fmt_Row = OK_Cont To Tmp_Row Step -1
    For I = 1 To Tmp2_I
        If InStr(UCase(Worksheets(Fmt_Sheet).Cells(Fmt_Row, I).Value), "TOTAL") > 0 Then
            Worksheets(Fmt_Sheet).Range(Tmp1_S & Fmt_Row & ":" & Tmp4_S & Fmt_Row).Font.Bold = True
            Worksheets(Fmt_Sheet).Rows(Fmt_Row + 1).Insert Shift:=xlDown
            NumToLtr I, Tmp_Col
            Worksheets(Fmt_Sheet).Range(Tmp_Col & Fmt_Row & ":" & Tmp4_S & Fmt_Row).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            Exit For
        End If
    Next I
Next Fmt_Row
With Worksheets(Fmt_Sheet).Columns(Tmp1_S & ":" & Tmp2_S)
    .HorizontalAlignment = xlLeft
    .ColumnWidth = 5
End With
Worksheets(Fmt_Sheet).Columns(Tmp2_S).AutoFit
With Worksheets(Fmt_Sheet).Columns(Tmp3_S & ":" & Tmp4_S)
    .ColumnWidth = 25
    .HorizontalAlignment = xlRight
    .NumberFormat = "#,##0.00_);(#,##0.00);"
End With
FindAnyWhere Fmt_Sheet, "TOTAL LIABILITIES & EQUITY", Tmp_Col, Tmp_Row
If Tmp_Row = 0 Then
FindAnyWhere Fmt_Sheet, "TOTAL LIABILITIES AND EQUITY", Tmp_Col, Tmp_Row
End If
Debug.Print "Total Liabilities and Equity Last Row>" & Tmp_Col & Tmp_Row & "<"
With Worksheets(Fmt_Sheet).Range(Tmp1_S & Tmp_Row & ":" & Tmp4_S & Tmp_Row)
    .Font.Bold = True
    .Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
    .Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
End With

FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_BS>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = Now()
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_02>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PATH>"
This_Path = Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_04>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_BS>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = This_Path
Worksheets(This_Sheet).Activate

ExitRoutine:
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Function Import_TB()
Const VBA_Name As String = "Import_TB"
On Error GoTo ErrSub
Dim Time_Beg, Time_End
Dim WkSheet As Worksheet
Dim This_Sheet, Raw_Sheet, WTB_Sheet, Tmp_Col As String
Dim I, Tmp_Row, OK_Cont, WTB_Row, RAW_Beg, RAW_End As Integer

Dim Find_WTB(5) As String
Dim Col_WTB(5) As String
Dim Col_RAW(3) As String

Col_RAW(1) = "A"
Col_RAW(2) = "B"
Col_RAW(3) = "C"
Find_WTB(1) = "<DESC>"
Find_WTB(2) = "<BOOK>"
Find_WTB(3) = "<DR>"
Find_WTB(4) = "<CR>"
Find_WTB(5) = "<FINAL>"


Const Use_Raw As String = "Raw_TB"
Const Use_WTB As String = "WTB_01"
Const Desc_Sheet As String = "Trial Balance"
Const Find_Hdr As String = "Credit"

This_Sheet = ActiveSheet.Name
OK_Cont = 1
Raw_Sheet = "NOF"
WTB_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_Raw Then Raw_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
Next WkSheet
If Raw_Sheet <> "NOF" And WTB_Sheet <> "NOF" Then
    ' Fall thru & Execute
Else
    Tmp1_I = MsgBox(VBA_Name & " Can NOT find a Raw Import Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "RAW Import Worksheet Not Found")
    GoTo ExitRoutine
End If
Debug.Print VBA_Name & ">" & "<RAW>" & Raw_Sheet & "<WTB>" & WTB_Sheet & "<"
FindColNumLtr This_Sheet, 1, Tmp_Row, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_TB>"
Worksheets(This_Sheet).Range(Tmp_Col & (Tmp_Row - 1)).Value = Now()
Import_Raw Raw_Sheet, Desc_Sheet, Find_Hdr, OK_Cont ' Raw_Sheet Purged within Import_Raw
If OK_Cont = 0 Then
    Tmp1_I = MsgBox("Import has been Canceled", vbInformation, "IMPORT CANCELED")
    GoTo ExitRoutine
End If
FindColNumLtr Raw_Sheet, 1, I, Tmp_Col, Find_Hdr
Debug.Print VBA_Name & ">" & Raw_Sheet & "<Find>" & Find_Hdr & "<I>" & I & "<Tmp_Col>" & Tmp_Col & "<"
Worksheets(Raw_Sheet).Cells(1, (I - 2)).Value = "<DESCRIPTION>"
If Trim(Worksheets(Raw_Sheet).Range("A1").Value) > "" Then
    ' ONLINE - Col(A) Contains Data - Do Nothing
Else
    ' LOCAL - Col(A) is Blank - Delete Col(A)
    Worksheets(Raw_Sheet).Cells(1, (I + 2)).Value = "Column (A) was BLANK & DELETED"
    Worksheets(Raw_Sheet).Range("A1").EntireColumn.Delete
    FindLastRow Raw_Sheet, Tmp_Row
    Worksheets(Raw_Sheet).Range("A" & Tmp_Row).Value = "<TOTAL>"
End If
Debug.Print "Purge>" & WTB_Sheet; "<"
FindRow WTB_Sheet, "A", WTB_Row, "<HDR>"
WTB_Row = WTB_Row + 1
FindLastRow WTB_Sheet, Tmp_Row
Worksheets(WTB_Sheet).Unprotect
If Tmp_Row >= WTB_Row Then Worksheets(WTB_Sheet).Range("A" & WTB_Row & ":" & "A" & Tmp_Row).EntireRow.Delete
Debug.Print "Build WTB Arrays"
For I = 1 To 5
    FindColNumLtr WTB_Sheet, 1, Tmp_Row, Col_WTB(I), Find_WTB(I)
Next I
RAW_Beg = 2
FindRow Raw_Sheet, "A", RAW_End, "TOTAL"
RAW_End = RAW_End - 1
For I = RAW_Beg To RAW_End
    Worksheets(WTB_Sheet).Range(Col_WTB(1) & WTB_Row).Value = Worksheets(Raw_Sheet).Range(Col_RAW(1) & I).Value
    Worksheets(WTB_Sheet).Range(Col_WTB(2) & WTB_Row).Value = (Worksheets(Raw_Sheet).Range(Col_RAW(2) & I).Value - Worksheets(Raw_Sheet).Range(Col_RAW(3) & I).Value)
    'ADD DESCRIPTION - TAMMY
    Worksheets(WTB_Sheet).Range("D" & WTB_Row).Formula = "=IF(OR(ISNUMBER(SEARCH(""LIABILI"",XLOOKUP(E" & WTB_Row & ",'Chart of Accounts'!A:A,'Chart of Accounts'!B:B,""""))),ISNUMBER(SEARCH(""CREDIT CARD"",XLOOKUP(E" & WTB_Row & ",'Chart of Accounts'!A:A,'Chart of Accounts'!B:B,"""")))),""LIABILITIES"",IF(OR(ISNUMBER(SEARCH(""ASSET"",XLOOKUP(E" & WTB_Row & ",'Chart of Accounts'!A:A,'Chart of Accounts'!B:B,""""))),ISNUMBER(SEARCH(""BANK"",XLOOKUP(E" & WTB_Row & ",'Chart of Accounts'!A:A,'Chart of Accounts'!B:B,"""")))),""ASSETS"",IF(ISNUMBER(SEARCH(""OTHER"",XLOOKUP(E" & WTB_Row & ",'Chart of Accounts'!A:A,'Chart of Accounts'!B:B,""""))),""NET OTHER (INCOME)/EXPENSE"",UPPER(XLOOKUP(E" & WTB_Row & ",'Chart of Accounts'!A:A,'Chart of Accounts'!B:B,"""")))))"
    Worksheets(WTB_Sheet).Range("A" & WTB_Row).Value = "=""<"" & " & "D" & WTB_Row & "& "">""" 'test"" '""<"" & D" & WTB_Row & " & " > ""
    WTB_Row = WTB_Row + 1
Next I
FindRow WTB_Sheet, "A", RAW_Beg, "<Hdr>"
RAW_Beg = RAW_Beg + 1
FindLastRow WTB_Sheet, RAW_End
Worksheets(WTB_Sheet).Range(Col_WTB(5) & RAW_Beg).Formula = "=" & Col_WTB(2) & RAW_Beg & "+" & Col_WTB(3) & RAW_Beg & "+" & Col_WTB(4) & RAW_Beg
Worksheets(WTB_Sheet).Range(Col_WTB(5) & RAW_Beg).Copy Destination:=Worksheets(WTB_Sheet).Range(Col_WTB(5) & (RAW_Beg + 1) & ":" & Col_WTB(5) & RAW_End)


'Set Run Time
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_TB>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = Now()
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_02>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PATH>"
This_Path = Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_04>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_TB>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = This_Path

Worksheets(This_Sheet).Activate

ExitRoutine:
Worksheets(WTB_Sheet).Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Function Import_GL()
Const VBA_Name As String = "Import_GL"
Debug.Print "VBA_Name>" & VBA_Name & "<"
'On Error GoTo ErrSub
Dim Time_Beg, Time_End
Dim WkSheet As Worksheet
Dim This_Sheet, Raw_Sheet, Tmp_Col As String
Dim I, Tmp_Row, OK_Cont As Integer

Const Use_Raw As String = "Raw_GL"
Const Use_GL As String = "GL_01"
Const Desc_Sheet As String = "General Ledger"
Const Find_Hdr As String = "Balance"

This_Sheet = ActiveSheet.Name
OK_Cont = 1
Raw_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_Raw Then Raw_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_GL Then GL_Sheet = WkSheet.Name
Next WkSheet
If Raw_Sheet = "NOF" Then
    Tmp1_I = MsgBox(VBA_Name & " Can NOT find a Raw Import Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "RAW Import Worksheet Not Found")
    GoTo ExitRoutine
End If
Debug.Print VBA_Name & ">" & "<Raw_Sheet>" & Raw_Sheet & "<"
FindColNumLtr This_Sheet, 1, Tmp_Row, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_GL>"
Worksheets(This_Sheet).Range(Tmp_Col & (Tmp_Row - 1)).Value = Now()
Import_Raw Raw_Sheet, Desc_Sheet, Find_Hdr, OK_Cont
If OK_Cont = 0 Then
    Tmp1_I = MsgBox("Import has been Canceled", vbInformation, "IMPORT CANCELED")
    GoTo ExitRoutine
End If
FindRow Raw_Sheet, "A", Tmp_Row, "TOTAL"
FindLastRowOnSheet Raw_Sheet, I
Debug.Print "Raw_Sheet>" & Raw_Sheet & "<TOTAL>" & Tmp_Row & "<LastRow>" & I & "<"
If Tmp_Row = I Then
    ' LOCAL - Delete Col(A)
    Worksheets(Raw_Sheet).Range("A1").EntireColumn.Delete
    Worksheets(Raw_Sheet).Range("A" & I).Value = "<TOTAL>"
    Worksheets(Raw_Sheet).Cells(1, (I + 2)).Value = "Column (A) was BLANK & DELETED"
Else
    ' ONLINE - Keep Col(A)
End If

FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_GL>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = Now()
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_02>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PATH>"
This_Path = Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value
FindRow This_Sheet, "A", Tmp_Row, "<REBUILD_GL>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = "Need to Rebuild GL from GL Sheet"
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = ""
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_04>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_GL>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = This_Path
' Automatically Rebuild GL
Worksheets(GL_Sheet).Activate
Rebuild_GL
Debug.Print "!!! Dow - This_Sheet>" & This_Sheet & "<"
Worksheets(This_Sheet).Activate

ExitRoutine:
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Sub Import_MonthlyPL()
Const VBA_Name As String = "Import_MonthlyPL"
Debug.Print "VBA_Name>" & VBA_Name & "<"
On Error GoTo ErrSub
Dim Time_Beg, Time_End
Dim WkSheet As Worksheet
Dim This_Sheet, Raw_Sheet, Tmp_Col As String
Dim I, Tmp_Row, OK_Cont As Integer

Dim wsh As Worksheet
Dim Wbp1 As Workbook
Dim Target_Wrkbk As String
Dim This_Import As Variant
Dim Src_Range, Src_LastCol, This_Drive, Dash_Sheet, This_Path, Tmp1_S As String
Dim Tmp_Str, Col_Cur, Col_Prev, Col_Calc As String
Dim Calc_DlrChg, Src_BegRow, Tmp1_I As Integer
Dim Src_LastRow As Long
Dim RngFound As Range

Const Use_Raw As String = "Raw_MonthlyPL"


This_Sheet = ActiveSheet.Name
OK_Cont = 1
Raw_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_Raw Then Raw_Sheet = WkSheet.Name
    'If WkSheet.CodeName = Use_GL Then GL_Sheet = WkSheet.Name
Next WkSheet
If Raw_Sheet = "NOF" Then
    Tmp1_I = MsgBox(VBA_Name & " Can NOT find a Raw Import Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "RAW Import Worksheet Not Found")
    GoTo ExitRoutine
End If
FindColNumLtr This_Sheet, 1, Tmp_Row, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_MONTHLYPL>"
Worksheets(This_Sheet).Range(Tmp_Col & (Tmp_Row - 1)).Value = Now()
Set wsh = ThisWorkbook.Worksheets(Raw_Sheet)
Debug.Print "Raw_Sheet>" & Raw_Sheet & "<"
Tmp1_I = MsgBox("Do you wish to IMPORT a " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Current >" & Import_Type & "< download" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "This Will PURGE and Rebuild >" & Raw_Sheet & "<", vbQuestion + vbYesNo + vbDefaultButton2, "IMPORT " & Import_Type & " download")
If Tmp1_I = 7 Then
    OK_Cont = 0
    GoTo ExitRoutine
End If ' Tmp1_I = 7 = [Cancel]
If Tmp1_I = 6 Then
This_Sheet = ActiveSheet.Name
Dash_Sheet = This_Sheet
This_Path = Application.ActiveWorkbook.Path
This_Drive = Left(This_Path, 2)

'Dash_Sheet = "Dashboard"
ChDrive This_Drive
ChDir This_Path
This_Import = Application.GetOpenFilename(FileFilter:="MS-Excel Files (*.xlsx), *.xlsx", Title:="Browse .xlsx Files")
Worksheets(Raw_Sheet).Visible = True
FindColNumLtr Dash_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow Dash_Sheet, "A", Tmp1_I, "<IMP_PATH>"
Worksheets(Dash_Sheet).Range(Tmp1_S & Tmp1_I).Value = This_Import
Target_Wrkbk = ActiveWorkbook.Name
Set Wbp1 = Workbooks.Open(This_Import)
wsh.Range("A1:AE1000").Value = Wbp1.Sheets(1).Range("A1:AE1000").Value
wsh.Columns("A:AE").AutoFit
Wbp1.Close SaveChanges:=False
Application.ScreenUpdating = True
End If
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_MONTHLYPL>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = Now()
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_02>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PATH>"
This_Path = Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value
'FindRow This_Sheet, "A", Tmp_Row, "<REBUILD_MONTHLYPL>"
'Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = "Need to Rebuild GL from GL Sheet"
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = ""
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_04>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_MONTHLYPL>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = This_Path
ExitRoutine:
Exit Sub

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Sub

Sub Import_ChartOfAccounts()
Const VBA_Name As String = "Import_ChartOfAccounts"
Debug.Print "VBA_Name>" & VBA_Name & "<"
On Error GoTo ErrSub
Dim Time_Beg, Time_End
Dim WkSheet As Worksheet
Dim This_Sheet, Raw_Sheet, Tmp_Col As String
Dim I, Tmp_Row, OK_Cont As Integer

Dim wsh As Worksheet
Dim Wbp1 As Workbook
Dim Target_Wrkbk As String
Dim This_Import As Variant
Dim Src_Range, Src_LastCol, This_Drive, Dash_Sheet, This_Path, Tmp1_S As String
Dim Tmp_Str, Col_Cur, Col_Prev, Col_Calc As String
Dim Calc_DlrChg, Src_BegRow, Tmp1_I As Integer
Dim Src_LastRow As Long
Dim RngFound As Range

Const Use_Raw As String = "Raw_ChartOfAccounts"


This_Sheet = ActiveSheet.Name
OK_Cont = 1
Raw_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_Raw Then Raw_Sheet = WkSheet.Name
    'If WkSheet.CodeName = Use_GL Then GL_Sheet = WkSheet.Name
Next WkSheet
If Raw_Sheet = "NOF" Then
    Tmp1_I = MsgBox(VBA_Name & " Can NOT find a Raw Import Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "RAW Import Worksheet Not Found")
    GoTo ExitRoutine
End If
FindColNumLtr This_Sheet, 1, Tmp_Row, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_CHART>"
Worksheets(This_Sheet).Range(Tmp_Col & (Tmp_Row - 1)).Value = Now()
Set wsh = ThisWorkbook.Worksheets(Raw_Sheet)
Debug.Print "Raw_Sheet>" & Raw_Sheet & "<"
Tmp1_I = MsgBox("Do you wish to IMPORT a " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Current >Chart of Accounts< download" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "This Will PURGE and Rebuild >" & Raw_Sheet & "<", vbQuestion + vbYesNo + vbDefaultButton2, "IMPORT " & Import_Type & " download")
If Tmp1_I = 7 Then
    OK_Cont = 0
    GoTo ExitRoutine
End If ' Tmp1_I = 7 = [Cancel]
If Tmp1_I = 6 Then
This_Sheet = ActiveSheet.Name
Dash_Sheet = This_Sheet
This_Path = Application.ActiveWorkbook.Path
This_Drive = Left(This_Path, 2)

'Dash_Sheet = "Dashboard"
ChDrive This_Drive
ChDir This_Path
This_Import = Application.GetOpenFilename(FileFilter:="MS-Excel Files (*.xlsx), *.xlsx", Title:="Browse .xlsx Files")
Worksheets(Raw_Sheet).Visible = True
FindColNumLtr Dash_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow Dash_Sheet, "A", Tmp1_I, "<IMP_PATH>"
Worksheets(Dash_Sheet).Range(Tmp1_S & Tmp1_I).Value = This_Import
Target_Wrkbk = ActiveWorkbook.Name
Set Wbp1 = Workbooks.Open(This_Import)
wsh.Range("A1:AE1000").Value = Wbp1.Sheets(1).Range("A1:AE1000").Value
wsh.Columns("A:AE").AutoFit
If wsh.Range("A1").Value = "Account List" Then
    wsh.Range("A1:A3").EntireRow.Delete
ElseIf wsh.Range("A1").Value = "" Then
    wsh.Range("A1").EntireColumn.Delete
End If
wsh.Range("A1").EntireRow.Delete
wsh.Range("D1").Formula = "=IFERROR(XMATCH(""TOTAL"",A:A),XMATCH(,A:A))"
Dim numRows As Integer
numRows = wsh.Range("D1").Value
wsh.Range("A" & numRows & ":A" & numRows + 50).EntireRow.Delete
wsh.Visible = False
Wbp1.Close SaveChanges:=False
Application.ScreenUpdating = True
End If
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_CHART>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = Now()
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_02>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_PATH>"
This_Path = Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value
'FindRow This_Sheet, "A", Tmp_Row, "<REBUILD_MONTHLYPL>"
'Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = "Need to Rebuild GL from GL Sheet"
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_03>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = ""
FindColNumLtr This_Sheet, 1, I, Tmp_Col, "<COL_04>"
FindRow This_Sheet, "A", Tmp_Row, "<IMP_CHART>"
Worksheets(This_Sheet).Range(Tmp_Col & Tmp_Row).Value = This_Path
ExitRoutine:
Exit Sub

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Sub
