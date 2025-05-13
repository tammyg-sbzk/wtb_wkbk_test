Attribute VB_Name = "Mod_WTB"
'-------------
' WTB_Reconcile
' WTB_Subtotal_Del
' WTB_Subtotal_Refresh
'-------------
Function WTB_Reconcile()
    Const VBA_Name As String = "WTB_Reconcile"
    On Error GoTo ErrSub
    Debug.Print "Open >" & VBA_Name & "<Begin Color Formatting"
    Dim BS_Cnt, PL_Cnt, Ctl_Row, Ctl_Beg, Ctl_End, I, Tmp1_I, Tmp2_I, BS_BegRow, BS_EndRow, PL_BegRow, PL_EndRow As Integer
    Dim BS_Sheet, PL_Sheet, WTB_Sheet, BSdol_Col, PLdol_Col, Find_Ctl, Find_Lookup, Tmp1_S, Tmp2_S, Tmp_Col As String
    Dim Lookup_Sheet, Lookup_DolCol, Lookup_Flag, Lookup_Col As String
    Dim Lookup_Exit, Lookup_Cnt, Row_Cogs, Lookup_BegRow, Lookup_EndRow, Tmp_EndRow, Lookup_Row As Integer
    
    Dim CTL_Col()
    Dim BS_Col()
    Dim PL_Col()

    Dim lastCol As Integer
    
    
    Const Array_Cnt As Integer = 5
    Const T_Sheet As String = "WTB_TEST"
    Const C_Sheet As String = "Control"
    Const Use_BS As String = "BS_01"
    Const Use_PL As String = "PL_01"
    Const Use_WTB As String = "WTB_01"
    ReDim CTL_Col(Array_Cnt)
    
    ActiveSheet.Unprotect

    FindColNumLtr ActiveSheet.Name, 1, I, Tmp_Col, "<END_DEL>"
    
    FindLastRow T_Sheet, Tmp1_I
    If Tmp1_I > 0 Then
        Worksheets(T_Sheet).Range("A1:" & Tmp_Col & Tmp1_I).Delete Shift:=xlUp
    End If  ' Purge T_Sheet
    BS_Sheet = "NOF"
    PL_Sheet = "NOF"
    WTB_Sheet = "NOF"
    For Each WkSheet In ThisWorkbook.Worksheets
        If WkSheet.CodeName = Use_BS Then BS_Sheet = WkSheet.Name
        If WkSheet.CodeName = Use_PL Then PL_Sheet = WkSheet.Name
        If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
    Next WkSheet
    'WTB_Sheet
    FindColNumLtr WTB_Sheet, 1, I, WTB_Col, "<BOOK>"
    'BS_Sheet Parameters
    FindLastRow BS_Sheet, BS_EndRow
    FindRow BS_Sheet, "A", BS_BegRow, "<HDR-1>"
    BS_BegRow = BS_BegRow + 1
    FindColNumLtr BS_Sheet, BS_BegRow, I, Tmp1_S, "Change"
    NumToLtr (I - 2), BSdol_Col
    Ctl_End = I - 3
    FindColNumLtr BS_Sheet, 1, Ctl_Beg, Tmp2_S, "<COL_01>"
    Debug.Print "BS Range>" & Tmp2_S & (BS_BegRow + 1) & ":" & Tmp1_S & BS_EndRow & "<"
    Worksheets(BS_Sheet).Range(Tmp2_S & (BS_BegRow + 1) & ":" & Tmp1_S & BS_EndRow).Interior.Color = RGB(255, 255, 255)
    BS_Cnt = (Ctl_End - Ctl_Beg) + 1
    ReDim BS_Col(BS_Cnt)
    I = 0
    For Tmp1_I = Ctl_Beg To Ctl_End
    I = I + 1
    NumToLtr Tmp1_I, BS_Col(I)
    Next Tmp1_I
    'PL_Sheet Parameters
    FindLastRow PL_Sheet, PL_EndRow
    FindRow PL_Sheet, "A", PL_BegRow, "<HDR-1>"
    PL_BegRow = PL_BegRow + 1
    FindColNumLtr PL_Sheet, PL_BegRow, I, Tmp1_S, "Change"
    NumToLtr (I - 2), PLdol_Col
    Ctl_End = I - 3
    FindColNumLtr PL_Sheet, 1, Ctl_Beg, Tmp2_S, "<COL_01>"
    Debug.Print "PL Range>" & Tmp2_S & (PL_BegRow + 1) & ":" & Tmp1_S & PL_EndRow & "<"
    Worksheets(PL_Sheet).Range(Tmp2_S & (PL_BegRow + 1) & ":" & Tmp1_S & PL_EndRow).Interior.Color = RGB(255, 255, 255)
    PL_Cnt = (Ctl_End - Ctl_Beg) + 1
    ReDim PL_Col(PL_Cnt)
    I = 0
    For Tmp1_I = Ctl_Beg To Ctl_End
    I = I + 1
    NumToLtr Tmp1_I, PL_Col(I)
    Next Tmp1_I
    ' Ctl_Sheet Column Array
    For I = 1 To Array_Cnt
        Tmp1_S = Right((100 + I), 2)
        FindColNumLtr C_Sheet, 1, Tmp1_I, CTL_Col(I), "<COL_" & Tmp1_S & ">"
    Next I
    FindRow C_Sheet, "A", Ctl_Beg, "<REC_BEG>"
    FindRow C_Sheet, "A", Ctl_End, "<REC_END>"
    ' Format WTB_Sheet $ to White
    For I = Ctl_Beg To Ctl_End
        FindRow WTB_Sheet, "A", Tmp1_I, Worksheets(C_Sheet).Range(CTL_Col(1) & I).Value
        If Tmp1_I > 0 Then Worksheets(WTB_Sheet).Range(WTB_Col & Tmp1_I).Interior.Color = RGB(255, 255, 255)
    Next I
    Test_Row = 0
    For Ctl_Row = Ctl_Beg To Ctl_End
        Test_Row = Test_Row + 1
        Find_Ctl = Worksheets(C_Sheet).Range(CTL_Col(1) & Ctl_Row).Value
        Worksheets(T_Sheet).Range("A" & Test_Row).Value = Find_Ctl
        FindRow WTB_Sheet, "A", WTB_Row, Find_Ctl
        If WTB_Row > 0 Then
            'Debug.Print ">" & WTB_Sheet & "<Found>" & Find_Ctl & "<"
            Find_Lookup = Trim(Worksheets(C_Sheet).Range(CTL_Col(4) & Ctl_Row).Value)
            'Worksheets(WTB_Sheet).Range(WTB_Col & WTB_Row).Interior.Color = RGB(255, 255, 255)
            Worksheets(T_Sheet).Range("B" & Test_Row).Value = Worksheets(WTB_Sheet).Range(WTB_Col & WTB_Row).Value
            If Worksheets(C_Sheet).Range(CTL_Col(5) & Ctl_Row).Value = Use_BS Then
                Lookup_Sheet = BS_Sheet
                Lookup_DolCol = BSdol_Col
                Lookup_Cnt = BS_Cnt
                Lookup_Flag = "BS"
                Lookup_BegRow = BS_BegRow
                Lookup_EndRow = BS_EndRow
            Else
                Lookup_Sheet = PL_Sheet
                Lookup_DolCol = PLdol_Col
                Lookup_Cnt = PL_Cnt
                Lookup_Flag = "PL"
                Lookup_BegRow = PL_BegRow
                Lookup_EndRow = PL_EndRow
            End If  ' Lookup_Sheet
        Lookup_Exit = 0    ' Not Found Yet
        For I = 1 To Lookup_Cnt
            If Lookup_Flag = "BS" Then
                Lookup_Col = BS_Col(I)
            Else
                Lookup_Col = PL_Col(I)
            End If
            Debug.Print "Search>" & Lookup_Sheet & "<for>" & Find_Lookup & "<"
            Tmp_EndRow = Lookup_EndRow
            Do  ' Loop until Exit
            Debug.Print "<In Col>" & Lookup_Col & "<Lookup_BegRow>" & Lookup_BegRow & "<Lookup_EndRow>" & Lookup_EndRow & "<"
            FindRowReverse Lookup_Sheet, Lookup_Col, Lookup_BegRow, Tmp_EndRow, Lookup_Row, Find_Lookup
            Debug.Print "Found Lookup_Row>" & Lookup_Row & "<"
            If Lookup_Row > 0 Then
                If Find_Lookup = Trim(Worksheets(Lookup_Sheet).Range(Lookup_Col & Lookup_Row).Value) Then
                    Lookup_Exit = 1
                    Exit Do
                    ' Exit Next
                Else
                    Tmp_EndRow = Lookup_Row - 1
                    Debug.Print "Loop Again for>" & Find_Lookup & "<No Match>" & Trim(Worksheets(Lookup_Sheet).Range(Lookup_Col & Lookup_Row).Value)
                    Debug.Print "Lookup_Row>" & Lookup_Row & "<Lookup_EndRow>" & Lookup_EndRow & "<"
                End If ' Find_Lookup
            Else
                ' NOF
                Exit Do
            End If ' Lookup_Row
            If Lookup_EndRow <= Lookup_BegRow Then
                Exit Do
            End If ' Exit Loop
            Loop
            If Lookup_Exit = 1 And Lookup_Row > 0 Then
                Exit For
            End If ' Found and Exit Next I
        Next I
        Debug.Print "Find_Lookup>" & Find_Lookup & "<Lookup_Exit>" & Lookup_Exit & "<Found At>" & Lookup_Col & Lookup_Row & "<"
        If Lookup_Exit = 1 And Lookup_Row > 0 Then
            Debug.Print "WTB>" & Worksheets(WTB_Sheet).Range(WTB_Col & WTB_Row).Value & "<"
            Debug.Print ">" & Lookup_Sheet & "<>" & Worksheets(Lookup_Sheet).Range(Lookup_Col & Lookup_Row).Value & "<"
            If Abs(Round(Worksheets(WTB_Sheet).Range(WTB_Col & WTB_Row).Value, 2)) = Abs(Round(Worksheets(Lookup_Sheet).Range(Lookup_DolCol & Lookup_Row).Value, 2)) Then
                ' Format Green
                Worksheets(WTB_Sheet).Range(WTB_Col & WTB_Row).Interior.Color = RGB(198, 224, 180)
                Worksheets(Lookup_Sheet).Range(Lookup_DolCol & Lookup_Row).Interior.Color = RGB(198, 224, 180)
            Else
                ' Format Pink
                Worksheets(WTB_Sheet).Range(WTB_Col & WTB_Row).Interior.Color = RGB(255, 197, 197)
                Worksheets(Lookup_Sheet).Range(Lookup_DolCol & Lookup_Row).Interior.Color = RGB(255, 197, 197)
            End If ' Color Format
        End If ' Color Format
        Else
            Debug.Print ">" & WTB_Sheet & "<NOF>" & Find_Ctl & "<"
        End If ' WTB Found
    Next Ctl_Row
    
ExitRoutine:
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    Exit Function
    
ErrSub:
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
    GoTo ExitRoutine
    
    End Function
    
    Function WTB_Subtotal_Del()
    Const VBA_Name As String = "WTB_Subtotal_Del"
    On Error GoTo ErrSub
    Debug.Print "Open >" & VBA_Name & "<"
    Dim I, WTB_Beg, WTB_End As Integer
    Dim WTB_Sheet, Tmp1_S, Tmp_Col As String
    
    Const Use_Sheet As String = "WTB_01"
    Const Notes_Sheet As String = "WTB_NOTES"
    
    ActiveSheet.Unprotect

    WTB_Sheet = "NOF"
    For Each WkSheet In ThisWorkbook.Worksheets
        If WkSheet.CodeName = Use_Sheet Then
            WTB_Sheet = WkSheet.Name
            Exit For
        End If
    Next WkSheet
    
    If WTB_Sheet <> "NOF" Then
        ' Fall thru & Execute
    Else
        Tmp1_I = MsgBox(VBA_Name & " Can NOT find a WTB Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "WTB Worksheet Not Found")
        GoTo ExitRoutine
    End If
    
    FindColNumLtr WTB_Sheet, 1, I, Tmp_Col, "<END_DEL>"
    
    FindColNumLtr WTB_Sheet, 1, WTB_Beg, Tmp1_S, "<ACCT>"
    FindRow WTB_Sheet, "A", WTB_Beg, "<TOT_SUB><ADJUSTMENTS>"
    FindLastRow WTB_Sheet, WTB_End
    FindColNumLtr WTB_Sheet, 1, Tmp2_I, Tmp2_S, "<NOTES>"
    Tmp2_I = Tmp2_I + 1
    NumToLtr Tmp2_I, Tmp2_S
    If WTB_Beg > 0 And WTB_End > WTB_Beg Then
        FindLastRow Notes_Sheet, I
        If I > 0 Then
            Worksheets(Notes_Sheet).Range("A1:A" & I).EntireRow.Delete
        End If  ' Purge Notes_Sheet
        Worksheets(WTB_Sheet).Range(Tmp1_S & (WTB_Beg + 1) & ":" & Tmp2_S & WTB_End).Copy
        Worksheets(Notes_Sheet).Range("A1").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
        Worksheets(WTB_Sheet).Range("A" & (WTB_Beg + 1) & ":A" & WTB_End).EntireRow.Delete
    End If  ' Save Notes
    FindRow WTB_Sheet, "A", WTB_Beg, "<HDR>"
    WTB_Beg = WTB_Beg + 1
    FindLastRow WTB_Sheet, WTB_End
    
    For I = WTB_End To WTB_Beg Step -1
        If Left(Worksheets(WTB_Sheet).Range("A" & I).Value, 4) = "<TOT" Then
            'Debug.Print ">" & I & "<>" & Worksheets(WTB_Sheet).Range("A" & I).Value & "<"
            Worksheets(WTB_Sheet).Range("A" & I & ":A" & I).EntireRow.Delete
        End If
    Next I
    
    Worksheets(WTB_Sheet).Range("A2:A12").EntireRow.Hidden = False
    
ExitRoutine:
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    Exit Function
    
ErrSub:
    Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    GoTo ExitRoutine
    
    End Function
    
    Function WTB_Subtotal_Refresh()
    Const VBA_Name As String = "WTB_Subtotal_Refresh"
    On Error GoTo ErrSub
    Debug.Print "Open >" & VBA_Name & "<"
    Dim I, Row_Beg, Row_Beg01, Row_Beg02, Row_Beg03, Row_End, Row_End01, Row_End02, Row_End03, SubTot_Beg, SubTot_End, Tmp1_I As Integer
    Dim Row_Cogs, Row_GrossProfit, Tmp_Col, Tmp1_Row, Tmp2_Row, Tmp3_Row As String
    Dim WTB_Sheet, Str_1, Str_2, Tmp1_S As String
    
    Const Col_Num As Integer = 8
    Const Use_Sheet As String = "WTB_01"
    Const Notes_Sheet As String = "WTB_NOTES"
    Const Find_TotIncome As String = "<TOT_SUB><INCOME>"
    Const Find_TotCogs As String = "<TOT_SUB><COST OF GOODS SOLD>"
    Const Find_TotGrossProfit As String = "<TOT_SUB><GROSS PROFIT>"
    Const Find_TotNetIncomeLoss As String = "<TOT_SUB><NET_INCOME_LOSS>"
    
    ActiveSheet.Unprotect

    
    WTB_Sheet = "NOF"
    For Each WkSheet In ThisWorkbook.Worksheets
        If WkSheet.CodeName = Use_Sheet Then
            WTB_Sheet = WkSheet.Name
            Exit For
        End If
    Next WkSheet
    
    If WTB_Sheet <> "NOF" Then
        ' Fall thru & Execute
    Else
        Tmp1_I = MsgBox(VBA_Name & " Can NOT find a WTB Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "WTB Worksheet Not Found")
        GoTo ExitRoutine
    End If
    
    
    FindColNumLtr WTB_Sheet, 1, I, Tmp_Col, "<END_DEL>"
    
    ' Delete Subtotals to be safe
    WTB_Subtotal_Del
    ' Begin actual REFRESH
    Dim Col_Find_Subtotal(Col_Num)
    Dim Col_Find_TB(Col_Num)
    
    Dim Col_Ltr_Subtotal(Col_Num)
    Dim Col_Ltr_TB(Col_Num)
    
    
    
    Col_Find_TB(1) = "<DESC>"
    Col_Find_TB(2) = "<BOOK>"
    Col_Find_TB(3) = "<SUB_TOT>"
    Col_Find_TB(4) = "<DR>"
    Col_Find_TB(5) = "<FIND>"
    Col_Find_TB(6) = "<CR>"
    Col_Find_TB(7) = "<FINAL>"
    Col_Find_TB(8) = "<ACCT>"
    
    Col_Find_Subtotal(1) = "<WTB_SUB_TOT>"
    Col_Find_Subtotal(2) = "<WTB_INC>"
    Col_Find_Subtotal(3) = "<WTB_BEG>"
    Col_Find_Subtotal(4) = "<DR>"
    Col_Find_Subtotal(5) = "<WTB_END>"
    Col_Find_Subtotal(6) = "<FIND>"
    Col_Find_Subtotal(7) = "<FINAL>"
    Col_Find_Subtotal(8) = "<ACCT>"
    
    SubTot_Beg = 5
    SubTot_End = 11
    
    
    ActiveSheet.Unprotect
    For I = 1 To Col_Num
        FindColNumLtr WTB_Sheet, 1, Tmp1_I, Col_Ltr_TB(I), Col_Find_TB(I)
        FindColNumLtr WTB_Sheet, 1, Tmp1_I, Col_Ltr_Subtotal(I), Col_Find_Subtotal(I)
    Next I
    
    
    For I = SubTot_End To SubTot_Beg Step -1
        If Worksheets(WTB_Sheet).Range(Col_Ltr_Subtotal(2) & I).Value = "SUBTOTAL" Then
            Sub_Beg = Worksheets(WTB_Sheet).Range(Col_Ltr_Subtotal(3) & I).Value
            Sub_End = Worksheets(WTB_Sheet).Range(Col_Ltr_Subtotal(5) & I).Value
            ActiveSheet.Range("A" & (Sub_End + 1) & ":A" & (Sub_End + 1)).EntireRow.Insert
            Worksheets(WTB_Sheet).Range("A" & Sub_End + 1).Value = "<TOT_BLANK>"
            Worksheets(WTB_Sheet).Range("A" & Sub_End + 1).RowHeight = 10
            ActiveSheet.Range("A" & (Sub_End + 2) & ":A" & (Sub_End + 2)).EntireRow.Insert
            Worksheets(WTB_Sheet).Range("A" & Sub_End + 2).Value = "<TOT_SUB><" & UCase(Worksheets(WTB_Sheet).Range(Col_Ltr_Subtotal(1) & I).Value) & ">"
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & Sub_End + 2).Value = "TOTAL " & UCase(Worksheets(WTB_Sheet).Range(Col_Ltr_Subtotal(1) & I).Value)
            ' Insert Totals
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & Sub_End + 2).Formula = "=Subtotal(9," & Col_Ltr_TB(2) & Sub_Beg & ":" & Col_Ltr_TB(2) & (Sub_End + 1) & ")"
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & Sub_End + 2).Formula = "=Subtotal(9," & Col_Ltr_TB(7) & Sub_Beg & ":" & Col_Ltr_TB(7) & (Sub_End + 1) & ")"
            With Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & (Sub_End + 2) & ":" & Col_Ltr_TB(7) & (Sub_End + 2))
                .Font.Bold = True
                .Font.Size = 16
            End With
            Worksheets(WTB_Sheet).Range("A" & Sub_End + 2).RowHeight = 20
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & (Sub_End + 2)).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & (Sub_End + 2)).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & (Sub_End + 1)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & (Sub_End + 1)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
            
            ActiveSheet.Range("A" & (Sub_End + 3) & ":A" & (Sub_End + 3)).EntireRow.Insert
            Worksheets(WTB_Sheet).Range("A" & Sub_End + 3).Value = "<TOT_BLANK>"
            Worksheets(WTB_Sheet).Range("A" & Sub_End + 3).RowHeight = 10
        End If
    Next I
    ' Gross Profit - If COGS
    FindRow WTB_Sheet, "A", Row_Cogs, Find_TotCogs
    If Row_Cogs > 0 Then
        For I = 1 To 2
        Worksheets(WTB_Sheet).Range("A" & (Row_Cogs + 2) & ":A" & (Row_Cogs + 2)).EntireRow.Insert
        If I = 2 Then
            Worksheets(WTB_Sheet).Range("A" & (Row_Cogs + 2)).Value = Find_TotGrossProfit
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & (Row_Cogs + 2)).Value = "GROSS PROFIT"
            Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & (Row_Cogs + 2)).HorizontalAlignment = xlRight
            Worksheets(WTB_Sheet).Rows((Row_Cogs + 2)).RowHeight = 20
            FindRow WTB_Sheet, "A", Tmp1_Row, Find_TotIncome
            FindRow WTB_Sheet, "A", Tmp2_Row, Find_TotCogs
            FindRow WTB_Sheet, "A", Row_GrossProfit, Find_TotGrossProfit
            Str_1 = "=" & Col_Ltr_TB(2) & Tmp1_Row & "+" & Col_Ltr_TB(2) & Tmp2_Row
            Str_2 = "=" & Col_Ltr_TB(7) & Tmp1_Row & "+" & Col_Ltr_TB(7) & Tmp2_Row
    
            With Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & Row_GrossProfit)
                .Formula = Str_1
                .Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            End With
            With Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & Row_GrossProfit)
                .Formula = Str_2
                .Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
            End With
        Else
            Worksheets(WTB_Sheet).Range("A" & (Row_Cogs + 2)).Value = "<TOT_BLANK>"
        End If  ' Insert Gross Profit Rows
        Next I
    Else
        Debug.Print "!!!! No COGS = No GROSS PROFIT"
    End If  ' Row_Cogs
    ' Subtotal Liabilities & Equity
    Debug.Print "Subtotal LIabilities and Equity"
    Row_Beg = 0
    Row_End = 0
    FindRow WTB_Sheet, "A", Row_Beg01, "<LIABILITIES>"
    FindRow WTB_Sheet, "A", Row_Beg02, "<EQUITY>"
    If Row_Beg01 <= Row_Beg02 Then
        Row_Beg = Row_Beg01
    Else
        If Row_Beg02 < Row_Beg01 Then
            Row_Beg = Row_Beg02
        End If
    End If  ' Row_Beg01 <= Row_Beg02
    FindRow WTB_Sheet, "A", Row_End01, "<TOT_SUB><LIABILITIES>"
    FindRow WTB_Sheet, "A", Row_End02, "<TOT_SUB><EQUITY>"
    If Row_End01 >= Row_End02 Then
        Row_End = Row_End1
    Else
        If Row_End02 > Row_End01 Then
            Row_End = Row_End02
        End If
    End If  ' Row_End01 >= Row_End02
    If Row_Beg <> 0 And Row_End <> 0 Then
        ActiveSheet.Range("A" & (Row_End + 1) & ":A" & (Row_End + 1)).EntireRow.Insert
        Worksheets(WTB_Sheet).Range("A" & (Row_End + 1)).Value = "<TOT_BLANK>"
        Worksheets(WTB_Sheet).Range("A" & (Row_End + 1)).RowHeight = 10
        ActiveSheet.Range("A" & (Row_End + 2) & ":A" & (Row_End + 2)).EntireRow.Insert
        Worksheets(WTB_Sheet).Range("A" & (Row_End + 2)).RowHeight = 20
        Worksheets(WTB_Sheet).Range("A" & (Row_End + 2)).Value = "<TOT_SUB><LIABILITIES_EQUITY>"
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & (Row_End + 2)).Value = "TOTAL LIABILITIES AND EQUITY"
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & Row_End + 2).Formula = "=Subtotal(9," & Col_Ltr_TB(2) & Row_Beg & ":" & Col_Ltr_TB(2) & Row_End & ")"
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & Row_End + 2).Formula = "=Subtotal(9," & Col_Ltr_TB(7) & Row_Beg & ":" & Col_Ltr_TB(7) & Row_End & ")"
        With Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & (Row_End + 2) & ":" & Col_Ltr_TB(7) & (Row_End + 2))
            .Font.Bold = True
        End With
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & (Row_End + 2)).HorizontalAlignment = xlRight
        
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & (Row_End + 2)).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & (Row_End + 2)).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    End If  ' LIABILITES AND EQUITY
    ' Net Income Loss
    Debug.Print "Net Income Loss"
    FindLastRow WTB_Sheet, Tmp1_I
    Tmp1_I = Tmp1_I + 1
    Str_1 = ""
    Str_2 = ""
    Tmp1_I = 0
    FindRow WTB_Sheet, "A", I, Find_TotGrossProfit
    If I > 0 Then
        ' Use Total Gross Profits
    Else
        ' Use Total Income
        FindRow WTB_Sheet, "A", I, "<TOT_SUB><INCOME>"
    End If
    If I > 0 Then   ' If Gross Profits or Income
    Tmp1_I = Tmp1_I + 1
    Str1 = Str1 & Col_Ltr_TB(2) & I
    Str2 = Str2 & Col_Ltr_TB(7) & I
    End If  ' Income
    Debug.Print "Income Str1>" & Str1 & "<"
    FindRow WTB_Sheet, "A", I, "<TOT_SUB><EXPENSES>"
    If I > 0 Then
        If Tmp1_I > 0 Then
            Str1 = Str1 & "+"
            Str2 = Str2 & "+"
        End If
        Tmp1_I = Tmp1_I + 1
        Str1 = Str1 & Col_Ltr_TB(2) & I
        Str2 = Str2 & Col_Ltr_TB(7) & I
    End If  ' Expense
    Debug.Print "Expense Str1>" & Str1 & "<"
    FindRow WTB_Sheet, "A", I, "<TOT_SUB><NET OTHER (INCOME)/EXPENSE>"
    If I > 0 Then
        If Tmp1_I > 0 Then
            Str1 = Str1 & "+"
            Str2 = Str2 & "+"
        End If
        Tmp1_I = Tmp1_I + 1
        Str1 = Str1 & Col_Ltr_TB(2) & I
        Str2 = Str2 & Col_Ltr_TB(7) & I
    End If  ' Income
    Debug.Print "Other Income/Expense>" & Str1 & "<"
    Str1 = "=" & Str1
    Str2 = "=" & Str2
    FindLastRow WTB_Sheet, Tmp1_I
    Tmp1_I = Tmp1_I + 1
    Worksheets(WTB_Sheet).Range("A" & Tmp1_I).Value = Find_TotNetIncomeLoss
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & Tmp1_I).Value = "NET (INCOME)/LOSS"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & Tmp1_I).Formula = Str1
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & Tmp1_I).Formula = Str2
    With Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & Tmp1_I & ":" & Col_Ltr_TB(7) & Tmp1_I)
        .Font.Bold = True
    End With
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(4) & (Tmp1_I - 1)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(6) & (Tmp1_I - 1)).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(4) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(6) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
    ' Current Year Earnings
    FindRow WTB_Sheet, "A", I, "<TOT_SUB><EQUITY>"
    I = I - 1
    ActiveSheet.Range("A" & I & ":A" & I).EntireRow.Insert
    Tmp1_I = Tmp1_I + 1 ' compensate INSERT
    Worksheets(WTB_Sheet).Range("A" & I).Value = "<TOT><C_Y_E>"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & I).Value = "Current Year Earnings"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(8) & I).Value = "EQUITY"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & I).Formula = "=" & Col_Ltr_TB(2) & Tmp1_I
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & I).Formula = "=" & Col_Ltr_TB(7) & Tmp1_I
    ' Format 'Totals'
    FindRow WTB_Sheet, "A", Tmp1_I, "<TOT_SUB><ASSETS>"
    If Tmp1_I > 0 Then
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & Tmp1_I).HorizontalAlignment = xlRight
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = xlDouble
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(7) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = xlDouble
    End If
    FindRow WTB_Sheet, "A", Tmp1_I, "<TOT_SUB><LIABILITES_EQUITY>"
    If Tmp1_I > 0 Then
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(2) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = xlDouble
    End If
    FindRow WTB_Sheet, "A", Tmp1_I, "<TOT_SUB><NET OTHER (INCOME) AND EXPENSE>"
    If Tmp1_I > 0 Then
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(4) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(6) & Tmp1_I).Borders(xlEdgeBottom).LineStyle = xlContinuous
    End If
    ' Dr & Cr Totals
    FindRow WTB_Sheet, "A", Row_Beg01, "<HDR>"
    Row_Beg01 = Row_Beg01 + 1
    FindRow WTB_Sheet, "A", Row_End01, Find_TotNetIncomeLoss
    Str1 = "=Sum(" & Col_Ltr_TB(4) & Row_Beg01 & ":" & Col_Ltr_TB(4) & (Row_End01 - 2) & ")"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(4) & Row_End01).Formula = Str1
    Str2 = "=Sum(" & Col_Ltr_TB(6) & Row_Beg01 & ":" & Col_Ltr_TB(6) & (Row_End01 - 2) & ")"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(6) & Row_End01).Formula = Str2
    Worksheets(WTB_Sheet).Range("A" & Row_End01).RowHeight = 20
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & Row_End01 & ":" & Col_Ltr_TB(7) & Row_End01).Font.Size = 16
    FindLastRow WTB_Sheet, Row_End02
    Row_End02 = Row_End02 + 2
    Worksheets(WTB_Sheet).Range("A" & (Row_End02)).Value = "<TOT_SUB><ADJUSTMENTS>"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(8) & (Row_End02)).Value = "AJE ID"
    Worksheets(WTB_Sheet).Range(Col_Ltr_TB(1) & (Row_End02)).Value = "AJE DESCRIPTION"
    With Worksheets(WTB_Sheet).Range(Col_Ltr_TB(8) & Row_End02 & ":" & Col_Ltr_TB(1) & (Row_End02))
        .Font.Bold = True
        .Font.Size = 16
        .RowHeight = 20
    End With
    FindLastColumn WTB_Sheet, Tmp1_I
    NumToLtr Tmp1_I, Tmp1_S
    With Worksheets(WTB_Sheet).Range(Col_Ltr_TB(8) & (Row_End02) & ":" & Tmp1_S & (Row_End02))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    FindLastRow Notes_Sheet, Row_End01
    FindLastColumn Notes_Sheet, Tmp1_I
    NumToLtr Tmp1_I, Tmp1_S
    NumToLtr Tmp1_I, Tmp1_S
    If Row_End01 > 0 Then
        Worksheets(Notes_Sheet).Range("A1" & ":" & Tmp1_S & Row_End01).Copy
        Worksheets(WTB_Sheet).Range(Col_Ltr_TB(8) & (Row_End02 + 1)).PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    End If  ' Copy Back Notes
    
    WTB_Reconcile
    
    ' JOHN Macro
    FindLastRow WTB_Sheet, Row_End01
     Range("N14").Select
        ActiveCell.Formula2R1C1 = _
            "=IF(ISERROR(FIND(""="",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),1)),"""",SUMIF(C13,SUBSTITUTE(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),""="",""""),C12) + OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-2))"
        Range("N14").Select
        Selection.Copy
        Range("N15:N" & Row_End01).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
    ' End JOHN Macro
    
    Worksheets(WTB_Sheet).Range("A" & SubTot_Beg - 1 & ":A" & SubTot_End).EntireRow.Hidden = True
    Worksheets(WTB_Sheet).Range("A" & SubTot_Beg - 2).RowHeight = 50
    Worksheets(WTB_Sheet).Range("A" & SubTot_End + 1).RowHeight = 30
    
ExitRoutine:
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    Exit Function
    
ErrSub:
    Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    GoTo ExitRoutine
    
    End Function
    
    Sub insertRow()
        ActiveSheet.Unprotect
        Dim subtotalDescr, Tmp_Col As String

        FindColNumLtr WTB_Sheet, 1, I, Tmp_Col, "<END_DEL>"

        subtotalDescr = ActiveSheet.Range("D" & ActiveCell.Row).Value
        If subtotalDescr = "" Then subtotalDescr = ActiveSheet.Range("D" & ActiveCell.Row - 1).Value
        ActiveSheet.Range("A" & ActiveCell.Row & ":A" & Activecell.row).EntireRow.Insert
        ActiveSheet.Range("N" & ActiveCell.Row).Formula2 = "=IF(ISERROR(FIND(""="",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),1)),"""",SUMIF($M:$M,SUBSTITUTE(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),""="",""""),$L:$L) + OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-2))"
        ActiveSheet.Range("D" & ActiveCell.Row).Value = subtotalDescr
        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    End Sub
    
    Sub deleteRow()
        Dim Tmp_Col As String
        FindColNumLtr WTB_Sheet, 1, I, Tmp_Col, "<END_DEL>"

        ActiveSheet.Unprotect
        ActiveSheet.Range("A" & ActiveCell.Row & ":A" & Activecell.row).EntireRow.delete
        ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=True _
            , AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True
    End Sub
    

