Attribute VB_Name = "Mod_Export"
'--------------------------------------
' AJE_Export
' AJE_Prep
'--------------------------------------
Function AJE_Prep()
' Make sure that AJE Insert Rows have an Account Desc
Const VBA_Name As String = "AJE_Prep"
On Error GoTo ErrSub
Dim WTB_Sheet, WTB_Col01, WTB_Col02, WTB_Col03, Tmp1_S As String
Dim I, WTB_Beg, WTB_End, WTB_Row, Tmp1_I

Const Use_WTB As String = "WTB_01"
Const Find_Col_01 As String = "<DESC>"
Const Find_Col_02 As String = "<SUB_TOT>"
Const Find_Col_03 As String = "<FIND>"

WTB_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
Next WkSheet

If WTB_Sheet <> "NOF" Then
    ' OK - Fall Thru
Else
    Msg_01 = VBA_Name & Chr(13) & Chr(10)
    If WTB_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [WTB - Working Trial Balance Sheet]" & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & "Please make a note of this message and contact Program Development"
    Tmp1_I = MsgBox(Msg_01, vbExclamation, "Worksheet Not Found")
    GoTo ExitRoutine
End If

FindColNumLtr WTB_Sheet, 1, Tmp1_I, WTB_Col01, Find_Col_01
FindColNumLtr WTB_Sheet, 1, Tmp1_I, WTB_Col02, Find_Col_02
FindColNumLtr WTB_Sheet, 1, Tmp1_I, WTB_Col03, Find_Col_03
FindRow WTB_Sheet, "A", WTB_Beg, "<HDR>"
WTB_Beg = WTB_Beg + 1
FindLastRow WTB_Sheet, WTB_End
With Worksheets(WTB_Sheet)
For WTB_Row = WTB_Beg To WTB_End
    If Trim(.Range(WTB_Col02 & WTB_Row).Value) = "" And Trim(.Range(WTB_Col03 & WTB_Row).Value) = "" Then
        ' No AJE Flag = Skip
    Else
        ' AJE Flag = Check Desc
        If Trim(.Range(WTB_Col01 & WTB_Row).Value) = "" Then
            .Range(WTB_Col01 & WTB_Row).Value = .Range(WTB_Col01 & (WTB_Row - 1)).Value
            '.Range(WTB_Col01 & WTB_Row).Font.Color = RGB(255, 0, 0) ' Set to RED
            .Range(WTB_Col01 & WTB_Row).Font.Color = RGB(255, 255, 255) ' Set to WHITE
        End If
    End If  ' AJE Yes/No
Next WTB_Row
End With

ExitRoutine:
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine
End Function

Function AJE_Export()
Const VBA_Name As String = "Export_AJE"
'On Error GoTo ErrSub
Dim Ctl_Sheet, AJE_Sheet, WTB_Sheet, Tmp1_S, Msg_01 As String
Dim I, Tmp1_I, Tmp2_I, Row_Beg, Row_End, Row_Exp As Integer
Dim Disp_YrEnd As Date
Dim WTB_Col()
Dim AJE_Col()
Dim CTL_Col(3)

Const Use_WTB As String = "WTB_01"
Const Use_AJE As String = "AJE_01"
Const Use_CTL As String = "CTL_01"

ActiveSheet.Unprotect
This_Sheet = ActiveSheet.Name
Disp_YrEnd = Application.Evaluate("Yr_End")
AJE_Sheet = "NOF"
WTB_Sheet = "NOF"
Ctl_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_AJE Then AJE_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_CTL Then Ctl_Sheet = WkSheet.Name
Next WkSheet

If WTB_Sheet <> "NOF" And AJE_Sheet <> "NOF" And Ctl_Sheet <> "NOF" Then
    ' OK - Fall Thru
Else
    Msg_01 = VBA_Name & Chr(13) & Chr(10)
    If Ctl_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [CONTROL Sheet]" & Chr(13) & Chr(10)
    If AJE_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [AJE export Sheet]" & Chr(13) & Chr(10)
    If WTB_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [WTB - Working Trial Balance Sheet]" & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & "Please make a note of this message and contact Program Development"
    Tmp1_I = MsgBox(Msg_01, vbExclamation, "Worksheet Not Found")
    GoTo ExitRoutine
End If
' Prep WTB_sheet Account Desc's
AJE_Prep
' Purge AJE_Sheet
FindRow AJE_Sheet, "A", Row_Beg, "<HDR>"
Row_Exp = Row_Beg
Row_Beg = Row_Beg + 1
FindLastRow AJE_Sheet, Row_End
Row_End = Row_End + 3
If Row_End >= Row_Beg Then
    ' Purge = Yes
    Tmp1_S = "A" & Row_Beg & ":" & "A" & Row_End
    Worksheets(AJE_Sheet).Range(Tmp1_S).EntireRow.Delete
End If
' Build Arrays
'CTL_Col() Array
For I = 1 To 3
    Tmp1_S = "<COL_" & Right((100 + I), 2) & ">"
    FindColNumLtr Ctl_Sheet, 1, Tmp1_I, CTL_Col(I), Tmp1_S
Next I
' WTB_Col() Array
FindRow Ctl_Sheet, "A", Row_Beg, "<WTB_BEG>"
Tmp1_I = Worksheets(Ctl_Sheet).Range(CTL_Col(2) & (Row_Beg - 1)).Value
ReDim WTB_Col(Tmp1_I)
FindRow Ctl_Sheet, "A", Row_End, "<WTB_END>"
I = 0
For Tmp1_I = Row_Beg To Row_End
    I = I + 1
    Tmp2_S = Worksheets(Ctl_Sheet).Range(CTL_Col(3) & Tmp1_I).Value
    FindColNumLtr WTB_Sheet, 1, Tmp2_I, WTB_Col(I), Tmp2_S
Next Tmp1_I
' AJE_Col() Array
FindRow Ctl_Sheet, "A", Row_Beg, "<AJE_BEG>"
Tmp1_I = Worksheets(Ctl_Sheet).Range(CTL_Col(2) & (Row_Beg - 1)).Value
ReDim AJE_Col(Tmp1_I)
FindRow Ctl_Sheet, "A", Row_End, "<AJE_END>"
I = 0
For Tmp1_I = Row_Beg To Row_End
    I = I + 1
    Tmp2_S = Worksheets(Ctl_Sheet).Range(CTL_Col(3) & Tmp1_I).Value
    FindColNumLtr AJE_Sheet, 1, Tmp2_I, AJE_Col(I), Tmp2_S
Next Tmp1_I
' Append AJE Hdr's
FindRow WTB_Sheet, "A", Row_Beg, "<ADJUSTMENTS>"
Row_Beg = Row_Beg + 1
FindLastRowInCol WTB_Sheet, WTB_Col(1), Row_End
For I = Row_Beg To Row_End
    Row_Exp = Row_Exp + 1
    Tmp1_S = "<DTL><" & Worksheets(WTB_Sheet).Range(WTB_Col(1) & I).Value & "><1Hdr>"
    Worksheets(AJE_Sheet).Range("A" & Row_Exp).Value = Tmp1_S
    Worksheets(AJE_Sheet).Range(AJE_Col(1) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(1) & I).Value
    Worksheets(AJE_Sheet).Range(AJE_Col(2) & Row_Exp).Value = Disp_YrEnd
    Worksheets(AJE_Sheet).Range(AJE_Col(3) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(2) & I).Value
    Worksheets(AJE_Sheet).Range(AJE_Col(3) & Row_Exp).Font.Bold = True
    Worksheets(AJE_Sheet).Range(AJE_Col(3) & Row_Exp).Font.Italic = True
    Row_Exp = Row_Exp + 1
    Tmp1_S = "<DTL><" & Worksheets(WTB_Sheet).Range(WTB_Col(1) & I).Value & "><9FTR>"
    Worksheets(AJE_Sheet).Range("A" & Row_Exp).Value = Tmp1_S
Next I
' Append AJE DR's
FindRow WTB_Sheet, "A", Row_Beg, "<HDR>"
Row_Beg = Row_Beg + 1
FindLastRowInCol WTB_Sheet, WTB_Col(3), Row_End
For I = Row_Beg To Row_End
    If Worksheets(WTB_Sheet).Range(WTB_Col(3) & I).Value > "" Then
        Row_Exp = Row_Exp + 1
        Tmp1_S = "<DTL><" & Worksheets(WTB_Sheet).Range(WTB_Col(3) & I).Value & ">" & "<2Dr>"
        Worksheets(AJE_Sheet).Range("A" & Row_Exp).Value = Tmp1_S
        'Worksheets(AJE_Sheet).Range(AJE_Col(1) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(3) & I).Value
        Worksheets(AJE_Sheet).Range(AJE_Col(3) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(2) & I).Value
        Worksheets(AJE_Sheet).Range(AJE_Col(4) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(4) & I).Value
    End If  ' Flag = Yes
Next I
' Append AJE Cr's
FindLastRowInCol WTB_Sheet, WTB_Col(5), Row_End
For I = Row_Beg To Row_End
    If Worksheets(WTB_Sheet).Range(WTB_Col(5) & I).Value > "" Then
        Row_Exp = Row_Exp + 1
        Tmp1_S = "<DTL><" & Worksheets(WTB_Sheet).Range(WTB_Col(5) & I).Value & "><3Cr>"
        Worksheets(AJE_Sheet).Range("A" & Row_Exp).Value = Tmp1_S
        'Worksheets(AJE_Sheet).Range(AJE_Col(1) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(5) & I).Value
        Worksheets(AJE_Sheet).Range(AJE_Col(3) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(2) & I).Value
        Worksheets(AJE_Sheet).Range(AJE_Col(5) & Row_Exp).Value = Worksheets(WTB_Sheet).Range(WTB_Col(6) & I).Value
    End If  ' Flag = Yes
Next I
' Sort
FindRow AJE_Sheet, "A", Row_Beg, "<HDR>"
Row_Beg = Row_Beg + 1
FindLastRow AJE_Sheet, Row_End
SortArea AJE_Sheet, "A", Row_Beg, AJE_Col(5), Row_End, "A", AJE_Col(1)
' AJE Totals
FindRow WTB_Sheet, "A", I, "<NET_INCOME_LOSS>"
Tmp1_S = "=Sum(" & AJE_Col(4) & Row_Beg & ":" & AJE_Col(4) & Row_End & ")"
Worksheets(AJE_Sheet).Range(AJE_Col(4) & (Row_End + 2)).Formula = Tmp1_S
Tmp1_S = "=Sum(" & AJE_Col(5) & Row_Beg & ":" & AJE_Col(5) & Row_End & ")"
Worksheets(AJE_Sheet).Range(AJE_Col(5) & (Row_End + 2)).Formula = Tmp1_S
' WTB to AJE Reconcile = Format Background
If Abs(Round(Worksheets(WTB_Sheet).Range(WTB_Col(4) & I).Value, 2)) = Abs(Round(Worksheets(AJE_Sheet).Range(AJE_Col(4) & (Row_End + 2)).Value, 2)) Then
    Worksheets(WTB_Sheet).Range(WTB_Col(4) & I).Interior.Color = RGB(198, 224, 180)
    Worksheets(AJE_Sheet).Range(AJE_Col(4) & (Row_End + 2)).Interior.Color = RGB(198, 224, 180)
Else
    Worksheets(WTB_Sheet).Range(WTB_Col(4) & I).Interior.Color = RGB(255, 197, 197)
    Worksheets(AJE_Sheet).Range(AJE_Col(4) & (Row_End + 2)).Interior.Color = RGB(255, 197, 197)
End If  ' Reconcile Dr Col
If Abs(Round(Worksheets(WTB_Sheet).Range(WTB_Col(4) & I).Value, 2)) = Abs(Round(Worksheets(AJE_Sheet).Range(AJE_Col(5) & (Row_End + 2)).Value, 2)) Then
    Worksheets(WTB_Sheet).Range(WTB_Col(6) & I).Interior.Color = RGB(198, 224, 180)
    Worksheets(AJE_Sheet).Range(AJE_Col(5) & (Row_End + 2)).Interior.Color = RGB(198, 224, 180)
Else
    Worksheets(WTB_Sheet).Range(WTB_Col(6) & I).Interior.Color = RGB(255, 197, 197)
    Worksheets(AJE_Sheet).Range(AJE_Col(5) & (Row_End + 2)).Interior.Color = RGB(255, 197, 197)
End If  ' Reconcile Cr Col


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

