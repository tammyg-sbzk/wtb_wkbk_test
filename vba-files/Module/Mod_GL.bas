Attribute VB_Name = "Mod_GL"
'----------
' Rebuild_GL
'----------

Function Rebuild_GL()
Const VBA_Name As String = "Rebuild_GL"
Debug.Print VBA_Name
On Error GoTo ErrSub
Dim GL_Sheet, Use_Sheet, QB_Type, Tmp1_S As String
Dim Out_Row, I, Tmp1_I, Tmp2_I, Tmp3_I, Tmp4_I, Array_Cnt As Integer

Const Use_Raw As String = "Raw_GL"
Const D_Sheet As String = "Dashboard"
Const Ctl_Sheet As String = "Control"
Const Find_Acct As String = "<ACCT>"
Const Find_GL_Desc As String = "<GL_DESC>"
Const Array_CTL As Integer = 4

Dim Col_CTL(Array_CTL) As String
Dim Col_GL() As String
Dim Col_Src() As String
FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Col_CTL(1), "<COL_02>"
FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Col_CTL(2), "<COL_03>"
FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Col_CTL(3), "<COL_04>"
FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Col_CTL(4), "<COL_05>"
FindRow Ctl_Sheet, "A", Tmp1_I, "<GL_DIM>"
Array_Cnt = Worksheets(Ctl_Sheet).Range(Col_CTL(1) & Tmp1_I).Value
ReDim Col_GL(Array_Cnt)
ReDim Col_Src(Array_Cnt)

GL_Sheet = ActiveSheet.Name
Raw_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_Raw Then Raw_Sheet = WkSheet.Name
Next WkSheet
If Raw_Sheet = "NOF" Then
    Tmp1_I = MsgBox(VBA_Name & " Can NOT find a Raw Import Worksheet with the CODE NAME >" & Use_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "RAW Import GL Worksheet Not Found")
    GoTo ExitRoutine
End If
'Tmp1_I = MsgBox("Do you wish to REBUILD " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & ">" & GL_Sheet & "< from the >" & Raw_Sheet & "< download" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "This Will PURGE and Rebuild >" & GL_Sheet & "<", vbQuestion + vbYesNo + vbDefaultButton2, "REBUILD >" & GL_Sheet & "<")
Debug.Print "Tmp1_I>" & Tmp1_I & "<"
'Tmp1_I = 6
'If Tmp1_I = 6 Then
FindRow Ctl_Sheet, "A", Tmp1_I, "<GL_COL_BEG>"
FindRow Ctl_Sheet, "A", Tmp2_I, "<GL_COL_END>"
Tmp4_I = 0
For I = Tmp1_I To Tmp2_I
    If Worksheets(Ctl_Sheet).Range(Col_CTL(1) & I).Value > 0 Then
        Tmp4_I = Tmp4_I + 1
        FindColNumLtr GL_Sheet, 1, Tmp3_I, Col_GL(Tmp4_I), Worksheets(Ctl_Sheet).Range(Col_CTL(2) & I).Value
        FindColNumLtr Raw_Sheet, 1, Tmp3_I, Col_Src(Tmp4_I), Worksheets(Ctl_Sheet).Range(Col_CTL(3) & I).Value
    End If
Next I
' Purge GL_Sheet
FindRow GL_Sheet, "A", Out_Row, "<HDR>"
Out_Row = Out_Row + 1
FindLastRow GL_Sheet, Tmp2_I
If Tmp2_I >= Tmp1_I Then
    Worksheets(GL_Sheet).Range("A" & (Out_Row + 1) & ":A" & Tmp2_I).EntireRow.Delete
End If  ' Purge GL_Sheet
Tmp2_I = 2
FindLastRow Raw_Sheet, Tmp3_I
FindColNumLtr GL_Sheet, 1, Tmp1_I, Tmp1_S, Find_Acct
Worksheets(Raw_Sheet).Range("A" & Tmp2_I & ":" & "A" & Tmp3_I).Copy
Worksheets(GL_Sheet).Range(Tmp1_S & Out_Row).PasteSpecial xlPasteValues
FindColNumLtr GL_Sheet, 1, Tmp1_I, Tmp1_S, Find_GL_Desc
QB_Type = ""
If Trim(Worksheets(Raw_Sheet).Range("B" & 1).Value) > "" Then
    Debug.Print "Not Null = ONLINE"
    QB_Type = "ONLINE"
Else
    Debug.Print "Is Null = LOCAL"
    QB_Type = "LOCAL"
    Worksheets(Raw_Sheet).Range("B" & Tmp2_I & ":" & "B" & Tmp3_I).Copy
    Worksheets(GL_Sheet).Range(Tmp1_S & Out_Row).PasteSpecial xlPasteValues
End If  ' Online or Local
Worksheets(GL_Sheet).Columns(Tmp1_S).AutoFit
FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_01>"
FindRow Ctl_Sheet, "A", Tmp1_I, "<QB_TYPE>"
Worksheets(Ctl_Sheet).Range(Tmp1_S & Tmp1_I).Value = QB_Type
For I = 1 To Array_Cnt
    Worksheets(Raw_Sheet).Range(Col_Src(I) & Tmp2_I & ":" & Col_Src(I) & Tmp3_I).Copy
    Worksheets(GL_Sheet).Range(Col_GL(I) & Out_Row).PasteSpecial xlPasteValues
Next I
Worksheets(GL_Sheet).Range("E5").Activate
ClearClipboard
' Format Hdr's & Ftr's
FindColNumLtr GL_Sheet, 1, Tmp1_I, Col_GL(1), Find_Acct
FindColNumLtr GL_Sheet, 1, Tmp1_I, Col_GL(2), Find_GL_Desc
FindColNumLtr GL_Sheet, 1, Tmp1_I, Col_GL(3), "<CONTRA>"
FindColNumLtr GL_Sheet, 1, Tmp1_I, Col_GL(4), "<BAL>"
FindRow GL_Sheet, "A", Tmp1_I, "<HDR>"
Tmp1_I = Tmp1_I + 1
FindLastRow GL_Sheet, Tmp2_I
For Out_Row = Tmp2_I To Tmp1_I Step -1
If Worksheets(GL_Sheet).Range(Col_GL(1) & Out_Row).Value = "" And Worksheets(GL_Sheet).Range(Col_GL(2) & Out_Row).Value = "" Then
    ' Do Nothing
Else
    Worksheets(GL_Sheet).Range(Col_GL(1) & Out_Row & ":" & Col_GL(4) & Out_Row).Font.Bold = True
    Tmp3_I = InStr(Worksheets(GL_Sheet).Range(Col_GL(1) & Out_Row).Value, "Total") & InStr(Worksheets(GL_Sheet).Range(Col_GL(2) & Out_Row).Value, "Total")
    If Tmp3_I > 0 Then
        Worksheets(GL_Sheet).Range(Col_GL(3) & Out_Row & ":" & Col_GL(4) & Out_Row).Borders(xlEdgeTop).LineStyle = XlLineStyle.xlContinuous
        Rows(Out_Row + 1).Insert Shift:=xlDown
    End If
End If  ' Bold Hdr & Ftr
Next Out_Row
FindColNumLtr D_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow D_Sheet, "A", Tmp1_I, "<REBUILD_GL>"
Worksheets(D_Sheet).Range(Tmp1_S & Tmp1_I).Value = ""
FindColNumLtr D_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_03>"
FindRow D_Sheet, "A", Tmp1_I, "<REBUILD_GL>"
Worksheets(D_Sheet).Range(Tmp1_S & Tmp1_I).Value = "GL Has Been Rebuilt"
'Else
    ' Tmp1_I <> 6 = abort pgm
'End If  ' Tmp1_I = 6 = run pgm
ExitRoutine:
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function



