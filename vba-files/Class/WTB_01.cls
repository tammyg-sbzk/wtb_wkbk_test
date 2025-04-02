VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WTB_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
On Error GoTo ErrSub
Dim RowClick As Long, ColClick As Long
Dim This_Sheet, GL_Sheet, GL_Type, Find_Str, Find_Str_Pfx, Find_Str_Sfx, Find_Col, FindCol_1, FindCol_2, Tmp1_S As String
Dim Hdr_Row, Tmp1_I, Find_Row As Integer

Const Use_GL As String = "GL_01"
Const Find_Col_1 As String = "E"
Const Find_Col_2 As String = "F"
Application.ScreenUpdating = False

RowClick = Selection.Cells.Row
'ColClick = Selection.Cells.Column
This_Sheet = ActiveSheet.Name
FindRow This_Sheet, "A", Hdr_Row, "<HDR>"

If RowClick <= Hdr_Row Then
    Tmp1_I = MsgBox("Row >" & RowClick & "< is in the >" & This_Sheet & "< Header" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Please select a SPLUNGE row and try again", vbExclamation, "Not a Valid Row")
    GoTo ExitRoutine
End If  ' RowClick <= Hdr_Row

Debug.Print "Click Thru"
GL_Type = Application.Evaluate("QB_Type")
Debug.Print "GL_Type>" & GL_Type & "<"
GL_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_GL Then GL_Sheet = WkSheet.Name
Next WkSheet
FindColNumLtr This_Sheet, 1, Tmp1_I, Tmp1_S, "<DESC>"
Find_Str = Worksheets(This_Sheet).Cells(RowClick, Tmp1_I).Value
FindCol_1 = Find_Col_1
If GL_Type = "ONLINE" Then
    FindCol_2 = FindCol_1
Else
    ' GL_Type = "LOCAL"
    FindCol_2 = Find_Col_2
End If
Tmp1_I = InStr(Find_Str, ":")
If Tmp1_I = 0 Then
    ' No Colon
    Find_Col = FindCol_1
    FindRow GL_Sheet, Find_Col, Find_Row, Find_Str
Else
    ' Colon Delimiter
    Find_Str_Pfx = Left(Find_Str, (Tmp1_I - 1))
    Find_Str_Sfx = Mid(Find_Str, (Tmp1_I + 1), 100)
    Find_Col = FindCol_1
    FindRow GL_Sheet, Find_Col, Find_Row, Find_Str_Pfx
    Debug.Print "Find Pfx>" & Find_Str_Pfx & "<Col>" & Find_Col & "<Row>" & Find_Row & "<"
    Tmp1_I = Find_Row
    Find_Col = FindCol_2
    Debug.Print "Find Sfx>" & Find_Str_Sfx & "<Col>" & Find_Col & "<After>" & Tmp1_I & "<"
    FindRowAfter GL_Sheet, Find_Col, Tmp1_I, Find_Row, Find_Str_Sfx
    Debug.Print "Find_Sfx Returns>" & Find_Fow & "<"
End If ' Colon No Colon

Debug.Print "GL Search >" & GL_Sheet & "<Found at>" & Find_Col & Find_Row & "<"
If Find_Row = 0 Then
    Tmp1_I = MsgBox(">" & Find_Str & "< is NOT found in >" & GL_Sheet & "<" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Have you IMPORTED and REBUILT the correct GL ?" & Chr(13) & Chr(10) & "Please check and try again", vbExclamation, "Account Description NOT Found")
    GoTo ExitRoutine
End If  ' RowClick <= Hdr_Row

'Worksheets(GL_Sheet).Activate
Application.GoTo Reference:=Worksheets(GL_Sheet).Range("D" & Find_Row), scroll:=True

ExitRoutine:
Exit Sub

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Sub

