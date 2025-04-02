Attribute VB_Name = "Mod_Dash"
'----------
' Obj_Hide ' Prep_For_Tax
' Obj_UnHide ' Recall Objects
'----------

Function Obj_Hide()
Const VBA_Name As String = "Prep_For_Tax"
Debug.Print ">" & VBA_Name & "<"
On Error GoTo ErrSub
Dim D_Sheet, R_Sheet, WTB_Sheet, Ctl_Sheet As String

Const Use_WTB As String = "WTB_01"
Const Use_ReadMe As String = "ReadMe_01"
Const Use_Dashboard As String = "Dashboard"
Const Use_CTL As String = "CTL_01"
Const Btn_Ref As String = "Btn_WTB_Refresh"
Const Btn_Del As String = "Btn_WTB_Delete"
Const Btn_Rec As String = "Btn_WTB_Reconcile"


WTB_Sheet = "NOF"
D_Sheet = "NOF"
R_Sheet = "NOF"
Ctl_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_ReadMe Then R_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_Dashboard Then D_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_CTL Then Ctl_Sheet = WkSheet.Name
Next WkSheet

If WTB_Sheet <> "NOF" And R_Sheet <> "NOF" And D_Sheet <> "NOF" And Ctl_Sheet <> "NOF" Then
    ' OK - Fall Thru
    Debug.Print "Worksheets Found"
Else
    Msg_01 = VBA_Name & Chr(13) & Chr(10)
    If Ctl_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [CONTROL Sheet]" & Chr(13) & Chr(10)
    If R_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [READ ME Sheet]" & Chr(13) & Chr(10)
    If D_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [DASHBOARD Sheet]" & Chr(13) & Chr(10)
    If WTB_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [WTB - Working Trial Balance Sheet]" & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & "Please make a note of this message and contact Program Development"
    Tmp1_I = MsgBox(Msg_01, vbExclamation, "Worksheet Not Found")
    GoTo ExitRoutine
End If

Worksheets(WTB_Sheet).Shapes(Btn_Ref).Visible = False
Worksheets(WTB_Sheet).Shapes(Btn_Del).Visible = False
Worksheets(WTB_Sheet).Shapes(Btn_Rec).Visible = False
Worksheets(R_Sheet).Visible = False
Worksheets(D_Sheet).Visible = False
Worksheets("Year-End Questions").Visible = False
Worksheets("Profit and Loss Monthly").Visible = False

ExitRoutine:
Debug.Print "Complete>" & VBA_Name & "<"
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Function Obj_HideCO()
Const VBA_Name As String = "Prep_For_CO"
Debug.Print ">" & VBA_Name & "<"
On Error GoTo ErrSub
Dim D_Sheet, R_Sheet, WTB_Sheet, Ctl_Sheet As String

Const Use_WTB As String = "WTB_01"
Const Use_ReadMe As String = "ReadMe_01"
Const Use_Dashboard As String = "Dashboard"
Const Use_CTL As String = "CTL_01"
Const Btn_Ref As String = "Btn_WTB_Refresh"
Const Btn_Del As String = "Btn_WTB_Delete"
Const Btn_Rec As String = "Btn_WTB_Reconcile"


WTB_Sheet = "NOF"
D_Sheet = "NOF"
R_Sheet = "NOF"
Ctl_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_ReadMe Then R_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_Dashboard Then D_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_CTL Then Ctl_Sheet = WkSheet.Name
Next WkSheet

If WTB_Sheet <> "NOF" And R_Sheet <> "NOF" And D_Sheet <> "NOF" And Ctl_Sheet <> "NOF" Then
    ' OK - Fall Thru
    Debug.Print "Worksheets Found"
Else
    Msg_01 = VBA_Name & Chr(13) & Chr(10)
    If Ctl_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [CONTROL Sheet]" & Chr(13) & Chr(10)
    If R_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [READ ME Sheet]" & Chr(13) & Chr(10)
    If D_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [DASHBOARD Sheet]" & Chr(13) & Chr(10)
    If WTB_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [WTB - Working Trial Balance Sheet]" & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & "Please make a note of this message and contact Program Development"
    Tmp1_I = MsgBox(Msg_01, vbExclamation, "Worksheet Not Found")
    GoTo ExitRoutine
End If

Worksheets(WTB_Sheet).Shapes(Btn_Ref).Visible = False
Worksheets(WTB_Sheet).Shapes(Btn_Del).Visible = False
Worksheets(WTB_Sheet).Shapes(Btn_Rec).Visible = False
Worksheets(R_Sheet).Visible = False
Worksheets(D_Sheet).Visible = False
'Worksheets("Year-End Questions").Visible = False
'Worksheets("Profit and Loss Monthly").Visible = False

ExitRoutine:
Debug.Print "Complete>" & VBA_Name & "<"
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Function Obj_UnHide()
Const VBA_Name As String = "Prep_For_Tax"
Debug.Print ">" & VBA_Name & "<"
On Error GoTo ErrSub
Dim D_Sheet, R_Sheet, WTB_Sheet, Ctl_Sheet As String

Const Use_WTB As String = "WTB_01"
Const Use_ReadMe As String = "ReadMe_01"
Const Use_Dashboard As String = "Dashboard"
Const Use_CTL As String = "CTL_01"
Const Btn_Ref As String = "Btn_WTB_Refresh"
Const Btn_Del As String = "Btn_WTB_Delete"
Const Btn_Rec As String = "Btn_WTB_Reconcile"


WTB_Sheet = "NOF"
D_Sheet = "NOF"
R_Sheet = "NOF"
Ctl_Sheet = "NOF"
For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_CTL Then Ctl_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_ReadMe Then R_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_Dashboard Then D_Sheet = WkSheet.Name
Next WkSheet

If WTB_Sheet <> "NOF" And R_Sheet <> "NOF" And D_Sheet <> "NOF" And Ctl_Sheet <> "NOF" Then
    ' OK - Fall Thru
    Debug.Print "Worksheets Found"
Else
    Msg_01 = VBA_Name & Chr(13) & Chr(10)
    If R_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [READ ME Sheet]" & Chr(13) & Chr(10)
    If Ctl_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [CONTROL Sheet]" & Chr(13) & Chr(10)
    If D_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [DASHBOARD Sheet]" & Chr(13) & Chr(10)
    If WTB_Sheet = "NOF" Then Msg_01 = Msg_01 & "Can NOT find the [WTB - Working Trial Balance Sheet]" & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & Chr(13) & Chr(10)
    Msg_01 = Msg_01 & "Please make a note of this message and contact Program Development"
    Tmp1_I = MsgBox(Msg_01, vbExclamation, "Worksheet Not Found")
    GoTo ExitRoutine
End If

Worksheets(WTB_Sheet).Shapes(Btn_Ref).Visible = True
Worksheets(WTB_Sheet).Shapes(Btn_Del).Visible = True
Worksheets(WTB_Sheet).Shapes(Btn_Rec).Visible = True
Worksheets(R_Sheet).Visible = True
Worksheets(D_Sheet).Visible = True
Worksheets("Year-End Questions").Visible = True
Worksheets("Profit and Loss Monthly").Visible = True

ExitRoutine:
Debug.Print "Complete>" & VBA_Name & "<"
Exit Function

ErrSub:
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

