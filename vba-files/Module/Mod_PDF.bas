Attribute VB_Name = "Mod_PDF"
'----------
' PDF_ALL
' PDF_AJE
' PDF_WTB
' PDF_GL
' PDF_BS
' PDF_PL
' PDF_Print
'----------

Function PDF_ALL()
Const VBA_Name As String = "PDF_ALL"
Dim GL_Sheet, BS_Sheet, PL_Sheet, WTB_Sheet, Msg01 As String
Dim Tmp1_I As Integer

Const Use_GL As String = "GL_01"
Const Use_BS As String = "BS_01"
Const Use_PL As String = "PL_01"
Const Use_WTB As String = "WTB_01"
Const Use_AJE As String = "AJE_01"

Const D_Sheet As String = "Dashboard"
GL_Sheet = "NOF"
BS_Sheet = "NOF"
PL_Sheet = "NOF"
WTB_Sheet = "NOF"
AJE_Sheet = "NOF"

For Each WkSheet In ThisWorkbook.Worksheets
    If WkSheet.CodeName = Use_GL Then GL_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_BS Then BS_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_PL Then PL_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_WTB Then WTB_Sheet = WkSheet.Name
    If WkSheet.CodeName = Use_AJE Then AJE_Sheet = WkSheet.Name
Next WkSheet

If GL_Sheet <> "NOF" And BS_Sheet <> "NOF" And PL_Sheet <> "NOF" And WTB_Sheet <> "NOF" And AJE_Sheet <> "NOF" Then
    Worksheets(GL_Sheet).Activate
    PDF_GL
    Worksheets(BS_Sheet).Activate
    PDF_BS
    Worksheets(PL_Sheet).Activate
    PDF_PL
    Worksheets(WTB_Sheet).Activate
    PDF_WTB
    Worksheets(AJE_Sheet).Activate
    PDF_AJE
    Worksheets(D_Sheet).Activate
Else
    Msg01 = "One of the required Worksheets is missing." & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    If GL_Sheet = "NOF" Then Msg01 = Msg01 & "General Ledger Sheet is missing, or is no longer a recognizable GL Sheet" & Chr(13) & Chr(10)
    If BS_Sheet = "NOF" Then Msg01 = Msg01 & "Balance Sheet is missing, or is no longer a recognizable Balance Sheet" & Chr(13) & Chr(10)
    If PL_Sheet = "NOF" Then Msg01 = Msg01 & "Profit & Loss Sheet is missing, or is no longer a recognizable P & L Sheet" & Chr(13) & Chr(10)
    If WTB_Sheet = "NOF" Then Msg01 = Msg01 & "Working Trial Balance Sheet is missing, or is no longer a recognizable WTB Sheet" & Chr(13) & Chr(10)
    If AJE_Sheet = "NOF" Then Msg01 = Msg01 & "Ajusting Journal Entry Sheet is missing, or is no longer a recognizable AJE Sheet" & Chr(13) & Chr(10)
    Msg01 = Msg01 & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development"
    Tmp1_I = MsgBox(Msg01, vbExclamation, "Missing Worksheet")
End If

End Function

Function PDF_AJE()
Const VBA_Name As String = "PDF_AJE"
Dim Act_Sheet, Tmp_Rng, Tmp1_S, Tmp2_S, Hdr_Repeat As String
Dim Tmp1_I, Tmp2_I As Integer

Const Ctl_Sheet As String = "CONTROL"

FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow Ctl_Sheet, "A", Tmp1_I, "<PDF_AJE>"
Hdr_Repeat = Worksheets(Ctl_Sheet).Range(Tmp1_S & Tmp1_I).Value

Act_Sheet = ActiveSheet.Name
FindColNumLtr Act_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_01>"
FindRow Act_Sheet, "A", Tmp1_I, "<HDR>"
Tmp1_I = Tmp1_I - 1
FindLastColumn Act_Sheet, Tmp2_I
NumToLtr Tmp2_I, Tmp2_S
FindLastRow Act_Sheet, Tmp2_I
Tmp2_I = Tmp2_I + 1
Tmp_Rng = Tmp1_S & Tmp1_I & ":" & Tmp2_S & Tmp2_I
PDF_Print Tmp_Rng, Hdr_Repeat
End Function

Function PDF_WTB()
Const VBA_Name As String = "PDF_WTB"
Dim Act_Sheet, Tmp_Rng, Tmp1_S, Tmp2_S, Hdr_Repeat As String
Dim Tmp1_I, Tmp2_I As Integer

Const Ctl_Sheet As String = "CONTROL"

FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow Ctl_Sheet, "A", Tmp1_I, "<PDF_WTB>"
Hdr_Repeat = Worksheets(Ctl_Sheet).Range(Tmp1_S & Tmp1_I).Value

Act_Sheet = ActiveSheet.Name
FindColNumLtr Act_Sheet, 1, Tmp1_I, Tmp1_S, "<ACCT>"
FindRow Act_Sheet, "A", Tmp1_I, "<HDR>"
Tmp1_I = Tmp1_I - 1
FindLastColumn Act_Sheet, Tmp2_I
NumToLtr Tmp2_I, Tmp2_S
FindLastRow Act_Sheet, Tmp2_I
Tmp2_I = Tmp2_I + 1
Tmp_Rng = Tmp1_S & Tmp1_I & ":" & Tmp2_S & Tmp2_I
PDF_Print Tmp_Rng, Hdr_Repeat
End Function

Function PDF_GL()
Const VBA_Name As String = "PDF_GL"
Dim Act_Sheet, Tmp_Rng, Tmp1_S, Tmp2_S, Hdr_Repeat As String
Dim Tmp1_I, Tmp2_I As Integer

Const Ctl_Sheet As String = "CONTROL"

FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow Ctl_Sheet, "A", Tmp1_I, "<PDF_GL>"
Hdr_Repeat = Worksheets(Ctl_Sheet).Range(Tmp1_S & Tmp1_I).Value

Act_Sheet = ActiveSheet.Name
FindColNumLtr Act_Sheet, 1, Tmp1_I, Tmp1_S, "<ACCT>"
FindRow Act_Sheet, "A", Tmp1_I, "<HDR>"
Tmp1_I = Tmp1_I - 1
FindColNumLtr Act_Sheet, 1, Tmp2_I, Tmp2_S, "<NOTES>"
FindLastRow Act_Sheet, Tmp2_I
Tmp_Rng = Tmp1_S & Tmp1_I & ":" & Tmp2_S & Tmp2_I
PDF_Print Tmp_Rng, Hdr_Repeat
End Function

Function PDF_BS()
Const VBA_Name As String = "PDF_BS"
Dim Act_Sheet, Tmp_Rng, Tmp1_S, Tmp2_S, Hdr_Repeat As String
Dim Tmp1_I, Tmp2_I As Integer

Const Ctl_Sheet As String = "CONTROL"

FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow Ctl_Sheet, "A", Tmp1_I, "<PDF_BS>"
Hdr_Repeat = Worksheets(Ctl_Sheet).Range(Tmp1_S & Tmp1_I).Value

Act_Sheet = ActiveSheet.Name
FindColNumLtr Act_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_01>"
FindRow Act_Sheet, "A", Tmp1_I, "<HDR-1>"
FindLastColumn Act_Sheet, Tmp2_I
NumToLtr Tmp2_I, Tmp2_S
FindLastRow Act_Sheet, Tmp2_I
Tmp_Rng = Tmp1_S & Tmp1_I & ":" & Tmp2_S & Tmp2_I
PDF_Print Tmp_Rng, Hdr_Sheet
End Function

Function PDF_PL()
Const VBA_Name As String = "PDF_PL"
Dim Act_Sheet, Tmp_Rng, Tmp1_S, Tmp2_S, Hdr_Repeat As String
Dim Tmp1_I, Tmp2_I As Integer

Const Ctl_Sheet As String = "CONTROL"

FindColNumLtr Ctl_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_02>"
FindRow Ctl_Sheet, "A", Tmp1_I, "<PDF_PL>"
Hdr_Repeat = Worksheets(Ctl_Sheet).Range(Tmp1_S & Tmp1_I).Value

Act_Sheet = ActiveSheet.Name
FindColNumLtr Act_Sheet, 1, Tmp1_I, Tmp1_S, "<COL_01>"
FindRow Act_Sheet, "A", Tmp1_I, "<HDR-1>"
FindLastColumn Act_Sheet, Tmp2_I
NumToLtr Tmp2_I, Tmp2_S
FindLastRow Act_Sheet, Tmp2_I
Tmp_Rng = Tmp1_S & Tmp1_I & ":" & Tmp2_S & Tmp2_I
Worksheets(Act_Sheet).PageSetup.PrintTitleRows = TopRepeat
PDF_Print Tmp_Rng, Hdr_Repeat
End Function

Function PDF_Print(TmpPDF, Tmp_Repeat)
On Error GoTo ErrSub
Dim Tmp_Str, Tmp_Hdr, Tmp_Name, Tmp_YrEnd, W_Sheet, WB_Path As String
Dim TmpRng As Range

Const Logo_Path As String = "F:\1NET\2Logos & Branding\SBZK_Logo_CMYK_2-Transparent.png"

Tmp_Name = Application.Evaluate("Name_Client")
Tmp_YrEnd = Application.Evaluate("Yr_End")
W_Sheet = ActiveSheet.Name
WB_Path = ThisWorkbook.FullName
WB_Path = Left(WB_Path, (InStrRev(WB_Path, ".") - 1)) & "_" & W_Sheet & ".PDF"
With Application.ActiveSheet.PageSetup.LeftHeaderPicture
    .Filename = Logo_Path
    .Height = 200
    .Width = 300
    .Brightness = 0.36
    .ColorType = msoPictureAutomatic
    .Contrast = 0.39
    .CropBottom = 0
    .CropLeft = 0
    .CropRight = 0
    .CropTop = 0
End With

ActiveSheet.PageSetup.LeftHeader = "&G"
Application.ActiveSheet.PageSetup.RightHeader = "&""Century Gothic,Bold""&16" & "Page " & "&P" & " of " & "&N"
Set TmpRng = Range(TmpPDF)
Worksheets(W_Sheet).PageSetup.PrintArea = TmpPDF
Worksheets(W_Sheet).PageSetup.PrintTitleRows = Tmp_Repeat
TmpRng.ExportAsFixedFormat Type:=xlTypePDF, Filename:=WB_Path, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=True, OpenAfterPublish:=True

ExitRoutine:
Exit Function

ErrSub:
If Err = 1004 Then
    Tmp1_I = MsgBox("PDF for Worksheet >" & W_Sheet & "< is currently open" & Chr(13) & Chr(10) & "Please close this PDF, and try again", vbInformation, "PDF is Already Open")
    GoTo ExitRoutine
End If  ' PDF Already Open
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

