Attribute VB_Name = "Mod_Tools"
Sub FilterOff()
Dim A_Sheet As String
A_Sheet = ActiveSheet.Name
Worksheets(A_Sheet).AutoFilterMode = False   ' Filter Off
End Sub

Function CodeNameToTabName(Tmp_Code, Tmp_Tab)
With ActiveWorkbook.VBProject
    Tmp_Tab = Worksheets(CStr(.VBComponents(Tmp_Code).Properties("Name"))).Name
End With
End Function

Function CountColorCells(CntRng As Range, CntClr As Range)
Dim CntClrValue As Integer
Dim CntTot As Integer
CntClrValue = CountColor.Interior.ColorIndex
Set rCell = CntRng
For Each rCell In CntRng
    If rCell.Interior.ColorIndex = CntClrValue Then
        CntTot = CntTot + 1
    End If
Next rCell
CountColorCells = CntTot
End Function

Function GetColorCount(CountRange As Range, CountColor As Range)
Dim CountColorValue As Integer
Dim TotalCount As Integer
CountColorValue = CountColor.Interior.ColorIndex
Set rCell = CountRange
For Each rCell In CountRange
  If rCell.Interior.ColorIndex = CountColorValue Then
    TotalCount = TotalCount + 1
  End If
Next rCell
GetColorCount = TotalCount
End Function

Sub WS_NameSave()
Dim Col_Name As Integer
Dim TmpStr01, TmpStr02 As String
Const Row_Name As Integer = 4

FindColNumLtr ActiveSheet.Name, 1, Col_Name, TmpStr02, "<DESC>"
ActiveSheet.Cells(4, 1).Value = UCase(Trim(ActiveSheet.Cells(4, Col_Name).Value))
End Sub

Sub WS_NameCheck(WS_Name)
Dim TmpNum01 As Integer
Dim TmpStr01, TmpStr02, TmpStr03 As String

Const T_Sheet As String = "ToolBar"
Const FindNameCol As String = "<COL02>"
Const Row_Name As Integer = 4

TmpStr03 = WS_Name
FindColNumLtr TmpStr03, 1, TmpNum01, TmpStr02, "<DESC>"
TmpStr01 = UCase(Trim(Worksheets(TmpStr03).Cells(4, 1).Value))
TmpStr02 = UCase(Trim(Worksheets(TmpStr03).Cells(4, TmpNum01).Value))
If TmpStr01 <> TmpStr02 Then
'Debug.Print "Names Are diffferent"
Worksheets(TmpStr03).Cells(4, 1).Value = TmpStr02
FindColNumLtr T_Sheet, 1, TmpNum01, TmpStr03, FindNameCol
'Debug.Print "Name Column TmpStr01>" & TmpStr01 & "<"
FindRow T_Sheet, TmpStr03, TmpNum01, TmpStr01
'Debug.Print "At Col>" & TmpStr03 & "<Row>" & TmpNum01 & "<insert>" & TmpStr02 & "<"
If TmpNum01 > 0 Then Worksheets(T_Sheet).Range(TmpStr03 & TmpNum01).Value = TmpStr02
End If
End Sub

Public Sub SortWorkSheets()

Dim currentUpdating As Boolean
currentUpdating = Application.ScreenUpdating

Application.ScreenUpdating = False

For Each xlSheet In ActiveWorkbook.Worksheets
   For Each xlWTB In ActiveWorkbook.Worksheets
      If LCase(xlWTB.Name) < LCase(xlSheet.Name) Then
         xlWTB.Move Before:=xlSheet
      End If
   Next xlWTB
Next xlSheet

Application.ScreenUpdating = currentUpdating

End Sub

Sub ClearClipboard()
Application.CutCopyMode = False ' Clear Clipboard after Client Sheet
End Sub

Sub FindLastRowInCol(TmpSheet, TmpCol, TmpEnd)
    ' Last Row used in Col TmpCol
    TmpEnd = Worksheets(TmpSheet).Cells(Worksheets(TmpSheet).Rows.Count, TmpCol).End(xlUp).Row
End Sub

Sub FindLastRowOnSheet(TmpSheet, TmpEnd)
    ' Last Row Used in ANY Col on TmpSheet
TmpEnd = Worksheets(TmpSheet).Cells.Find(What:="*", After:=Range("ZZ65536"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Sub

Sub FindLastColumnOnSheet(TmpSheet, TmpEnd)
    ' Last Row Used in ANY Col on TmpSheet
TmpEnd = Worksheets(TmpSheet).Cells.Find(What:="*", After:=Range("ZZ65536"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Row
End Sub

Sub EchoOff()
    Application.Calculation = xlManual
    ' Application.ScreenUpdating = False ' Leave "ON" for Status Updates
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Sub EchoOn()
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Sub TurnOffCalcAll()
' Set Calc OFF in Individual Worksheets
Dim Wksht As Worksheet

For Each Wksht In ActiveWorkbook.Worksheets
    Worksheets(Wksht.Name).EnableCalculation = False
Next Wksht
ActiveSheet.EnableCalculation = True
End Sub

Sub TurnOffCalcOne(Calc_Sheet)
' Turn OFF Calculation on Active Sheet
Worksheets(Calc_Sheet).EnableCalculation = False
End Sub

Sub ForceSheetCalculateOne(Calc_Sheet)
' Force Calculation on Active Sheet
Worksheets(Calc_Sheet).EnableCalculation = True
Worksheets(Calc_Sheet).Calculate
End Sub

Sub ForceAllSheetCalc()
' Set Calc ON, force Calc, Set Calc OFF for ALL Worksheets
Dim Wksht As Worksheet

For Each Wksht In ActiveWorkbook.Worksheets
    Worksheets(Wksht.Name).EnableCalculation = True
    Worksheets(Wksht.Name).Calculate
    Worksheets(Wksht.Name).EnableCalculation = False
Next Wksht
ActiveSheet.EnableCalculation = True
End Sub

Sub SaveDateOnControl(S_Beg, S_End, S_FindCol, S_FindRow)
Dim S_Tstr As String
Dim S_Col, S_Row As Integer
Const C_Sheet = "Control"

FindColNumLtr C_Sheet, 1, S_Col, S_Tstr, S_FindCol
FindRow C_Sheet, "A", S_Row, S_FindRow
Worksheets(C_Sheet).Cells(S_Row, S_Col).Value = S_End
Worksheets(C_Sheet).Cells(S_Row, (S_Col + 1)).Value = Format((S_End - S_Beg), "h:mm:ss")
End Sub

Function FindLastRow(L_Sheet, L_Row)
'Debug.Print "FindLastRow>" & L_Sheet & "<"
On Error Resume Next
L_Row = 0
L_Row = Worksheets(L_Sheet).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
End Function

Function FindLastColumn(C_Sheet, C_End)
On Error Resume Next
C_End = 0
C_End = Worksheets(C_Sheet).Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
End Function

Sub PurgeSheet(P_Sheet)
Dim RowBeg, RowEnd As Long

FindRow P_Sheet, "A", RowBeg, "<Hdr>"
RowBeg = RowBeg + 1
RowEnd = Worksheets(P_Sheet).Range("A65536").End(xlUp).Row ' Find Last Row of Data
'Debug.Print "Purge>A" & RowBeg & ":A" & RowEnd & "<"
If RowEnd > RowBeg Then
' Purge Sheet
    Worksheets(P_Sheet).Range("A" & RowBeg & ":" & "A" & RowEnd).EntireRow.Delete
End If
End Sub

Sub PurgeSheet_End(P_Sheet)
Dim RowBeg, RowEnd As Long

FindRow P_Sheet, "A", RowBeg, "<Hdr>"
RowBeg = RowBeg + 1
FindRow P_Sheet, "A", RowEnd, "<End>"
If RowEnd > RowBeg Then
' Purge Sheet
    Worksheets(P_Sheet).Range("A" & RowBeg & ":" & "A" & RowEnd).EntireRow.Delete
End If
End Sub

Sub ColumnAutoFit(FitSheet)
Dim HdrRow, ColCnt As Integer
Dim ColLtr, ColRange As String
Const StartCol As String = "C"

FindRow FitSheet, "A", HdrRow, "<Hdr>"
ColCnt = 1
Do While Trim(Worksheets(FitSheet).Cells(HdrRow, ColCnt).Value) > ""
ColCnt = ColCnt + 1
Loop
ColCnt = ColCnt - 1
NumToLtr ColCnt, ColLtr
ColRange = StartCol & ":" & ColLtr
' Debug.Print "ColRange>" & ColRange & "<"
Worksheets(FitSheet).Range(ColRange).Columns.AutoFit
End Sub

Sub FormulaToStatic(TmpSheet, TmpLeft, TmpTop, TmpRight, TmpBottom)
' Switch Cell from Formula to static Value
With Worksheets(TmpSheet).Range(TmpLeft & TmpTop & ":" & TmpRight & TmpBottom)
    .Copy
    .PasteSpecial xlPasteValues
 End With
  Application.CutCopyMode = False
End Sub

Sub PrepHdr(PrepSheet)
Dim PrepCol As Integer
Dim PrepStr As String
PrepCol = 1
Do While Trim(Worksheets(PrepSheet).Cells(1, PrepCol).Value) > ""
PrepStr = UCase(Trim(Worksheets(PrepSheet).Cells(1, PrepCol).Value))
CleanString PrepStr
' Debug.Print "PrepCol>" & PrepCol & "<>" & Worksheets(PrepSheet).Cells(1, PrepCol).Value & "<"
If Left(PrepStr, 1) = "<" Then
' Already Prepped
Else
    PrepStr = "<" & PrepStr & ">"
    Worksheets(PrepSheet).Cells(1, PrepCol).Value = PrepStr
End If  ' "<" = Prepped
PrepCol = PrepCol + 1
Loop
End Sub

Sub SheetHide(Tsheet)
Worksheets(Tsheet).Visible = False
End Sub

Sub SheetUnHide(Tsheet)
Worksheets(Tsheet).Visible = True
End Sub

Sub SetColumnToNeg(Tsheet, Tstr)
Dim Rdisc, RD As Range

Set Rdisc = Worksheets(Tsheet).Range(Tstr)
For Each RD In Rdisc
    RD.Value = -1 * RD.Value
Next RD
End Sub

Function FindColNumLtr(Tsheet, Trow, Tnum, Tltr, Tstr)

On Error GoTo ErrHandler

'Debug.Print "FindColNumLtr"
'Debug.Print "Tsheet>" & Tsheet & "<"
'Debug.Print "Trow>" & Trow & "<"

Tnum = Worksheets(Tsheet).Rows(Trow).Find(Tstr, SearchOrder:=xlByColumns, LookIn:=xlValues, SearchDirection:=xlNext).Column

If Tnum > 26 Then             ' Convert Column Number to Column Letter
    Tltr = Chr(Int((Tnum - 1) / 26) + 64) & Chr(((Tnum - 1) Mod 26) + 65)
Else
    Tltr = Chr(Tnum + 64)
End If

'Debug.Print "GoodTnum>" & Tnum & "<"
'Debug.Print "GoodTltr>" & Tltr & "<"
Exit Function

ErrHandler:
'Debug.Print "ErrTnum>" & Tnum & "<"
'Debug.Print "ErrTltr>" & Tltr & "<"
Tnum = 0
Tltr = ""

End Function

Function NumToLtr(Tnum, Tltr)
If Tnum > 26 Then             ' Convert Column Number to Column Letter
    Tltr = Chr(Int((Tnum - 1) / 26) + 64) & Chr(((Tnum - 1) Mod 26) + 65)
Else
    Tltr = Chr(Tnum + 64)
End If
End Function


Sub ClearSheet(ClrSheet, ClrStr)
Dim ClrStart, ClrEnd, ClrNum As Long
Dim ClrLtr As String

FindRow ClrSheet, "A", ClrStart, "<Hdr>"
If ClrStr = "NoCol" Then
' No Columns to Delete
Else
FindColNumLtr ClrSheet, ClrStart, ClrNum, ClrLtr, ClrStr
'Debug.Print "ClrSheet>" & ClrSheet & "<ClrStr>" & ClrStr & "<ClrLtr>" & ClrLtr & "<"
If ClrLtr > "" Then
    Worksheets(ClrSheet).Range(ClrLtr & "1:ZZ" & 1).EntireColumn.Delete
End If
End If  ' NoCol
ClrStart = ClrStart + 1
' Find EndRow
ClrEnd = Worksheets(ClrSheet).Cells(Rows.Count, 1).End(xlUp).Row
' Delete between StartRow + 1 and EndRow
If ClrStart <= ClrEnd Then Worksheets(ClrSheet).Range("A" & ClrStart & ":" & "A" & ClrEnd).EntireRow.Delete
End Sub

Function FindRowReverse(FRsheet, FRCol, FRbeg, FRend, FRrow, FRstr)
Dim Tmp_Row As Long
'Debug.Print "FRsheet>" & FRsheet & "<"
'Debug.Print "FRcol>" & FRCol & "<"
'Debug.Print "FRbeg>" & FRbeg & "<"
'Debug.Print "FRend>" & FRend & "<"
'Debug.Print "FRrow>" & FRrow & "<"
'Debug.Print "FRstr>" & FRstr & "<"
Set Target = Worksheets(FRsheet).Range(FRCol & FRbeg & ":" & FRCol & FRend).Find(FRstr, LookIn:=xlValues, SearchDirection:=xlPrevious)
If Not Target Is Nothing Then
    FRrow = Rows(Target.Row).Row
Else
  FRrow = 0
End If

End Function


Function FindRow(FRsheet, FRCol, FRrow, FRstr)
'Debug.Print "FRsheet>" & FRsheet & "<"
'Debug.Print "FRcol>" & FRcol & "<"
'Debug.Print "FRrow>" & FRrow & "<"
'Debug.Print "FRstr>" & FRstr & "<"
       Set Target = Worksheets(FRsheet).Columns(FRCol).Find(FRstr, LookIn:=xlValues)
        If Not Target Is Nothing Then
          FRrow = Rows(Target.Row).Row
        Else
          FRrow = 0
        End If
End Function

Function FindRowAfter(FRsheet, FRCol, FRafter, FRrow, FRstr)
On Error GoTo ErrSub
'Debug.Print "Find Row After"
'Debug.Print "FRA FRsheet>" & FRsheet & "<"
'Debug.Print "FRA FRcol>" & FRcol & "<"
'Debug.Print "FRA FRafter>" & FRafter & "<"
'Debug.Print "FRA FRrow>" & FRrow & "<"
'Debug.Print "FRA FRstr>" & FRstr & "<"
'Dim TmpLast As Long
'Debug.Print "Call FindLastRow"
FindLastRow FRsheet, TmpLast
Debug.Print "Last Row >" & TmpLast & "<"
FRrow = 0

FRrow = Worksheets(FRsheet).Range(FRCol & FRafter & ":" & FRCol & TmpLast).Find(What:=FRstr, LookIn:=xlValues).Row

ExitRoutine:
Exit Function

ErrSub:
If Err = 91 Then
    FRrow = 0
    GoTo ExitRoutine
End If
Tmp1_I = MsgBox("Dow VBA Error # " & Err & " has occured in " & VBA_Name & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Please make a note of this message and contact Program Development", vbExclamation, "VBA Error")
GoTo ExitRoutine

End Function

Function FindBegColumn(FAW_Beg, FAW_Sheet, FAW_Find, FAW_Col, FAW_Row)
'Debug.Print "FindAnyWhere()>" & FAW_Sheet & "<Find>" & FAW_Find & "<"
Dim FAW_Str, End_Col As String
Dim FAW_I As Integer
Dim RngFound As Range

FindLastColumn FAW_Sheet, End_Row
NumToLtr End_Row, End_Col
FindLastRow FAW_Sheet, End_Row

With Worksheets(FAW_Sheet).Range(FAW_Beg & 1 & ":" & End_Col & End_Row)
    Set RngFound = .Find(FAW_Find, LookIn:=xlValues)
    If Not RngFound Is Nothing Then
        'Debug.Print "Found>" & FAW_Find & "<"
        FAW_Str = RngFound.Address
        'Debug.Print "FAW_Str>" & FAW_Str & "<"
        FAW_Str = Mid(FAW_Str, 2, 10)
        FAW_I = InStr(FAW_Str, "$")
        FAW_Col = Left(FAW_Str, (FAW_I - 1))
        FAW_I = InStr(FAW_Str, "$")
        FAW_Row = Mid(FAW_Str, (FAW_I + 1), 10)
    Else
        'Debug.Print "Not Found>" & FAW_Find & "<"
        FAW_Col = "NOF"
        FAW_Row = 0
    End If
End With

End Function

Function FindAnyWhere(FAW_Sheet, FAW_Find, FAW_Col, FAW_Row)
'Debug.Print "FindAnyWhere()>" & FAW_Sheet & "<Find>" & FAW_Find & "<"
Dim FAW_Str As String
Dim FAW_I As Integer
Dim RngFound As Range

With Worksheets(FAW_Sheet).Cells
    Set RngFound = .Find(FAW_Find, LookIn:=xlValues)
    If Not RngFound Is Nothing Then
'        Debug.Print "Found>" & FAW_Find & "<"
        FAW_Str = RngFound.Address
'        Debug.Print "FAW_Str>" & FAW_Str & "<"
        FAW_Str = Mid(FAW_Str, 2, 10)
        FAW_I = InStr(FAW_Str, "$")
        FAW_Col = Left(FAW_Str, (FAW_I - 1))
        FAW_I = InStr(FAW_Str, "$")
        FAW_Row = Mid(FAW_Str, (FAW_I + 1), 10)
    Else
'        Debug.Print "Not Found>" & FAW_Find & "<"
        FAW_Col = "NOF"
        FAW_Row = 0
    End If
End With

End Function

Function FindAnyWhereAfterRow(FAW_Sheet, FAW_After, FAW_Find, FAWfind_Col, FAWfind_Row, FAWrtn_Col, FAWrtn_Row)
Debug.Print "FindAnyWhere()>" & FAW_Sheet & "<After>" & FAW_After & "<Find>" & FAW_Find & "<FAW_Col>" & FAW_Col & "<FAW_Row>" & FAW_Row & "<"
Dim FAW_Str As String
Dim FAW_I As Integer
Dim RngFound As Range

With Worksheets(FAW_Sheet).Range("A" & FAW_After & ":" & FAWfind_Col & FAWfind_Row)
    Set RngFound = .Find(FAW_Find, LookIn:=xlValues)
    If Not RngFound Is Nothing Then
        Debug.Print "Found>" & FAW_Find & "<"
        FAW_Str = RngFound.Address
        Debug.Print "FAW_Str>" & FAW_Str & "<"
        FAW_Str = Mid(FAW_Str, 2, 10)
        FAW_I = InStr(FAW_Str, "$")
        FAWrtn_Col = Left(FAW_Str, (FAW_I - 1))
        FAW_I = InStr(FAW_Str, "$")
        FAWrtn_Row = Mid(FAW_Str, (FAW_I + 1), 10)
    Else
        Debug.Print "Not Found>" & FAW_Find & "<"
        FAWrtn_Col = "NOF"
        FAWrtn_Row = 0
    End If
End With

End Function

Function CleanStringGood(In_Str, Clean_Str)
Dim Tstr, TmpStr, PrevStr, TmpLtr As String
Dim I As Integer

Const GoodChr As String = "1234567890_ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const BadChr As String = "!@#$%^&*()+={[}]|\:;<,>.?/"
Clean_Str = ""
Tstr = UCase(In_Str)
PrevStr = "_"
For I = 1 To Len(Tstr)
    TmpLtr = Mid(Tstr, I, 1)
    If InStr(GoodChr, TmpLtr) > 0 Then
        ' Include a 'Good' Chr
        Clean_Str = Clean_Str & TmpLtr
    Else
        ' Not a 'Good' Chr
        If TmpLtr = " " Then
            ' Replace " " space with "_"
           If Right(Clean_Str, 1) <> "_" Then Clean_Str = Clean_Str & "_"
        Else
            If InStr(BadChr, TmpLtr) > 0 Then
            ' Replace 'Bad' Chr with "_"
            ' Skip a 'Bad' Chr
            Else
            ' Not a 'Good' or 'Bad' chr so leave out
            End If  ' Bad Chr
        End If  ' Blank or Bad Chr
    End If  ' Good
    ' PrevStr = TmpStr
Next I
' Tstr = TmpStr
End Function

Sub SortArea(Sort_Sheet, SortColStart, SortRowStart, SortColRight, SortRowEnd, SortCol_1, SortCol_2)
Dim SortEnd As Long

Debug.Print "SortSheet>" & Sort_Sheet & "<"
Debug.Print "SortColStart>" & SortColStart & "<"
Debug.Print "SortRowStart>" & SortRowStart & "<"
Debug.Print "SortColRight>" & SortColRight & "<"
Debug.Print "SortRowEnd>" & SortRowEnd & "<"
Debug.Print "SortCol_1>" & SortCol_1 & "<"
Debug.Print "SortCol_2>" & SortCol_2 & "<"

' SortEnd = Worksheets(Sort_Sheet).Cells(Rows.Count, 1).End(xlUp).Row

Worksheets(Sort_Sheet).Range(SortColStart & SortRowStart & ":" & SortColRight & SortRowEnd).Sort Key1:=Worksheets(Sort_Sheet).Range(SortCol_1 & SortRowStart), Key2:=Worksheets(Sort_Sheet).Range(SortCol_2 & SortRowStart)
End Sub

Function IsSheetPresent(TmpSheet, TmpCheck)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(TmpSheet)   ' Error here & fall thru to next line
    If ws Is Nothing Then
        TmpCheck = 0                ' if ws NOT set then sheet does NOT exist
    Else
        TmpCheck = 1                ' is ws set then sheet DOES exist
    End If
End Function

Sub SpeedTurnOptionsOff()
' Application.ScreenUpdating = False    ' Left on for Status updates on Screen
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.DisplayAlerts = False
End Sub

Sub SpeedTurnOptionsOn()
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.DisplayAlerts = True
End Sub

Function lastRow(ws As Worksheet)
  On Error Resume Next
  lastRow = ws.Cells.Find(What:="*", _
                          After:=ws.Range("A1"), LookAt:=xlPart, LookIn:=xlFormulas, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlPrevious, MatchCase:=False).Row
  On Error GoTo 0
End Function

Sub Reset()
ActiveSheet.UsedRange
End Sub



