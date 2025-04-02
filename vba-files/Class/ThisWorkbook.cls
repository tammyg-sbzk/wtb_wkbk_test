VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
  Dim thisMonth As Date, rateMonth As Date, ws As Worksheet
  Dim newFileName
  OKExit = True
  
  If ThisWorkbook.Path = "" Then
    If MsgBox("You must choose a working folder and save this workbook." & vbCr & vbCr _
        & "Do you want to save this workbook now?", vbYesNo, "SAVE NEW WORKBOOK") _
        = vbNo Then                     'If user cancels
        OKExit = False                    'set OKExit to False and Exit Sub
        Exit Sub
    End If                              'Otherwise, proceed with Save As

    ChDrive "F"                       'Keep user away from Template folder
    ChDir "F:\1Net"
    
    newFileName = Application.GetSaveAsFilename(, "Excel File (*.xlsm), *.xlsm")
    
    If newFileName = False Then         'If user cancels here, also
      OKExit = False                    'set OKExit to False and Exit Sub
      Exit Sub
    Else 'If user doesn't cancel, then
      ActiveWorkbook.Sheets("Year-End Questions").Range("Bookkeeper").Value = Split(StrConv(Application.UserName, vbUpperCase), " ")(0) 'add name to bookkeeper box
      ThisWorkbook.SaveAs newFileName, FileFormat:=52   'Save workbook
    End If
End If
End Sub





