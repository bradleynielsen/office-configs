Attribute VB_Name = "Module7"
Sub dadmsIndexMaker()
'
' dadmsIndexMaker Macro
' Make Abbreviated DADMS extract
'

'make new sheet
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "abbreviatedExtract"
    
'copy selected DADMS info into new sheet
    Sheets("APP_EXTRACT").Select
    Columns("A:A").Select
    Selection.Copy
    Sheets("abbreviatedExtract").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    
    Sheets("APP_EXTRACT").Select
    Columns("B:B").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("abbreviatedExtract").Select
    Columns("B:B").Select
    ActiveSheet.Paste
    
    Sheets("APP_EXTRACT").Select
    Columns("E:E").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("abbreviatedExtract").Select
    Columns("C:C").Select
    ActiveSheet.Paste
    
    Range("D1").Select
    Sheets("APP_EXTRACT").Select
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("abbreviatedExtract").Select
    Columns("D:D").Select
    ActiveSheet.Paste
    
    Sheets("APP_EXTRACT").Select
    Columns("P:P").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("abbreviatedExtract").Select
    Columns("E:E").Select
    ActiveSheet.Paste
    
    Sheets("APP_EXTRACT").Select
    Columns("AU:AU").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("abbreviatedExtract").Select
    Columns("F:F").Select
    ActiveSheet.Paste
    

    
    
    
    
    
'save file
    'Application.Dialogs(xlDialogSaveAs).Show
    'Workbooks.Open Filename:=ThisWorkbook.Path & "\APP_EXTRACT.xlsx"
        
    
   Dim sFilename As String
    'You can give a nem to save
    sFilename = ActiveWorkbook.Path & "\" & "" & "abbreviatedExtract.xls"
    
    'Workbooks.Add
    'Saving the Workbook
    'MsgBox (sFilename)
    ActiveWorkbook.SaveAs sFilename
    
    
    
    
'remove extra data from xls
    
    Application.DisplayAlerts = False
    Sheets("APP_EXTRACT").Delete
    Application.DisplayAlerts = True
    
    Range("A1:F1").Select
    Selection.Columns.AutoFit
    
    ActiveWorkbook.Save
    
End Sub

