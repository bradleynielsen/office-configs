Attribute VB_Name = "Module5"
Sub splitcell()
    'splits Text active cell using ALT+10 char as separator
    Dim splitVals As Variant
    Dim totalVals As Long
    Dim i As Integer
 
    For i = 1 To 1000
      splitVals = Split(ActiveCell.Value, Chr(10))
      totalVals = UBound(splitVals)
      Range(Cells(ActiveCell.Row, ActiveCell.Column + 1), Cells(ActiveCell.Row, ActiveCell.Column + 1 + totalVals)).Value = splitVals
      ActiveCell.Offset(1, 0).Activate
    Next i
End Sub

Sub move_pub_date()
'
' move_pub_date Macro
'

'

' copy and split date into col
    Columns("W:W").Select
    Selection.Copy
    Range("W1").Select
    Selection.End(xlToRight).Select
    Range("AS1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AS1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
        
    Range("AV2").Select
    ActiveCell.FormulaR1C1 = "=MONTH(DATEVALUE(RC[-3]&"" 1, 1970""))"
    Range("AV3").Select
    
    'good
    
    Range("AW2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=DATE(RC[-2],RC[-1],RC[-3])"
    Range("AW3").Select
    
        Range("AW1").Select
    Selection.Locked = True
    Selection.FormulaHidden = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "Plugin Publication Date"
    Range("AW2").Select
    
    
    
    
    '   autofil to the end
    Dim TotalRow As Integer
    TotalRow = ActiveSheet.UsedRange.Rows.Count
    Dim theAddress As String
    theAddress = "AV2:AW" & CStr(TotalRow)
    Range("AV2:AW2").Select
    Selection.AutoFill Destination:=Range(theAddress)
    
    Columns("AW:AW").EntireColumn.AutoFit
    
    Range("AV1").Select
    ActiveCell.FormulaR1C1 = "Month"
    Range("AV2").Select
    Selection.AutoFilter
    
End Sub

