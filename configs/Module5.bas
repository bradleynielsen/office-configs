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
