Attribute VB_Name = "Module1"
'Module 1:          Macros for Ribbon buttons
'Author:            Brad Nielsen
'Last Modified:     2/18/2016

Sub TabToYellow()
'
' TabToYellow Macro
'
With ActiveSheet.Tab
        .Color = 65535
        .TintAndShade = 0
    End With
End Sub
Sub TabToRed()
'
' TabToRed Macro
'
With ActiveSheet.Tab
        .Color = 255
        .TintAndShade = 0
    End With
End Sub
Sub TabToGreen()
'
' TabToGreen Macro
'
With ActiveSheet.Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
End Sub
Sub TabToBlue()
'
' TabToBlue Macro
'
With ActiveSheet.Tab
        .Color = 15773696
        .TintAndShade = 0
    End With
End Sub
Sub TabNoColor()
'
' TabNoColor Macro
'
With ActiveSheet.Tab
        .ColorIndex = xlNone
    End With
End Sub
Sub CellToYellow()
'
' Cell To Yellow Macro
'
With ActiveCell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub NoColorCell()
'
' Cell To No Color Macro
'
With ActiveCell.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub CellToRed()
'
' Cell To Red Macro
'
With ActiveCell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub CellsNoColor()
'
' Cells to No Color Macro
'
Dim rng As Range
   Set rng = Selection
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub CellsToYellow()
'
' Cells to Yellow Macro
'
   Dim rng As Range
   Set rng = Selection
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub CellsToRed()
'
' Cells to Red Macro
'
   Dim rng As Range
   Set rng = Selection
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub AutoSizeCol()
'Auto size columns and rows
    Cells.Select
    Cells.EntireRow.AutoFit
    Cells.EntireColumn.AutoFit
    
    Range("A1").Select
End Sub



Sub UsedRange()
'Reset the used range
    ActiveSheet.UsedRange
End Sub


Sub ResetComments()
Dim cmt As Comment
For Each cmt In ActiveSheet.Comments
   cmt.Shape.Top = cmt.Parent.Top + 5
   cmt.Shape.Left = _
      cmt.Parent.Offset(0, 1).Left + 5
Next
End Sub
Sub Comments_AutoSize()
'posted by Dana DeLouis  2000-09-16
Dim MyComments As Comment
Dim lArea As Long
For Each MyComments In ActiveSheet.Comments
  With MyComments
    .Shape.TextFrame.AutoSize = True
    If .Shape.Width > 300 Then
      lArea = .Shape.Width * .Shape.Height
      .Shape.Width = 200
      ' An adjustment factor of 1.1 seems to work ok.
      .Shape.Height = (lArea / 200) * 1.1
    End If
  End With
Next ' comment
End Sub

Sub FixComments()
'realign comments
Dim cmt As Comment
For Each cmt In ActiveSheet.Comments
   cmt.Shape.Top = cmt.Parent.Top + 5
   cmt.Shape.Left = _
      cmt.Parent.Offset(0, 1).Left + 5
Next
'resize comments
Dim MyComments As Comment
Dim lArea As Long
For Each MyComments In ActiveSheet.Comments
  With MyComments
    .Shape.TextFrame.AutoSize = True
    
    If .Shape.Width > 400 Then 'default 300
      lArea = .Shape.Width * .Shape.Height
      .Shape.Width = 350
      ' An adjustment factor of 1.1 seems to work ok.
      .Shape.Height = (lArea / 200) * 0.9
    End If
    
    'set min hgt to 50
    If .Shape.Height < 50 Then
        .Shape.Height = 50
    End If
    
  End With
Next ' comment
   
End Sub
  
Sub textToGreen()
' textToGreen Macro
    With Selection.Font
        .Color = -11489280
        .TintAndShade = 0
    End With
End Sub

Sub BlueHeader()
'
' BlueHeader Macro
'

'
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6299648
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Rows("1:1").RowHeight = 46.5
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        
        
        'grid'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
        
    End With
End Sub


Sub clearHighlight()
'
' clearHighlight Macro
'

'
    Rows("31:31").Select
    Range("F31").Activate
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A1").Select
End Sub
Sub clearLineFmt()
'
' clearLineFmt Macro
'

'
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub symboltoname()
Attribute symboltoname.VB_ProcData.VB_Invoke_Func = " \n14"
'
' symboltoname Macro
'

'

    Selection.ConvertToLinkedDataType ServiceID:=268435456, LanguageCulture:= _
        "en-US"
    
End Sub


Sub setFalseToRed()
Attribute setFalseToRed.VB_ProcData.VB_Invoke_Func = " \n14"
'
' setFalseToRed Macro

Dim MyRange As Range
Set MyRange = Selection


'Selection.FormatConditions.Add Type:=xlTextString, Operator:=xlEqual, Formula1:="FALSE"

With MyRange.FormatConditions.Add(xlTextString, TextOperator:=xlContains, String:="false")


    With .Interior
    .PatternColorIndex = xlAutomatic
    .Color = 255
    .TintAndShade = 0
    End With

End With




End Sub



Sub poamHeaderFix()
Attribute poamHeaderFix.VB_ProcData.VB_Invoke_Func = " \n14"
    On Error GoTo Error
    Sheets("POA&M").Select
    
Error:
    
    Range("C2:H6").Select
    Selection.ClearContents
    Range("J2:K6").Select
    Selection.ClearContents
    
    Range("I2:I6").Select
    Selection.Insert Shift:=xlToRight
    Range("C2:I2").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.ClearContents
    Selection.Copy
    Range("C2:I6").Select
    ActiveSheet.Paste
    Range("C2:I2").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("C3:I3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Copy
    Range("C3:I6").Select
    ActiveSheet.Paste
    Range("C3:I3").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.Copy
    Range("C3:I6").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C3:I3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C4:I4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C5:I5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C6:I6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C2:I6").Select
    Range("C6").Activate
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    Range("N2:P3").Select
    
    
    Range("C3:I3").Select
    ActiveCell.FormulaR1C1 = "Brad Nielsen"
    Range("C4:I4").Select
    ActiveCell.FormulaR1C1 = "CNIC"
    Range("C5:I5").Select
    ActiveCell.FormulaR1C1 = "Navy Electronic Security System - LENEL (Isolated)"
    Range("K2:L3").Select
    ActiveCell.FormulaR1C1 = "Platform IT System"
    Range("K4:L4").Select
    ActiveCell.FormulaR1C1 = "Brad Nielsen"
    Range("K5:L5").Select
    ActiveCell.FormulaR1C1 = "661.444.0141"
    Range("K6:L6").Select
    ActiveCell.FormulaR1C1 = "brad.nielsen@serco-na.com"
    
    
    
    
    Range("P7").Select
    ActiveCell.FormulaR1C1 = _
        "Mitigations (in-house and in conjunction with the Navy CSSP)"
    Range("Q7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Predisposing Conditions Severity"
    Range("R7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Relevance of Threat"
    Range("S7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Threat Description"
    Range("T7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Likelihood"
    Range("U7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Impact"
    Range("V7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Impact Description"
    Range("W7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "Residual Risk Level"
    Range("X7").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = _
        "Recommendations Resulting Residual Risk after Proposed Mitigations"
    Range("A2:B2").Select
    
    ''''''
    
    
    Range("L8").Select
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlUp).Select
    Range("O7").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("O7").Select
    Selection.Copy
    Range("P7:X7").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Columns.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Columns.AutoFit
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("L8").Select
    ActiveCell.SpecialCells(xlLastCell).Select
    Selection.End(xlUp).Select
    Range("O7").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("O7").Select
    Selection.Copy
    Range("P7:X7").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Columns.AutoFit
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Columns.AutoFit
    Range("A2:B2").Select
    Range("A7:B7").Select
End Sub
Sub setDataTypeGeneral()
Attribute setDataTypeGeneral.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim cell As Excel.Range
    'IsEmpty( value )
    
    
    
    For Each cell In Selection
        If IsEmpty(cell.Value) Then GoTo NextIteration
            'Next cell
            'MsgBox ("cell blank")
        'Else

'        End If
    

        With cell
        cell.NumberFormat = "General"
        .Value = .Value
        'MsgBox ("cell not blank")
        End With
NextIteration:
    Next cell
    

End Sub


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

