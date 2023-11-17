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

