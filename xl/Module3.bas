Attribute VB_Name = "Module3"
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
