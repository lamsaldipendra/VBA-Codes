Attribute VB_Name = "TestModules"
Option Explicit

Sub arrayindex() ' most useful

    ActiveWorkbook.Worksheets("Test").Cells.Clear

    Dim arrayData(1 To 5) As Variant
    arrayData(1) = "A"
    arrayData(2) = "B"
    arrayData(3) = "C"
    arrayData(4) = "D"
    arrayData(5) = "E"

    Dim rngTarget As Range
    Set rngTarget = ActiveWorkbook.Worksheets("Test").Range("A1:E1")
    rngTarget = arrayData

End Sub

Sub testcolumn()

    ActiveWorkbook.Worksheets("Test").Cells.Clear

    Dim arrayData(1 To 1, 1 To 5) As Variant
    arrayData(1, 1) = "A"
    arrayData(1, 2) = "B"
    arrayData(1, 3) = "C"
    arrayData(1, 4) = "D"
    arrayData(1, 5) = "E"

    Dim rngTarget As Range
    Set rngTarget = ActiveWorkbook.Worksheets("Test").Range("A1:E1")
    rngTarget = arrayData

End Sub

Sub testrow()

MsgBox ThisWorkbook.Worksheets("WEXCB").Range("A" & ThisWorkbook.Worksheets("WEXCB").Rows.Count).End(xlUp).Row

End Sub

Sub Cbm_Value_Select()
   'Set up the variables.
   Dim rng As Range
   
   'Use the InputBox dialog to set the range for MyFunction, with some simple error handling.
   Set rng = Application.InputBox("Range:", Type:=8)
   If rng.Cells.Count <> 3 Then
     MsgBox "Length, width and height are needed -" & _
         vbLf & "please select three cells!"
      Exit Sub
   End If
   
   'Call MyFunction by value using the active cell.
   ActiveCell.Value = MyFunction(rng)
End Sub

Function MyFunction(rng As Range) As Double
   MyFunction = rng(1) * rng(2) * rng(3)
End Function

