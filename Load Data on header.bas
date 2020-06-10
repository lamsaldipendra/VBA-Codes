Attribute VB_Name = "LoadData"
Option Explicit

Sub LoadHeader()
Dim tRow As Integer
Dim Invoice As String

    tRow = Range("E1").Value
    
    If Cells(tRow, 7).Value <> "" And Cells(tRow, 7).Value <> "CARD" And Cells(tRow, 7).Value <> "CHARGES" Then
    
    Range("I4:I13,L4:L13,R5:R12").ClearContents
    Range("T4:X11,Y4:Z11").ClearContents
    
    'top left table
    Range("I4").Value = Range("G" & tRow).Value 'Unique Id
    Range("I5").Value = Range("K" & tRow).Value 'Creditor code
    Range("I6").Value = Range("O" & tRow).Value 'Booking
    Range("I7").Value = Range("U" & tRow).Value 'HOT
    
    'bottom left table
    Range("I10").Value = Range("L" & tRow).Value 'invoice
    Range("I11").Value = Range("N" & tRow).Value 'currency
    Range("I12").Value = Range("P" & tRow).Value 'Local currency amount
    Range("I13").Value = Range("Q" & tRow).Value 'AUD value
    
    Invoice = Range("L" & tRow).Value
    'top right
    'Card Provider
    If Invoice <> "" Then
        If Left(Invoice, 1) >= 6 Then
            Range("L4").Value = "WEX"
                ElseIf Left(Invoice, 1) <= 2 Then
                    Range("L4").Value = "ENETT"
        End If
    End If
    
    Range("L5").Value = Range("R" & tRow) 'batch
    Range("L6").Value = Range("I" & tRow) 'year
    Range("L7").Value = Range("J" & tRow) 'period
    
    'bottom right
    Range("L10").Value = Range("M" & tRow) 'Invoice Date
    Range("L11").Value = Range("S" & tRow) 'Batch
    Range("L12").Value = Range("V" & tRow) 'Net AUD
    Range("L13").Value = Range("AE" & tRow) 'Credit
    
    'Far right tables
    Range("R5").Value = Cells(tRow, 26)
    Range("R6").Value = Cells(tRow, 29)
    Range("R8").Value = Cells(tRow, 28)
    Range("R9").Value = Cells(tRow, 30)
    Range("R11").Value = Cells(tRow, 23)
    Range("R12").Value = Cells(tRow, 24)
    
    Range("T4").Value = Cells(tRow, 25).Value
    Range("Y4").Value = Cells(tRow, 27).Value
    
    End If
    
    On Error Resume Next
    If Cells(tRow, 7).Value <> "" And Cells(tRow, 7).Value = "CHARGES" Then
        Dim Charged As Long
        Dim Net As Integer
        Dim LastRow As Integer, cRow As Integer
        Dim Card As String
        Dim CriteriaRange As Range, SumRange As Range
    
        cRow = Sheet1.Range("G19").CurrentRegion.Rows.Count + 22
        LastRow = Range("G" & Rows.Count).End(xlUp).Row
    
        Range("O4:O6,O8:O9,O11:O13").ClearContents
        Range("O4").Value = Range("S" & tRow).Value
        Range("O5").Value = Range("T" & tRow).Value
        Range("O6").Value = Range("U" & tRow).Value
        Range("O8").Value = Range("J" & tRow).Value
        Range("O9").Value = Range("K" & tRow).Value
    
        Range("O11").Value = Range("W" & tRow).Value * 0.99
    
        Set CriteriaRange = Range("J" & cRow & ":J" & LastRow)
        Set SumRange = Range("P" & cRow & ":P" & LastRow)
        Card = Range("O8").Value
        Charged = Application.SumIf(CriteriaRange, Card, SumRange)
        Range("O12").Value = Charged
    
        Net = Charged - Range("O11").Value
    
        Range("O13").Value = Net
    
    End If

End Sub







