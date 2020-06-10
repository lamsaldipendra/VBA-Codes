Attribute VB_Name = "WexUpdate"
Sub UpdateWexReport()

With Application
    .ScreenUpdating = False
    .DisplayStatusBar = False
    .DisplayAlerts = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
    .Cursor = xlWait
End With

Dim ReportFolder As FileDialog
Dim fso As Object
Dim NewReportPath As String
Dim NewReport As String
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set ReportFolder = Application.FileDialog(msoFileDialogFilePicker)

    With ReportFolder
        .Title = "Select New Report File"
            If .Show <> -1 Then GoTo NoSel
            Sheet3.Range("T23").Value = .SelectedItems(1)
            NewReportPath = Sheet3.Range("T23").Value
            NewReport = fso.getfilename(NewReportPath)
            Sheet3.Range("T24") = NewReport
    End With


Dim ReportPath As String, ReportFile As String
    ReportPath = Sheet3.Range("T18").Value
    ReportFile = fso.getfilename(ReportPath)

Application.Workbooks.Open (ReportPath)
Application.Workbooks.Open (NewReportPath)

Dim MainRpt As Workbook, MainSht As Worksheet
Dim NewRpt As Workbook, NewSht As Worksheet

Dim mRow As Long, nRow As Long, cRow As Long

    Set MainRpt = Workbooks(ReportFile)
    Set MainSht = MainRpt.Worksheets("Transactions")
    Set NewRpt = Workbooks(NewReport)
    Set NewSht = NewRpt.Worksheets("WEX_SETTLEMENTS__AP_Overcharge")

    mRow = MainSht.Range("A" & MainSht.Rows.Count).End(xlUp).Row + 1
    nRow = NewSht.Range("C" & NewSht.Rows.Count).End(xlUp).Row
    cRow = mRow + nRow

        NewSht.Range("J2:J" & nRow).Copy Destination:=MainSht.Range("A" & mRow & ":A" & cRow)
        NewSht.Range("H2:H" & nRow).Copy Destination:=MainSht.Range("B" & mRow & ":B" & cRow)
        NewSht.Range("D2:D" & nRow).Copy Destination:=MainSht.Range("C" & mRow & ":C" & cRow)
        NewSht.Range("U2:U" & nRow).Copy Destination:=MainSht.Range("D" & mRow & ":D" & cRow)
        NewSht.Range("C2:C" & nRow).Copy Destination:=MainSht.Range("E" & mRow & ":E" & cRow)
        NewSht.Range("B2:B" & nRow).Copy Destination:=MainSht.Range("F" & mRow & ":F" & cRow)
        NewSht.Range("W2:W" & nRow).Copy Destination:=MainSht.Range("G" & mRow & ":G" & cRow)
        NewSht.Range("Q2:Q" & nRow).Copy Destination:=MainSht.Range("H" & mRow & ":H" & cRow)
        NewSht.Range("R2:R" & nRow).Copy Destination:=MainSht.Range("I" & mRow & ":I" & cRow)
        NewSht.Range("Q2:Q" & nRow).Copy Destination:=MainSht.Range("H" & mRow & ":H" & cRow)
        NewSht.Range("F2:F" & nRow).Copy Destination:=MainSht.Range("J" & mRow & ":J" & cRow)
        NewSht.Range("G2:G" & nRow).Copy Destination:=MainSht.Range("K" & mRow & ":K" & cRow)
        NewSht.Range("E2:E" & nRow).Copy Destination:=MainSht.Range("L" & mRow & ":L" & cRow)
        NewSht.Range("I2:I" & nRow).Copy Destination:=MainSht.Range("M" & mRow & ":M" & cRow)
        NewSht.Range("M2:M" & nRow).Copy Destination:=MainSht.Range("N" & mRow & ":N" & cRow)
        NewSht.Range("P2:P" & nRow).Copy Destination:=MainSht.Range("O" & mRow & ":O" & cRow)
        NewSht.Range("V2:V" & nRow).Copy Destination:=MainSht.Range("P" & mRow & ":P" & cRow)

        MainSht.Columns("D").Replace Replacement:="'", What:="XXXX-XXXX-XXXX-", LookAt:=xlPart, SearchOrder:=xlByRows, _
        MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        MainSht.Range("E" & mRow & ":E" & cRow).NumberFormat = "General"
        MainSht.Range("E" & mRow & ":E" & cRow).Value = MainSht.Range("E" & mRow & ":E" & cRow).Value

        MainSht.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(5), Header:=xlYes

Dim i As Long
'Range("G5").Value = CDate(Range("G5").Value)
    For i = 2 To cRow
        With MainSht
            .Range("F" & i).Value = CDate(.Range("F" & i).Value)
            .Range("G" & i).Value = CDate(.Range("G" & i).Value)
        End With
    Next i

        MainSht.Range("I" & mRow & ":I" & cRow).Value = MainSht.Range("I" & mRow & ":I" & cRow).Value
        MainSht.Range("K" & mRow & ":K" & cRow).Value = MainSht.Range("K" & mRow & ":K" & cRow).Value
        MainSht.Range("P" & mRow & ":P" & cRow).Value = MainSht.Range("P" & mRow & ":P" & cRow).Value

'Filter out transactions older than 190 days

Dim dDate As Long
    dDate = Date - 190
        If MainSht.FilterMode = True Then
        MainSht.ShowAllData
        End If
    With MainSht
        .Range("A1").CurrentRegion.AutoFilter Field:=6, Criteria1:="<" & dDate
        .Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With

    MainSht.AutoFilterMode = False

'applies formatt
MainSht.Activate
MainSht.Range("A2:P120000").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 9
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With

MainRpt.Save
mRow = MainSht.Range("A" & MainSht.Rows.Count).End(xlUp).Row

ThisWorkbook.Worksheets("Admin").Range("T21").Value = "Recent Charge on: " & Format(WorksheetFunction.Max(MainSht.Range("F1:F" & mRow)), "dd-Mmm-yy")
MainRpt.Close
NewRpt.Close savechanges:=False

NoSel:

    With Application
    .ScreenUpdating = True
    .DisplayStatusBar = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Cursor = xlDefault
    .Calculation = xlCalculationAutomatic
    End With
End Sub

