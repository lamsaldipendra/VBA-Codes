Attribute VB_Name = "CardFilter"

Option Explicit

'this is working
Sub WexTransactionsFilter()

Application.ScreenUpdating = False
    Dim ReportPath As String, ReportName As String, ReportWb As Workbook
    Dim ActionWb As Workbook, ActionSht As Worksheet, LocalDataSht As Worksheet
    Dim DataRange As Range, Criteria As Range, Output As Range
    Dim LastTrackerRow As Integer, CardHeaderRow As Integer, FirstCardRow As Integer, LastCardRow As Integer
    
    Set ActionWb = ThisWorkbook
    Set ActionSht = ActionWb.Worksheets("2003VCC")
    Set LocalDataSht = ActionWb.Worksheets("2003VCCDb")
    
        ReportPath = ActionWb.Worksheets("Admin").Range("T18").Value
        ReportName = ActionWb.Worksheets("Admin").Range("T19").Value
        
        LastTrackerRow = ActionSht.Range("G19").CurrentRegion.Rows.Count + 19
        
        CardHeaderRow = LastTrackerRow + 2
    
        ActionSht.Range("G" & LastTrackerRow + 1 & ":AE5000").ClearContents
    
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.StatusBar = "Wait while data is being filtered and imported"
    
        Set ReportWb = Application.Workbooks.Open(ReportPath)
        Set DataRange = ReportWb.Worksheets("Transactions").Range("A1").CurrentRegion
        Set Criteria = LocalDataSht.Range("AG26:AH27")
        Set Output = ActionSht.Range("H" & CardHeaderRow)

    DataRange.AdvancedFilter xlFilterCopy, Criteria, Output
Application.DisplayAlerts = False
    ReportWb.Close savechanges:=False
    
    LastCardRow = ActionSht.Range("H" & ActionSht.Rows.Count).End(xlUp).Row
    
    ActionSht.Range("G" & CardHeaderRow).Value = "CARD"

    ActionSht.Range("G" & CardHeaderRow + 1 & ":G" & LastCardRow).Value = "CHARGES"
Application.StatusBar = ""
Application.ScreenUpdating = True

End Sub
