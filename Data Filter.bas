Attribute VB_Name = "Filter"
Option Explicit

Sub AdvancedFilter()
    Application.ScreenUpdating = False
    
    Range("I4:I13,L4:L13,O4:O13,R5:R12").ClearContents
    Range("T4:X11,Y4:Z11").ClearContents
    
   ''' On Error Resume Next
    Application.StatusBar = "Applying initial Filter"
    Call FirstFilter
    Application.StatusBar = "Importing data to Front Page"
    Call ToFront
    Range("G19:AE5000").WrapText = False
    Call LoadHeader
    Application.StatusBar = ""
    
    Application.ScreenUpdating = True

End Sub


Sub FirstFilter()

    Application.ScreenUpdating = False
    Dim DataWb As Workbook, DataSht As Worksheet
    Dim ActionWb As Workbook, ActionSht As Worksheet, LocalDataSht As Worksheet
    Dim DataRange As Range, Criteria As Range, Output As Range
    
    Set ActionWb = ThisWorkbook
    Set ActionSht = ActionWb.Worksheets("2003VCC")
    Set LocalDataSht = ActionWb.Worksheets("2003VCCDb")
    
    Dim DatabasePath As String
    Dim DataFile As String
    DatabasePath = ActionWb.Worksheets("Admin").Range("T6").Value
    
    LocalDataSht.Range("A2:Y5000").ClearContents
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Set DataWb = Application.Workbooks.Open(DatabasePath)
    
    Set DataRange = DataWb.Worksheets("2003VCCDb").Range("A1").CurrentRegion
    Set Criteria = LocalDataSht.Range("AG1:AO6")
    Set Output = LocalDataSht.Range("A1:Y1")

    DataRange.AdvancedFilter xlFilterCopy, Criteria, Output
    Application.DisplayAlerts = False
    DataWb.Close savechanges:=False

    Application.ScreenUpdating = True

End Sub


Sub ToFront()
    Dim ActionWb As Workbook, ActionSht As Worksheet, LocalDataSht As Worksheet
    Set ActionWb = ThisWorkbook
    Set ActionSht = ActionWb.Worksheets("2003VCC")
    Set LocalDataSht = ActionWb.Worksheets("2003VCCDb")
    
    Application.ScreenUpdating = False
    On Error Resume Next
    Dim Data As Range, Criteria As Range, Output As Range
    ActionSht.Range("G20:AE5000").ClearContents
    Set Data = LocalDataSht.Range("A1").CurrentRegion
    Set Criteria = LocalDataSht.Range("AR1:AR24")
    Set Output = ActionSht.Range("G19:AE19")

    Data.AdvancedFilter xlFilterCopy, Criteria, Output
    Application.ScreenUpdating = True

End Sub

Sub ClearFilter()
    Sheet1.Range("E15:E25").ClearContents
    Sheet2.Range("AF2:AF6,AQ2:AQ24").Value = True
End Sub

Sub AllStatus()

    Application.ScreenUpdating = False
    If Sheet2.Range("AQ26").Value = True Then
    Sheet2.Range("AQ2:AQ24").Value = False
    Sheet2.Range("AQ26").Value = False
    GoTo Done
    End If
    
    If Sheet2.Range("AQ26").Value = False Then
    Sheet2.Range("AQ26").Value = True
    Sheet2.Range("AQ2:AQ24").Value = True
    End If
Done:
    Application.ScreenUpdating = True

End Sub

Sub DrillDown()

    Application.ScreenUpdating = False

    If Sheet1.Cells(5, 9) <> "" Or Sheet1.Cells(6, 9) <> "" Then
    Application.StatusBar = "Creditor and Booking Filter"
    
    Sheet1.Range("C5").Value = True
    With Sheet1
    .Shapes("Rounded Rectangle 10").Visible = False
    .Shapes("Picture 18").Visible = False
    .Shapes("Rounded Rectangle 7").Visible = False
    .Shapes("Picture 17").Visible = False
    .Shapes("Rounded Rectangle 5").Visible = False
    .Shapes("Picture 13").Visible = False
    .Shapes("Rounded Rectangle 6").Visible = False
    .Shapes("Picture 28").Visible = False
    .Shapes("Rounded Rectangle 12").Visible = True
    .Shapes("Picture 44").Visible = True
    .Shapes("Rounded Rectangle 25").Visible = True
    End With
    Call DrillFilter
    
    Else: MsgBox "No data available to apply filter"
    End If
    
    Application.ScreenUpdating = True

    
End Sub

Sub DrillFilter()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim X() As Variant
    Dim XT() As Variant

    X = Sheet1.Range("I5:I6").Value2
    XT = Application.Transpose(X)
    Sheet2.Range("AG23:AH23").Value = XT
    
    ' On Error Resume Next
    Application.StatusBar = "Importing Data to Front Page"
    
    Dim DataWb As Workbook, DataSheet As Worksheet
    Dim ActionWb As Workbook, ActionSht As Worksheet, LocalDataSht As Worksheet
    Dim DataRange As Range, Criteria As Range, Output As Range
    
    Set ActionWb = ThisWorkbook
    Set ActionSht = ActionWb.Worksheets("2003VCC")
    Set LocalDataSht = ActionWb.Worksheets("2003VCCDb")
    
    Dim DatabasePath As String
    Dim DataFile As String
    DatabasePath = ActionWb.Worksheets("Admin").Range("T6").Value
    DataFile = ActionWb.Worksheets("Admin").Range("T9").Value
    
    Sheet1.Range("G20:AE5000").ClearContents
    Application.ScreenUpdating = False
    Set DataWb = Application.Workbooks.Open(DatabasePath)
    Set DataRange = DataWb.Worksheets("2003VCCDb").Range("A1").CurrentRegion
    Set Criteria = Sheet2.Range("AG22:AH23")
    Set Output = Sheet1.Range("G19:AE19")

    DataRange.AdvancedFilter xlFilterCopy, Criteria, Output
    Application.DisplayAlerts = False
    DataWb.Close savechanges:=False

    Application.StatusBar = ""
    
    Application.ScreenUpdating = False


End Sub

Sub CloseDrill()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Reapplying user selected filters"

    Sheet1.Range("C5").Value = False
    With Sheet1
    .Shapes("Rounded Rectangle 12").Visible = False
    .Shapes("Picture 44").Visible = False
    .Shapes("Rounded Rectangle 25").Visible = False
    .Shapes("Rounded Rectangle 10").Visible = True
    .Shapes("Picture 18").Visible = True
    .Shapes("Rounded Rectangle 7").Visible = True
    .Shapes("Picture 17").Visible = True
    .Shapes("Rounded Rectangle 5").Visible = True
    .Shapes("Picture 13").Visible = True
    .Shapes("Rounded Rectangle 6").Visible = True
    .Shapes("Picture 28").Visible = True
    
    End With
    
    Sheet1.Range("G20:AE5000").ClearContents
    Application.StatusBar = "Importing data to Front Page"
    Call ToFront
    Application.StatusBar = ""

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub Refresh()

Application.ScreenUpdating = False
    If Cells(5, 3).Value = False Then
    Application.ScreenUpdating = False
    Call AdvancedFilter
    Call LoadHeader
    Application.ScreenUpdating = True
    End If
    
    If Cells(5, 3).Value = True Then

    Application.ScreenUpdating = False
    Call FirstFilter
    Call DrillFilter
    Call LoadHeader
    Application.ScreenUpdating = True
    
    End If
Application.ScreenUpdating = True
End Sub

Sub ClearHeader()
Sheet1.Range("I4:I13,l4:L13,O4:O13,R5:R12,T4:X11,Y4:Z11").ClearContents
End Sub
