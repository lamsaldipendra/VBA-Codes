Attribute VB_Name = "AddComments"
Function FileLocked(strFileName As String) As Boolean

On Error Resume Next

' If the file is already opened by another process,
' and the specified type of access is not allowed,
' the Open operation fails and an error occurs.
Open strFileName For Binary Access Read Lock Read As #1
Close #1

' If an error occurs, the document is currently open.
If Err.Number <> 0 Then
    FileLocked = True
    Err.Clear
End If

End Function

Sub UpdateComments()

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim ActionWb As Workbook
Dim ActionSht As Worksheet, LocalDataSht As Worksheet
Dim Sr As Integer, Dr As Integer, Fr As Integer
Dim h As String, NewComment As String, uName As String, uInitials As String, addcom As String
Dim tDate As Variant
Dim DataWb As Workbook
Dim DataSht As Worksheet
Dim DatabasePath As String
Dim DataFile As String

    Set ActionWb = ThisWorkbook
    Set ActionSht = ThisWorkbook.Worksheets("2003VCC")
    Set LocalDataSht = ActionWb.Worksheets("2003VCCDb")
    DatabasePath = ActionWb.Worksheets("Admin").Range("T6").Value
    DataFile = ActionWb.Worksheets("Admin").Range("T9").Value
    
    Sr = ActionSht.Range("E1").Value
    Dr = ActionSht.Range("E2").Value
    Fr = ActionSht.Range("E3").Value
    h = ActionSht.Cells(Sr, 25).Value ' reads existing value

    uName = Application.Username
    If uName = "Maria Almendras" Then
    uName = "Raelyne Almendras"
    End If
    uInitials = Left(uName, 1) & Mid(uName, InStr(uName, " ") + 1, 1)
    tDate = Format(Date, "dd-mmm-yy")

    NewComment = InputBox("Add new comment here:" & Chr(10) & h) ' prompts input box
TryAgain:
    If NewComment <> "" Then ' end macro if user do not enter any comment
    addcom = tDate & Space(1) & "-" & Space(1) & uInitials & Space(1) & "-" & Space(1) & UCase(NewComment) ' additional comment
    
    
    If Not FileLocked(DatabasePath) Then
        Workbooks.Open FileName:=DatabasePath, UpdateLinks:=0, ReadOnly:=False
    Else
    Dim answer As Integer
    answer = MsgBox("Click Retry to resubmit or Cancel to undo the changes", vbRetryCancel)
    If answer = vbRetry Then
    GoTo TryAgain
    Else
    MsgBox "Change is not saved. Please do it again"
    GoTo UndoChange
    End If
    End If
    
    Set DataWb = Workbooks(DataFile)
    Set DataSht = DataWb.Worksheets("2003VCCDb")
    
    If DataSht.Range("S" & Dr) <> "" Then ' Updating comment
    
    With DataSht
    .Range("S" & Dr).Value = h & Chr(10) & addcom
    .Range("A" & Dr & ":Y" & Dr).WrapText = False
    DataWb.Save
    DataWb.Close
    End With
    
    With ActionSht
    .Range("T4").Value = h & Chr(10) & addcom
    .Range("Y" & Sr).Value = h & Chr(10) & addcom
    .Range("G" & Sr & ":AE" & Sr).WrapText = False
    End With
    
    If Fr <> 0 Then
    With LocalDataSht
    .Range("S" & Fr).Value = h & Chr(10) & addcom
    .Range("A" & Fr & ":Y" & Dr).WrapText = False
    End With
    End If
    
    Else
    
    With DataSht
    .Range("S" & Dr).Value = addcom
    .Range("R" & Dr).Value = tDate
    .Range("A" & Dr & ":Y" & Dr).WrapText = False
    DataWb.Save
    DataWb.Close
    End With
    
    With ActionSht
    .Range("T4").Value = addcom
    .Range("Y" & Sr).Value = addcom
    .Range("X" & Sr).Value = tDate
    .Range("G" & Sr & ":AE" & Sr).WrapText = False
    End With
    
    If Fr <> 0 Then
    With LocalDataSht
    .Range("S" & Fr).Value = addcom
    .Range("R" & Fr).Value = tDate
    .Range("A" & Fr & ":Y" & Dr).WrapText = False
    End With
    End If
    
    End If

    End If ' ending first if statement
UndoChange:

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub StartEditDetails()

    With Sheet1
    .Range("C11") = True
    'edit button
    .Shapes("Rounded Rectangle 4").Visible = False
    .Shapes("Picture 20").Visible = False
    
    'filter buttons top to bottom
    .Shapes("Rounded Rectangle 6").Visible = False
    .Shapes("Picture 28").Visible = False
    .Shapes("Rounded Rectangle 5").Visible = False
    .Shapes("Picture 13").Visible = False
    .Shapes("Rounded Rectangle 7").Visible = False
    .Shapes("Picture 17").Visible = False
    
    ' refresh button
    .Shapes("Rounded Rectangle 11").Visible = False
    .Shapes("Picture 31").Visible = False
    
    'drill button
    If .Range("C5") = False Then
    .Shapes("Rounded Rectangle 10").Visible = False
    .Shapes("Picture 18").Visible = False
    Else:
    .Shapes("Rounded Rectangle 12").Visible = False
    .Shapes("Picture 44").Visible = False
    End If
    
    'updatecomment button
    .Shapes("Rounded Rectangle 8").Visible = False
    .Shapes("Picture 33").Visible = False
    
    ' enabaling save and cancel buttons
    .Shapes("Rounded Rectangle 50").Visible = True
    .Shapes("Picture 24").Visible = True
    .Shapes("Rounded rectangle 52").Visible = True
    .Shapes("Picture 22").Visible = True
    End With
End Sub

Sub SaveEditDetails()
' declaring this workbook and its properties
Dim ActionWb As Workbook
Dim ActionSht As Worksheet
Dim LocalDataSht As Worksheet
Dim Sr As Integer, Dr As Integer, Fr As Integer
Dim Channel As String, GlAuth As String, JiraCase As String, Status As String, sDate As String, Person As String, Started As String, Comment As String, Issue As String
Dim DatabasePath As String
Dim DataFile As String

' setting properties
Set ActionWb = ThisWorkbook
Set ActionSht = ThisWorkbook.Worksheets("2003VCC")
Set LocalDataSht = ActionWb.Worksheets("2003VCCDb")
    
    Sr = ActionSht.Range("E1").Value
    Dr = ActionSht.Range("E2").Value
    Fr = ActionSht.Range("E3").Value

    DatabasePath = ActionWb.Worksheets("Admin").Range("T6").Value
    DataFile = ActionWb.Worksheets("Admin").Range("T9").Value

' setting Labels
    Channel = ActionSht.Range("R5").Value
    GlAuth = ActionSht.Range("R6").Value
    JiraCase = ActionSht.Range("R8").Value
    Status = ActionSht.Range("R9").Value
    Person = ActionSht.Range("R11").Value
    sDate = ActionSht.Range("R12").Value
    Comment = ActionSht.Range("T4").Value
    Issue = ActionSht.Range("Y4").Value
    
'declaring database workbook and its properties
Dim DataWb As Workbook
Dim DataSht As Worksheet

If ActionSht.Range("C11") = True Then ' extra check before writing changes
TryAgain:
Application.ScreenUpdating = False
    If Not FileLocked(DatabasePath) Then
    Workbooks.Open FileName:=DatabasePath, UpdateLinks:=0, ReadOnly:=False
    Else
    Dim answer As Integer
    answer = MsgBox("Click Retry to resubmit or Cancel to undo the changes", vbRetryCancel)
    If answer = vbRetry Then
    GoTo TryAgain
    Else
    MsgBox "Change is not saved. Please do it again"
    GoTo UndoChange
    End If
    End If
    
    Set DataWb = Workbooks(DataFile)
    Set DataSht = DataWb.Worksheets("2003VCCDb")

   ''' On Error Resume Next
  
  ' Database
  
    With DataSht
    .Range("T" & Dr).Value = Channel
    .Range("W" & Dr).Value = GlAuth
    .Range("V" & Dr).Value = JiraCase
    .Range("X" & Dr).Value = Status
    .Range("Q" & Dr).Value = Person
    .Range("R" & Dr).Value = sDate
    .Range("S" & Dr).Value = Comment
    .Range("U" & Dr).Value = Issue
    
    .Range("A" & Dr & ":Y" & Dr).WrapText = False
    End With
    
    DataWb.Save
    DataWb.Close
    
    ' Action Sheet
    
    With ActionSht
    .Range("Z" & Sr).Value = Channel
    .Range("AC" & Sr).Value = GlAuth
    .Range("AB" & Sr).Value = JiraCase
    .Range("AD" & Sr).Value = Status
    .Range("W" & Sr).Value = Person
    .Range("X" & Sr).Value = sDate
    .Range("Y" & Sr).Value = Comment
    .Range("AA" & Sr).Value = Issue
    
    .Range("G" & Sr & ":AE" & Sr).WrapText = False
    End With
    
    ' Local Database
    If Fr <> 0 Then
    With LocalDataSht
    .Range("T" & Fr).Value = Channel
    .Range("W" & Fr).Value = GlAuth
    .Range("V" & Fr).Value = JiraCase
    .Range("X" & Fr).Value = Status
    .Range("Q" & Fr).Value = Person
    .Range("R" & Fr).Value = sDate
    .Range("S" & Fr).Value = Comment
    .Range("U" & Fr).Value = Issue
    
    .Range("A" & Dr & ":Y" & Dr).WrapText = False
    End With
    End If
    
   ' Buttons
    ' hiding save and cancel buttons
    With Sheet1
    .Shapes("Rounded Rectangle 50").Visible = False
    .Shapes("Picture 24").Visible = False
    .Shapes("Rounded rectangle 52").Visible = False
    .Shapes("Picture 22").Visible = False
    
    'edit button
    .Shapes("Rounded Rectangle 4").Visible = True
    .Shapes("Picture 20").Visible = True
    
    ' refresh button
    .Shapes("Rounded Rectangle 11").Visible = True
    .Shapes("Picture 31").Visible = True
    
    'drill button
    If .Range("C5") = False Then
    .Shapes("Rounded Rectangle 10").Visible = True
    .Shapes("Picture 18").Visible = True
    'and filter buttons top to bottom
    .Shapes("Rounded Rectangle 6").Visible = True
    .Shapes("Picture 28").Visible = True
    .Shapes("Rounded Rectangle 5").Visible = True
    .Shapes("Picture 13").Visible = True
    .Shapes("Rounded Rectangle 7").Visible = True
    .Shapes("Picture 17").Visible = True
    Else:
    .Shapes("Rounded Rectangle 12").Visible = True
    .Shapes("Picture 44").Visible = True
    End If
    
    'updatecomment button
    .Shapes("Rounded Rectangle 8").Visible = True
    .Shapes("Picture 33").Visible = True

    .Range("C11") = False
    End With
    
    Application.ScreenUpdating = True
    End If
    
UndoChange:
End Sub
    

Sub Cancelbtn()

    Application.ScreenUpdating = False

     ' Buttons
    ' hiding save and cancel buttons
    With Sheet1
    .Shapes("Rounded Rectangle 50").Visible = False
    .Shapes("Picture 24").Visible = False
    .Shapes("Rounded rectangle 52").Visible = False
    .Shapes("Picture 22").Visible = False
    
    'edit button
    .Shapes("Rounded Rectangle 4").Visible = True
    .Shapes("Picture 20").Visible = True
    
    ' refresh button
    .Shapes("Rounded Rectangle 11").Visible = True
    .Shapes("Picture 31").Visible = True
    
    'drill button
    If .Range("C5") = False Then
    .Shapes("Rounded Rectangle 10").Visible = True
    .Shapes("Picture 18").Visible = True
    'and filter buttons top to bottom
    .Shapes("Rounded Rectangle 6").Visible = True
    .Shapes("Picture 28").Visible = True
    .Shapes("Rounded Rectangle 5").Visible = True
    .Shapes("Picture 13").Visible = True
    .Shapes("Rounded Rectangle 7").Visible = True
    .Shapes("Picture 17").Visible = True
    Else:
    .Shapes("Rounded Rectangle 12").Visible = True
    .Shapes("Picture 44").Visible = True
    End If
    
    'updatecomment button
    .Shapes("Rounded Rectangle 8").Visible = True
    .Shapes("Picture 33").Visible = True

    .Range("C11") = False
    End With
    
    Call LoadHeader
    
    Application.ScreenUpdating = True

End Sub

Sub ExpandComment()
    
End Sub

Sub HoldComment()

End Sub
Sub PushChange()

End Sub

