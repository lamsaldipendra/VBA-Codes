            '''''Function to check if the file is opened and locked by another user'''''''

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


                '''''' Code to perform the filelock check '''''''

TryAgain:
If Not FileLocked("Filenamewithpath") Then
Workbooks.Open FileName:="Filenamewithpath", UpdateLinks:=0, ReadOnly:=False
    Else
    Dim answer As Integer
''' if file is locked it pops msg box to try. again quit the process
    answer = MsgBox("Click Retry to resubmit or Cancel to undo the changes", vbRetryCancel)
    If answer = vbRetry Then
    GoTo TryAgain
    Else
    MsgBox "Change is not saved. Please do it again"
    GoTo UndoChange
    End If
    End If


            '''''' Code to open file select browser and allow to select file '''''''

Sub BrowseDataFile()

Dim DataFile As FileDialog
Dim fso As Object
Dim DatabasePath As String
Dim FileName As String
Set fso = CreateObject("Scripting.FileSystemObject")

Set DataFile = Application.FileDialog(msoFileDialogFilePicker)

With DataFile

.Title = "Select Data File"
If .Show <> -1 Then GoTo NoSel
Sheet3.Range("T6").Value = .SelectedItems(1)
DatabasePath = Sheet3.Range("T6").Value
FileName = fso.getfilename(DatabasePath)

With Sheet3
.Range("T9") = FileName
.Range("U11") = "Mapped on: " & Format(Now, "dd-Mmm-yy, hh:mm:ss AM/PM")
End With
    
NoSel:

End With
End Sub
    
''''''' Advanced Filter function'''''''''

Sub FilterData()

Application.ScreenUpdating = False

            Set SourceBook = Application.Workbooks.Open("workbooktoimportfromwithPath")
            Set SourceRange = SourceBook.Worksheets("SheetName").Range("CellAddress").CurrentRegion
            Set Criteria = CriteriaSheet.Range("CriteriaRange")
            Set Output = Outputsheet.Range("OutputRange")

            SourceRange.AdvancedFilter xlFilterCopy, Criteria, Output
            
            SourceBook.Close savechanges:=False
                               
    
Application.ScreenUpdating = True

End Sub
