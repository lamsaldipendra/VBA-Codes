Attribute VB_Name = "Browse"
Option Explicit

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

Call MatchFormula


NoSel:

End With
End Sub

Sub BrowseWex()

Dim FileSelector As FileDialog
Dim fso As Object
Dim DatabasePath As String
Dim FileName As String
Set fso = CreateObject("Scripting.FileSystemObject")

Set FileSelector = Application.FileDialog(msoFileDialogFilePicker)

With FileSelector

.Title = "Select Data File"
If .Show <> -1 Then GoTo NoSel
Sheet3.Range("T18").Value = .SelectedItems(1)
DatabasePath = Sheet3.Range("T18").Value
FileName = fso.getfilename(DatabasePath)

Sheet3.Range("T19") = FileName

NoSel:

End With
End Sub


Sub MatchFormula()
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim Path As String, File As String

Path = fso.GetParentFolderName(Sheet3.Range("T6").Value) & "\"
File = Sheet3.Range("T9").Value

Sheet1.Range("E2").Formula = "=IFERROR(MATCH(I4,'" & Path & "[" & File & "]2003VCCDb'!$A:$A,0),"""")"

End Sub

Sub MaxFormula()


End Sub

