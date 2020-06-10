Attribute VB_Name = "WorkbookUpdate"
Option Explicit

Sub Auto_Open()

Call UpdateWorkbook
End Sub

Sub UpdateWorkbook()

Application.ScreenUpdating = False

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Dim ThisVersion As String, NewVersion As String
Dim workbookPath As String, workbookFolder As String
Dim Macroworkbook As Workbook, Adminsheet As Worksheet
Dim Updater As Workbook, UpdaterSheet As Worksheet

Set Macroworkbook = ThisWorkbook
Set Adminsheet = Macroworkbook.Worksheets("Admin")

Adminsheet.Range("AD5").Formula = "='T:\FCGF\FIN\FCGP only\General\ACCOUNTS\2 SUPPLIER ACCOUNTS\VCC\VCC PROJECT\2. 2003VCC WORKING FILE\2003VCC - Macro Project\Progress\[2003VCC Macro Version Management.xlsx]Version Info'!$B$5"

workbookFolder = Application.ThisWorkbook.Path
workbookPath = Application.ThisWorkbook.FullName

ThisVersion = "2003VCC - Macro Workbook V1.0.3.xlsm"
NewVersion = Adminsheet.Range("AD5")

Dim answer As Integer

If ThisVersion <> NewVersion Then

answer = MsgBox("New Version is Available. Do you want to update it" & Chr(10) & Chr(10) & "Click Yes to proceed or click No to close the workbook", vbYesNo)

If answer = vbYes Then

fso.CopyFile "T:\FCGF\FIN\FCGP only\General\ACCOUNTS\2 SUPPLIER ACCOUNTS\VCC\VCC PROJECT\2. 2003VCC WORKING FILE\2003VCC - Macro Project\MacroBook\" & NewVersion, workbookFolder & "\"
    
    Workbooks.Open (workbookFolder & "\" & NewVersion)
    
    Macroworkbook.Save
    If Dir(workbookFolder & "\Previous Versions", vbDirectory) = vbNullString Then
    fso.CreateFolder (workbookFolder & "\Previous Versions")
    End If
    Application.DisplayAlerts = False
    Macroworkbook.SaveAs workbookFolder & "\Previous Versions\" & fso.getfilename(workbookPath)
    Application.DisplayAlerts = True
    Macroworkbook.Close
    
    Else:
    Macroworkbook.Save
    Adminsheet.Activate
      
End If

End If

If ThisVersion = NewVersion Then

    If Dir(Adminsheet.Range("DatabaseFile")) = "" Then
    Dim Prompt As Integer
    
    Prompt = MsgBox("Database Folder or Database Filename has changed. Please select Database File.", vbOKCancel)
    If Prompt = vbOK Then
     Call BrowseDataFile
     End If
    End If

End If

Set fso = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic


End Sub
