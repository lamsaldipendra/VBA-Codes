Attribute VB_Name = "Wexchargeback"
Option Explicit

Sub FillWexchargeback()

WEXCBForm.Show

Dim arrayData(1 To 22) As Variant

With WEXCBForm
    arrayData(1) = .BookingText
    arrayData(2) = Format(Date, "Medium Date")
    arrayData(3) = "FLIGHT CENTRE CORP(Flight Centre Global Product)" 'Client Name
    arrayData(4) = .RaisedbyText
    arrayData(5) = "ap-support@flightcentre.com"
    arrayData(6) = .MerchantText
    arrayData(7) = .CardText
    arrayData(8) = .CurrencyText
    arrayData(9) = "'" & .ReferenceText
    arrayData(10) = .TransactionDtText
    arrayData(11) = .PostingDtText
    arrayData(12) = .TransactionAmtText
    arrayData(13) = .ExpectedAmtText
    arrayData(14) = .DispuetdText
    arrayData(15) = .ChargebackReasonComboBox
    arrayData(16) = Format(.CnxDateText, "dd-Mmm-yy")
    arrayData(17) = .ChargebackDtlText
    arrayData(18) = "'" & .PurchaseIDText
    arrayData(19) = "FLIGHT CENTRE CORP(Flight Centre Global Product)" ' Card Holder
    arrayData(20) = Format(.TravelDateText, "Medium Date")
    arrayData(21) = Format(.ContactDateText, "Medium Date")
    arrayData(22) = .ContactDtlText

    Dim rngTarget As Range
    Dim lrow As Integer
    
    lrow = ThisWorkbook.Worksheets("WEXCB").Cells(ThisWorkbook.Worksheets("WEXCB").Rows.Count, "A").End(xlUp).Row + 1
    
    Set rngTarget = ThisWorkbook.Worksheets("WEXCB").Range("A" & lrow & ":V" & lrow)
    rngTarget = arrayData
End With

Unload WEXCBForm
WEXCBForm.Hide

End Sub

Sub Exporttotemplate()
Application.ScreenUpdating = False

Dim ChargebackFile As FileDialog
Dim fso As Object
Dim Chargebackfilepath As String, ChargebackFileName As String

Set fso = CreateObject("Scripting.FileSystemObject")
Set ChargebackFile = Application.FileDialog(msoFileDialogFilePicker)

With ChargebackFile
    .Title = "Select Chargeback template File"
        If .Show <> -1 Then GoTo NoSel
            Chargebackfilepath = .SelectedItems(1)
            ChargebackFileName = fso.getfilename(Chargebackfilepath)
End With
Dim wb1 As Workbook, ws1 As Worksheet, wb2 As Workbook, ws2 As Worksheet, lrow1 As Integer, lrow2 As Integer, copyrange As Range
    Set wb1 = ThisWorkbook
    Set ws1 = wb1.Worksheets("WEXCB")
    
Application.Workbooks.Open FileName:=Chargebackfilepath, UpdateLinks:=0, ReadOnly:=False
    Set wb2 = Workbooks(ChargebackFileName)
    Set ws2 = wb2.Worksheets("Sheet1")
    
    lrow1 = ws1.Range("A" & ws1.Rows.Count).End(xlUp).Row + 1
    lrow2 = ws2.Range("A" & ws2.Rows.Count).End(xlUp).Row + 1
    
    Set copyrange = ws1.Range("A6:V" & lrow1)
    copyrange.Copy ws2.Range("A" & lrow2 & ":V" & lrow2 + lrow1 - 5)
    
        With wb2
            .Save
            .Close
        End With
NoSel:

Application.ScreenUpdating = True

End Sub

Sub ClearWexCB()

Dim answer As Integer
    answer = MsgBox("Are you sure you want to cancel?" & Chr(10) & "You will not be able to restore it after clearing", vbYesNo + vbCritical)
        If answer = vbNo Then
            GoTo No
        Else
        ThisWorkbook.Worksheets("WEXCB").Range("A6:V1000").ClearContents
        
        End If
No:
End Sub

