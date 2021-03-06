VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WEXCBForm 
   Caption         =   "WEX Chargeback Form"
   ClientHeight    =   7320
   ClientLeft      =   40
   ClientTop       =   400
   ClientWidth     =   13720
   OleObjectBlob   =   "CBForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WEXCBForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancelbtn_Click()
    Unload WEXCBForm
    WEXCBForm.Hide
    End
End Sub

Private Sub ExpectedAmtText_Change()
On Error Resume Next
With WEXCBForm
    .DispuetdText = Format(.TransactionAmtText - .ExpectedAmtText, "#,##0.00")
End With
End Sub


Private Sub Label1_Click()

End Sub

Private Sub OKbtn_Click()
WEXCBForm.Hide
End Sub

Private Sub TransactionAmtText_Change()
On Error Resume Next
With WEXCBForm
    .DispuetdText = Format(.TransactionAmtText - .ExpectedAmtText, "#,##0.00")
End With
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
Dim Wbc As Workbook, Wsc As Worksheet, Curr As String, PurchaseId As String
    
Set Wbc = ThisWorkbook
Set Wsc = Wbc.Worksheets("2003VCC")
        
    With WEXCBForm
        .RaisedbyText = Application.Username
        .BookingText = Wsc.Range("I6")
        .TravelDateText = Format(Wsc.Range("O6"), "dd-Mmm-yy")
        .CreditorText = Wsc.Range("I5")
        .CardText = Wsc.Range("O9")
        .TransactionAmtText = Format(Wsc.Range("O12").Value, "#,##0.00")
        .ExpectedAmtText = Format(Wsc.Range("O11").Value, "#,##0.00")
        .DispuetdText = Format(.TransactionAmtText - .ExpectedAmtText, "#,##0.00")
            
        Curr = Wsc.Range("I11").Value
            If Curr = "USN" Then
                CurrencyText.Value = "USD"
                    ElseIf Curr = "GBX" Then
                        CurrencyText.Value = "GBP"
                    Else
                CurrencyText.Value = Curr
            End If
                With .ChargebackReasonComboBox
                    .Clear
                    .List = ThisWorkbook.Worksheets("Validation").Range("ChargebackReason").Value ' [Other!C2:C11].Value
                    .Value = .List(1)
                End With
        .CnxDateText = ""
        .ChargebackDtlText = Wsc.Range("Y4")
        .MerchantText = Wsc.Range("O4")
        PurchaseId = Wsc.Range("O8")
        .PurchaseIDText = PurchaseId
        .ReferenceText = WorksheetFunction.Index(Wsc.Range("V20:V1000"), WorksheetFunction.Match(PurchaseId, Wsc.Range("J20:J1000"), 0))
        .TransactionDtText = Format(WorksheetFunction.Index(Wsc.Range("N20:N1000"), WorksheetFunction.Match(PurchaseId, Wsc.Range("J20:J1000"), 0)), "dd-Mmm-yy")
        .PostingDtText = Format(WorksheetFunction.Index(Wsc.Range("M20:M1000"), WorksheetFunction.Match(PurchaseId, Wsc.Range("J20:J1000"), 0)), "dd-Mmm-yy")
        .ContactDtlText = "Never responded when contacted via email; "
        .ContactDateText = ""
    End With
        
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
    Unload WEXCBForm
    WEXCBForm.Hide
    End
    End If
End Sub
