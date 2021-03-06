VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EmailForm 
   Caption         =   "Fill the form"
   ClientHeight    =   9045.001
   ClientLeft      =   40
   ClientTop       =   400
   ClientWidth     =   9300.001
   OleObjectBlob   =   "Email Form.frx":0000
End
Attribute VB_Name = "EmailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EmailForm_Initialize()

ConfirmationInput.Value = ""

ReservationInput.Value = ""
Frame1.Visible = False

ConfirmationInput.SetFocus

End Sub

Private Sub Cancelbtn_Click()
    Unload EmailForm
    EmailForm.Hide
    End
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub CancelDateBox_Change()

End Sub

Private Sub CardText_Change()

End Sub

Private Sub ChargedText_Change()
On Error Resume Next
DisputedText.Value = Format(ChargedText.Value - ExpectedText.Value, "#,##0.00")
End Sub

Sub DetailCombobox_Change()

On Error Resume Next
    If IssueTypeBox.Value = "Overcharge" Then
        With DetailMsgBox
            .Value = ""
            .Value = WorksheetFunction.Index(ThisWorkbook.Worksheets("Validation").Range("OCDetails"), WorksheetFunction.Match(Me.DetailCombobox.Value, ThisWorkbook.Worksheets("Validation").Range("OCType"), 0))
        End With
    End If

    If IssueTypeBox.Value = "Cancellation" Then
        With DetailMsgBox
            .Value = ""
            .Value = WorksheetFunction.Index(ThisWorkbook.Worksheets("Validation").Range("CnxDetails"), WorksheetFunction.Match(Me.DetailCombobox.Value, ThisWorkbook.Worksheets("Validation").Range("CnxType"), 0))
        End With
    End If

End Sub

Private Sub DetailMsgBox_Change()

End Sub

Private Sub ExpectedText_Change()
On Error Resume Next
DisputedText.Value = Format(ChargedText.Value - ExpectedText.Value, "#,##0.00")
End Sub

Private Sub IssueTypeBox_Change()

If IssueTypeBox.Value = "Overcharge" Then
OverchargeTypelbl.Visible = True
CancelTypelbl.Visible = False
CancelDatelbl.Visible = False
CancelDateBox.Visible = False

With DetailCombobox
.Clear
.List = ThisWorkbook.Worksheets("Validation").Range("OCType").Value ' [Other!C2:C11].Value
.Value = .List(0)
End With

'With DetailMsgBox
    '.Value = WorksheetFunction.Index(ThisWorkbook.Worksheets("Validation").Range("OCDetails"), WorksheetFunction.Match(Me.DetailCombobox.Value, ThisWorkbook.Worksheets("links").Range("OCType"), 0))
'End With

End If

If IssueTypeBox.Value = "Cancellation" Then
OverchargeTypelbl.Visible = False
CancelTypelbl.Visible = True
CancelDatelbl.Visible = True
CancelDateBox.Visible = True

DetailCombobox.Clear
With DetailCombobox
.Clear
.List = ThisWorkbook.Worksheets("Validation").Range("CnxType").Value ' [Other!C2:C11].Value
.Value = .List(0)
End With

'With DetailMsgBox
    '.Value = WorksheetFunction.Index(ThisWorkbook.Worksheets("Validation").Range("CnxDetails"), WorksheetFunction.Match(Me.DetailCombobox.Value, ThisWorkbook.Worksheets("links").Range("CnxType"), 0))
'End With

End If

End Sub

Private Sub OkayBtn_Click()
 Me.Hide
End Sub

Private Sub UserForm_Initialize()

On Error Resume Next
Dim wb As Workbook
Dim ws As Worksheet
Dim Curr As String

Set wb = ThisWorkbook
Set ws = ThisWorkbook.Worksheets("2003VCC")

With Me
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End With

ConfirmationInput.Value = ""

ReservationInput.Value = ""

With IssueTypeBox
    .AddItem "Overcharge"
    .AddItem "Cancellation"
    .Value = .List(0)
End With

CardText.Value = ws.Range("O9").Value

Curr = ws.Range("I11").Value
If Curr = "USN" Then
CurrencyText.Value = "USD"
ElseIf Curr = "GBX" Then
CurrencyText.Value = "GBP"
Else
CurrencyText.Value = Curr
End If

ExpectedText.Value = Format(ws.Range("O11").Value, "#,##0.00")
ChargedText.Value = Format(ws.Range("O12").Value, "#,##0.00")
DisputedText.Value = Format(ChargedText.Value - ExpectedText.Value, "#,##0.00")

ConfirmationInput.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
    Unload EmailForm
    EmailForm.Hide
    End
    End If
End Sub
