Attribute VB_Name = "EmailTemplate"
Option Explicit

Sub GenerateEmail()

Dim Confirmation As String, Reservation As String, IssueType As String, Details As String, CnxDate As String, Curr As String
Dim ChargedAmount As String, CorrectAmount As String, OverchargedAmount As String

EmailForm.Show

With EmailForm

If .ConfirmationInput.Value <> "" Then
Confirmation = "Confirmation: " & .ConfirmationInput.Value
Else
Confirmation = ""
End If

If .ReservationInput.Value = "" Then
Reservation = ""
Else
Reservation = "Reservation/HBSI ID: " & .ReservationInput.Value
End If

If Confirmation <> "" And Reservation <> "" Then
Reservation = "  / Reservation: " & .ReservationInput.Value
ElseIf Confirmation = "" And Reservation = "" Then
Reservation = "Reservation/HBSI ID: Not Applicable"
End If

IssueType = .IssueTypeBox.Value
Details = .DetailMsgBox.Value
CnxDate = Format(.CancelDateBox, "dd-Mmm-yyyy")

Curr = .CurrencyText.Value
ChargedAmount = Curr & Space(1) & .ChargedText.Value 'charged amt
CorrectAmount = Curr & Space(1) & .ExpectedText.Value 'correct
OverchargedAmount = Curr & Space(1) & .DisputedText.Value ' overcharge

End With
Unload EmailForm
EmailForm.Hide

Application.ScreenUpdating = False

Dim aBook As Workbook, aSht As Worksheet, lSht As Worksheet
Dim Subject As String, Card As String, ChargedDate As String
Dim Booking As String, GuestName As String, TravelDate As String, Conf As String, Res As String, RefundbyDate As String
Dim Reason As String

Set aBook = ThisWorkbook
Set aSht = aBook.Worksheets("2003VCC")
Set lSht = aBook.Worksheets("Validation")

With aSht
Card = aSht.Range("O9").Value ' card

ChargedDate = Format(aSht.Range("L10").Value, "dd-mmm-yyyy") ' charged date

Booking = aSht.Range("I6").Value ' Booking
GuestName = aSht.Range("O5").Value ' GuestName
TravelDate = Format(aSht.Range("O6").Value, "dd-mmm-yyyy") 'TravelDate

RefundbyDate = Format(Date + 5, "dd-mmm-yyyy") ' RefbyDate

If aSht.Range("Y4") <> "" Then
Reason = aSht.Range("Y4").Value & ". "
Else
Reason = aSht.Range("Y4").Value
End If

If IssueType = "Overcharge" Then
Subject = aSht.Range("I4").Value & " - " & aSht.Range("L4").Value & " INV DATE " & ChargedDate & Space(1) & "-" & Space(1) & "VCC Overcharged " & OverchargedAmount & " Refund by COB " & RefundbyDate
ElseIf IssueType = "Cancellation" Then
Subject = aSht.Range("I4").Value & " - " & aSht.Range("L4").Value & " INV DATE " & ChargedDate & Space(1) & "-" & Space(1) & "VCC charged CXLD booking " & OverchargedAmount & " Refund by COB " & RefundbyDate

End If
End With


'aBook.Worksheets("links").Range("I3").Value = ChargedDate
Dim WordApp As Object

Dim wdDoc As Object

'Open word application

Set WordApp = CreateObject("Word.Application")

'Add a new document

If IssueType = "Cancellation" Then
Set wdDoc = WordApp.Documents.Open("T:\FCGF\FIN\FCGP only\General\ACCOUNTS\2 SUPPLIER ACCOUNTS\VCC\VCC PROJECT\2. 2003VCC WORKING FILE\Documents\CancellationEmailTemplate.docm")
Else
Set wdDoc = WordApp.Documents.Open("T:\FCGF\FIN\FCGP only\General\ACCOUNTS\2 SUPPLIER ACCOUNTS\VCC\VCC PROJECT\2. 2003VCC WORKING FILE\Documents\OverchargeEmailTemplate.docm")
End If
'Make Word Visible, not active

WordApp.Visible = True
WordApp.Activate

'Destroy object variables

With wdDoc.Content.Find

  .Text = "#Subject"
  .Replacement.Text = Subject
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll

  .Text = "#Card"
  .Replacement.Text = Card
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll

  .Text = "#OverchargedAmount"
  .Replacement.Text = OverchargedAmount
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll

  .Text = "#RefundbyDate"
  .Replacement.Text = RefundbyDate
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#ChargedDate"
  .Replacement.Text = ChargedDate
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#ChargedAmount"
  .Replacement.Text = ChargedAmount
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#CorrectAmount"
  .Replacement.Text = CorrectAmount
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#Booking"
  .Replacement.Text = Booking
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#GuestName"
  .Replacement.Text = GuestName
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#TravelDate"
  .Replacement.Text = TravelDate
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#CnxDate"
  .Replacement.Text = CnxDate
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#Confirmation"
  .Replacement.Text = Confirmation
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#Reservation"
  .Replacement.Text = Reservation
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll

  .Text = "#Reason"
  .Replacement.Text = Reason
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll
  
  .Text = "#Details"
  .Replacement.Text = Details
  .Wrap = wdFindContinue
  .Execute Replace:=wdReplaceAll

End With
Application.ScreenUpdating = True

Set wdDoc = Nothing

Set WordApp = Nothing

End Sub


