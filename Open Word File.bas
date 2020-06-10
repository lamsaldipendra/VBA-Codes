Attribute VB_Name = "GuideModule"
Option Explicit

Sub OpenUserguide()

Dim WordApp As Object, WordDoc As Object

Set WordApp = CreateObject("Word.Application")
Set WordDoc = WordApp.Documents.Open("T:\FCGF\FIN\FCGP only\General\ACCOUNTS\2 SUPPLIER ACCOUNTS\VCC\VCC PROJECT\2. 2003VCC WORKING FILE\Documents\VCC Macro Workbook Guide.docx")

WordApp.Visible = True
WordApp.Activate

Set WordDoc = Nothing
Set WordApp = Nothing

End Sub
