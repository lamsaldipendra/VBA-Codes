VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Not Intersect(Target, Range("G20:AE3000")) Is Nothing And Range("C11") = False And Target.Count = 1 Then
Application.ScreenUpdating = False
Dim tgRow As Long

tgRow = Target.Row
Range("E1").Value = tgRow
Call LoadHeader

Application.ScreenUpdating = True
End If
End Sub
