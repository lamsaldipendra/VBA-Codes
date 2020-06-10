'' Check if the file is opened and locked by another user
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
