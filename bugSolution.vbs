Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Or param1 = "" Then 'Handle both empty and zero-length strings
    Err.Raise vbError + 1, , "Parameter cannot be empty or zero length"
  End If
  On Error GoTo 0
End Function