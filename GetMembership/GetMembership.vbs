On Error Resume Next
strDomain = "[YOURDOMAIN]"
strGroupName = WScript.Arguments(0)
If strGroupName = "" Then strGroupName = InputBox("Please enter a group name")

Err.Clear
Set Group = GetObject("WinNT://" & strDomain & "/" & strGroupName & ",group")
If Err.Number <> 0 Then
  WScript.Echo "Error getting " & strGroupName
  WScript.Quit 1
End If

For Each Member in Group.Members
  Members = Members & vbCrLf & Member.Name
Next

WScript.Echo "Members of " & strGroupName & ":" & Members
Set Group = Nothing


