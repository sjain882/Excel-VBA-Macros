Sub RunProgram(sFile, Optional args, Optional runInFolder)
'Version: 1.000
'Purpose: This run passed sFile
  Dim RetVal As Long
  On Error Resume Next
  RetVal = ShellExecute(0, "open", sFile, "", "", SW_SHOWMAXIMIZED)
End Sub