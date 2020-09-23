Attribute VB_Name = "mdlLog"
Public Function AscDate() As String
AscDate = Replace(Date, "/", "_")
End Function

Public Sub LogSecurity(Severity As String, Description As String)

On Error Resume Next
MkDir App.Path & "\Logs\Security"
Dim ff As Integer
ff = FreeFile
Open App.Path & "\Logs\Security\" & AscDate & ".txt" For Append As #ff
Print #ff, "[" & Time & "] " & Severity & " - " & Description
Close #ff

End Sub


Public Sub LogTraffic(InOutBlock As String, Description As String)

On Error Resume Next
MkDir App.Path & "\Logs\Traffic\"
Dim ff As Integer
ff = FreeFile
Open App.Path & "\Logs\Traffic\" & AscDate & ".txt" For Append As #ff
Print #ff, "[" & Time & "] " & InOutBlock & " - " & Description
Close #ff

End Sub
