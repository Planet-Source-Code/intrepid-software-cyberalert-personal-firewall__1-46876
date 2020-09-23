Attribute VB_Name = "mdlFirewall"
'Declare everything
Option Explicit
Dim tcpt As MIB_TCPTABLE
'Used to hide the dos window (so it doesnt interupt
'anything)
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Finds a window
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Last declares
Public lngLastSent As Long
Public lngLastRecieved As Long
'INI file access
Public cINIFile As New cINI
Public cINFO As New cINI
'Declares for Connection ID
Private Type Connection_
Filename As String
ProcessId As Long
TCPEntryNum As Long
LocalPort As String
RemotePort As String
LocalHost As String
RemoteHost As String
State As String
TCP As Boolean
End Type
Public Connection(2000) As Connection_
Public OldConnection(2000) As Connection_
Public StatsLen As Long
' ----------------------------
' Support Routines
' ----------------------------


Const CHUNK_SIZE = 10240
Dim OldCnt As Long
Public Ontop As New clsOnTop
Public AlertWindows As String
Public AlreadyAsked As String
Public ServiceAlert As String
'''' counters
Public lngIncoming As Long
Public lngOutgoing As Long
Public lngBlocked As Long

Public oldLngIncoming As Long
Public oldLngOutgoing As Long
Public oldLngBlocked As Long
Public OldLen As Long
Public Function Already(idNum As Long) As Boolean

cINIFile.Section = "ALREADY"
cINIFile.Key = Connection(idNum).Filename & Connection(idNum).LocalPort & Connection(idNum).RemotePort & Connection(idNum).RemoteHost
cINIFile.Default = "-"
If cINIFile.Value <> "YES" Then
'Yes this host was never connected to by it
cINIFile.Value = "YES"
Already = False
Else
Already = True
End If

End Function


Public Sub CheckForHackers(LocP As String, idNum As Long)
If Trusted(Connection(idNum).RemoteHost) Then Exit Sub
If LocP = "" Then Exit Sub

If InStr(1, ServiceAlert, ("-=" & LocP & Connection(idNum).RemoteHost & "=-"), vbTextCompare) Then
Exit Sub
End If

ServiceAlert = ServiceAlert & "-=" & LocP & Connection(idNum).RemoteHost & "=-"

cINFO.Section = "HACKERS"
cINFO.Key = LocP
cINFO.Default = "-"
If cINFO.Value <> "-" Then
Dim frmS As New frmServiceAlert
LogSecurity "Moderate", "A potentially dangerous port " & LocP & " allowed a connection from " & Connection(idNum - 1).RemoteHost & ". This port is " & cINFO.Value
frmS.lblDesc.Caption = "The following port " & LocP & " has allowed a connection from " & Connection(idNum - 1).RemoteHost & ". This port is " & cINFO.Value & ". It could allow hackers to control your computer or utilize malicious activity. On ""High security"" CyberAlert automatically closes the connection, and moderate does not close it by default. You can either choose to allow this port to be accessed by the host or not."
CloseConnection idNum, Connection(idNum).Filename
frmS.Show
frmS.lblDesc.Tag = LocP
frmS.Tag = Connection(idNum).RemoteHost
Ontop.MakeTopMost frmS.hwnd
End If

End Sub

Public Function Trusted(RemoteHost As String)
cINFO.Section = "NETWORKING"
cINFO.Default = "-"
cINFO.Key = RemoteHost
If cINFO.Value = "TRUSTED" Then Trusted = True
End Function

Public Sub CloseConnection(iNum As Long, AppName As String)
'On Error Resume Next
Dim l As Long
Dim lvl As Long
'MsgBox Connection(iNum).RemoteHost
        If AppName = Connection(iNum).Filename Then
        
        LogTraffic "Blocked", Connection(iNum).Filename & " traffic was blocked. Remote Computer was " & Connection(iNum).RemoteHost & " on port " & Connection(iNum).RemotePort & "."
        
        l = Len(MIB_TCPTABLE)
        GetTcpTable tcpt, l, 0
        tcpt.table(iNum).dwState = 12
        SetTcpEntry tcpt.table(iNum)
        End If
        

End Sub
Public Sub Alert(Filename As String, idNum As Long, ProcessId As Long)
If Filename = "" Then Exit Sub

If Already(idNum) = True Then Exit Sub

'MsgBox Connection(IDNum).LocalPort & vbCrLf & Connection(IDNum).LocalHost & vbCrLf & Connection(IDNum).State
'If already has an alert window then exit subroutine
If InStr(1, AlertWindows, "-=" & Filename & "=-") Then Exit Sub

'Nope so add a alert window
AlertWindows = AlertWindows & "-=" & Filename & "=-"

LogSecurity "Low", "Unknown Application """ & Filename & """ attempted access."

SuspendThreads ProcessId

Dim frmA As New frmAlert
Dim tmpString As String
'MsgBox Connection(IDNum).LocalHost & vbCrLf & Connection(IDNum).RemoteHost
tmpString = mdlFile.GetFileDescription(Filename)
If Connection(idNum).LocalPort = Connection(idNum).RemotePort Then
'INCOMING!
frmA.lblDesc.Caption = tmpString & " (" & Mid(Filename, InStrRev(Filename, "\") + 1) & ") is allowing a connection from " & Connection(idNum).RemoteHost & " on port " & Connection(idNum).RemotePort & " . Do you want to allow this program to access the network?"
Else
frmA.lblDesc.Caption = tmpString & " (" & Mid(Filename, InStrRev(Filename, "\") + 1) & ") is trying to connect to " & Connection(idNum).RemoteHost & " using port " & Connection(idNum).RemotePort & " . Do you want to allow this program to access the network?"
End If

frmA.Tag = ProcessId & "," & idNum & "," & Filename
frmA.Show
Ontop.MakeTopMost frmA.hwnd


End Sub


Public Sub Crc32Check(Filename As String)


    If Filename <> "" Then
    cINIFile.Section = "CRC32"
    cINIFile.Key = Filename
    cINIFile.Default = "-"
    If cINIFile.Value = "-" Then
    cINIFile.Value = FileDateTime(Filename)
    Else
        If cINIFile.Value <> FileDateTime(Filename) Then
        'CORRUPTION
        
        End If
    End If
    End If


'originally i had it use Crc32 Checks
'but it takes too much process time
End Sub




Public Sub Execute(Optional force As Boolean = False)
'////////// STEP 1 - New connections? \\\\\\\\\\\'

Dim strRetVal As String
Dim i As Long
Dim Item As ListItem


If Refresh = False And force = False Then Exit Sub


'////////// Step 2.0 - Clear Connections ID \\\\\'
For i = 0 To StatsLen
OldConnection(i).Filename = Connection(i).Filename
Next i

OldLen = StatsLen
Erase Connection


'////////// Step 2.1 - Use Netstat -o to Map \\\\\'

strRetVal = mdlCmd.Execute("Netstat - o")
'Parse the strRetVal
Parse strRetVal

'////////// Step 3.1 - List Processes \\\\\\\\\\\'
mdlProcess2.LoadNTProcess



oldLngIncoming = lngIncoming
oldLngOutgoing = lngOutgoing
oldLngBlocked = lngBlocked
lngIncoming = 0
lngOutgoing = 0
lngBlocked = 0

    
    
    For i = 0 To StatsLen
    
    '////////// Step 3.2 - View Rules \\\\\\\\\\\'
    'Now get the rule for it
    cINIFile.Section = "RULES"
    cINIFile.Key = Connection(i).Filename
    cINIFile.Default = 0
    
    If Connection(i).LocalPort = Connection(i).RemotePort And Connection(i).LocalPort <> "" Then
    lngIncoming = lngIncoming + 1
    LogTraffic "Incoming", Connection(i).Filename & " allowed traffic from " & Connection(i).RemoteHost & " on port " & Connection(i).LocalPort & "."
    Else
    LogTraffic "Outbound", Connection(i).Filename & " sent traffic to " & Connection(i).RemoteHost & " on port " & Connection(i).RemotePort & "."
    lngOutgoing = lngOutgoing + 1
    End If
    

    If frmMain.mnuSecurityAllow.Checked = True Then GoTo skip
    If frmMain.mnuSecurityBlock.Checked = True Then GoTo block
    Select Case cINIFile.Value
    Case "0", "" 'Ask
    If frmBlockAll.Visible = True Then
    CloseConnection i, Connection(i).Filename
    Else
    
    Alert Connection(i).Filename, i, Connection(i).ProcessId
    End If
    
    Case "1" 'Block
    
block:
    lngBlocked = lngBlocked + 1
    CloseConnection i, Connection(i).Filename
    Case "2" 'Trust

    Case "3" 'Terminate
    KillProcessById Connection(i).ProcessId
    End Select
    DoEvents
    Next i
    
skip:


'/// STEP 3.3 - SET TRAY ICON \\\\\\\\\\\\\\\'

If oldLngIncoming = lngIncoming And oldLngOutgoing = lngOutgoing And oldLngBlocked = lngBlocked Then
'NO TRAFFIC
Set frmMain.SysTray.TrayIcon = frmMain.ilTray.ListImages(1).ExtractIcon
End If

If oldLngIncoming = lngIncoming And oldLngOutgoing = lngOutgoing And oldLngBlocked <> lngBlocked Then
'BLOCK ALL
Set frmMain.SysTray.TrayIcon = frmMain.ilTray.ListImages(2).ExtractIcon
End If

If oldLngIncoming <> lngIncoming And oldLngOutgoing <> lngOutgoing And oldLngBlocked = lngBlocked Then
'ALLOW ALL
Set frmMain.SysTray.TrayIcon = frmMain.ilTray.ListImages(5).ExtractIcon
End If

If oldLngIncoming = lngIncoming And oldLngOutgoing <> lngOutgoing And oldLngBlocked <> lngBlocked Then
'BLOCK OUT
Set frmMain.SysTray.TrayIcon = frmMain.ilTray.ListImages(3).ExtractIcon
End If

If oldLngIncoming <> lngIncoming And oldLngOutgoing = lngOutgoing And oldLngBlocked <> lngBlocked Then
'BLOCK IN
Set frmMain.SysTray.TrayIcon = frmMain.ilTray.ListImages(4).ExtractIcon
End If

If oldLngIncoming = lngIncoming And oldLngOutgoing <> lngOutgoing And oldLngBlocked = lngBlocked Then
'ALLOW OUT
Set frmMain.SysTray.TrayIcon = frmMain.ilTray.ListImages(6).ExtractIcon
End If

If oldLngIncoming <> lngIncoming And oldLngOutgoing = lngOutgoing And oldLngBlocked = lngBlocked Then
'ALLOW IN
Set frmMain.SysTray.TrayIcon = frmMain.ilTray.ListImages(7).ExtractIcon
End If



'///////// Step 4 - Fill up Listview \\\\\\\\\\'
'(Only if visible)
cINIFile.Path = (App.Path & "\Firewall.dat")



If frmMain.Visible = True Then
frmMain.lvFirewall.ListItems.Clear
frmMain.lblIncoming = lngIncoming
frmMain.lblOutgoing = lngOutgoing
frmMain.lblBlocked = lngBlocked
For i = 0 To StatsLen

        Crc32Check (Connection(i).Filename)

    If Connection(i).Filename <> "" Then
    Set Item = frmMain.lvFirewall.ListItems.Add()
    Item.Text = mdlFile.GetFileDescription(Connection(i).Filename)
    If Item.Text = "" Then Item.Text = "(No Application Name)"
    Item.SubItems(1) = mdlFile.GetFileVersion(Connection(i).Filename)
    Item.SubItems(2) = Connection(i).Filename
    Item.Tag = i
    'Now get the rule for it
    cINIFile.Section = "RULES"
    cINIFile.Key = Connection(i).Filename
    cINIFile.Default = 0
    Select Case cINIFile.Value
    Case "0", "" 'Ask
    Item.SmallIcon = 1
    Case "1" 'Block
    Item.SmallIcon = 2
    CloseConnection i, Connection(i).Filename
    Case "2" 'Trust
    Item.SmallIcon = 3
    Case "3" 'Kill
    KillProcessById Connection(i).ProcessId
    frmMain.lvFirewall.ListItems.Remove Item.Index
    End Select
    
    End If
    
Next i
End If
'set timer
frmMain.Timer2.Enabled = True
End Sub
Public Function Refresh(Optional force As Boolean = False) As Boolean

On Error Resume Next


Dim NewCnt As Long


GetTcpTable tcpt, Len(MIB_TCPTABLE), 0
NewCnt = tcpt.dwNumEntries

If NewCnt <> OldCnt Or force Then
Refresh = True
    OldCnt = NewCnt
End If
End Function
Public Sub Parse(Data As String)
Dim SplitData() As String 'Split by vbCrLf (Line Returns)
Dim LineSplit() As String
Dim i As Long
Dim LocP As String
Dim RemP As String
Dim LocA As String
Dim RemA As String
Dim State As String
Dim Y As Long
Dim PID As Long

On Error Resume Next

'While there are more than 1 space chrs in a row
'remove them
Do While InStr(1, Data, "  ")
Data = Replace(Data, "  ", " ")
DoEvents
Loop

'Split by vbCrLf (Line Returns)
SplitData = Split(Data, vbCrLf)

'Split by Spaces
For Y = 0 To UBound(SplitData)
LineSplit = Split(SplitData(Y), " ")
DoEvents
    'Now find everything
    If LineSplit(0) <> "PROTO" Then
        If LineSplit(0) = "TCP " Then
        Connection(i).TCP = True
        Else
        Connection(i).TCP = False
        End If
    
    LocA = Mid(LineSplit(2), 1, InStr(1, LineSplit(2), ":"))
    LocP = Mid(LineSplit(2), InStr(1, LineSplit(2), ":") + 1, Len(LineSplit(2)) - InStr(1, LineSplit(2), ":"))
    RemP = Mid(LineSplit(3), InStr(1, LineSplit(3), ":") + 1, Len(LineSplit(3)) - InStr(1, LineSplit(3), ":"))
    RemA = Mid(LineSplit(3), 1, InStr(1, LineSplit(3), ":"))
    State = LineSplit(4)
    PID = 0
    PID = LineSplit(5)
    CheckForHackers LocP, i
    If PID <> 0 Then
    Connection(i).LocalHost = Replace(LocA, ":", "")
    Connection(i).LocalPort = LocP
    Connection(i).RemoteHost = Replace(RemA, ":", "")
    Connection(i).RemotePort = RemP
    Connection(i).State = State
    Connection(i).ProcessId = PID
    i = i + 1
    End If
    
    
    End If

Next Y
StatsLen = i
End Sub


