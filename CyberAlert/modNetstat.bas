Attribute VB_Name = "modNetstat"
Option Explicit

'-------------------------------------------------------------------------------
'Types and function for the ICMP table:
Public lc
Public MIBICMPSTATS As MIBICMPSTATS
Public Type MIBICMPSTATS
    dwEchos As Long
    dwEchoReps As Long
End Type

Public MIBICMPINFO As MIBICMPINFO
Public Type MIBICMPINFO
    icmpOutStats As MIBICMPSTATS
End Type

Public MIB_ICMP As MIB_ICMP
Public Type MIB_ICMP
    stats As MIBICMPINFO
End Type

Public Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (pStats As MIBICMPINFO) As Long
Public Last_ICMP_Cnt As Integer 'ICMP count

'-------------------------------------------------------------------------------
'Types and functions for the TCP table:

Type MIB_TCPROW
  dwState As Long
  dwLocalAddr As Long
  dwLocalPort As Long
  dwRemoteAddr As Long
  dwRemotePort As Long
End Type

Type MIB_TCPTABLE
  dwNumEntries As Long
  table(2000) As MIB_TCPROW
End Type
Public MIB_TCPTABLE As MIB_TCPTABLE

Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function SetTcpEntry Lib "IPhlpAPI" (pTcpRow As MIB_TCPROW) As Long 'This is used to close an open port.
Public IP_States(13) As String
Private Last_Tcp_Cnt As Integer 'TCP connection count

'-------------------------------------------------------------------------------
'Types and functions for winsock:

Private Const AF_INET = 2
Private Const IP_SUCCESS As Long = 0
Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const SOCKET_ERROR As Long = -1
Private Const WS_VERSION_REQD As Long = &H101

Type HOSTENT
    h_name As Long        ' official name of host
    h_aliases As Long     ' alias list
    h_addrtype As Integer ' host address type
    h_length As Integer   ' length of address
    h_addr_list As Long   ' list of addresses
End Type

Type servent
  s_name As Long            ' (pointer to string) official service name
  s_aliases As Long         ' (pointer to string) alias list (might be null-seperated with 2null terminated)
  s_port As Long            ' port #
  s_proto As Long           ' (pointer to) protocol to use
End Type

Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Long
   wMaxUDPDG As Long
   dwVendorInfo As Long
End Type

Public Declare Function ntohs Lib "WSOCK32.DLL" (ByVal netshort As Long) As Long
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal CP As String) As Long
Private Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal inn As Long) As Long
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" (Addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal host_name As String) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Declare Function lstrlen Lib "kernel32" (ByVal lpString As Any) As Integer
Private Blocked As Boolean

Private Type Ports_
RemA As String
RemP As String
State As String
LocP As String
End Type
Public Ports(0 To 2000) As Ports_
Sub InitStates()
  IP_States(0) = "UNKNOWN"
  IP_States(1) = "CLOSED"
  IP_States(2) = "LISTENING"
  IP_States(3) = "SYN_SENT"
  IP_States(4) = "SYN_RCVD"
  IP_States(5) = "ESTABLISHED"
  IP_States(6) = "FIN_WAIT1"
  IP_States(7) = "FIN_WAIT2"
  IP_States(8) = "CLOSE_WAIT"
  IP_States(9) = "CLOSING"
  IP_States(10) = "LAST_ACK"
  IP_States(11) = "TIME_WAIT"
  IP_States(12) = "DELETE_TCB"
End Sub

Public Function GetAscIP(ByVal inn As Long) As String
  Dim nStr&
    Dim lpStr As Long
    Dim retString As String
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
    If lpStr Then
        nStr = lstrlen(lpStr)
        If nStr > 32 Then nStr = 32
        CopyMemory ByVal retString, ByVal lpStr, nStr
        retString = Left(retString, nStr)
        GetAscIP = retString
    Else
        GetAscIP = "Unable to get IP"
    End If
End Function
