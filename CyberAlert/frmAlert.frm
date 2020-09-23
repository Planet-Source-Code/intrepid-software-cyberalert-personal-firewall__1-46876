VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CyberAlert Firewall"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remeber my answer, and do not ask me again for this application."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label lblDESC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAlert.frx":000C
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAlert.frx":00D0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Check1.Value = vbChecked Then
'WRITE Rule allowing program
cINIFile.Section = "RULES"
Dim TagSplit() As String
Dim PID As Long
TagSplit = Split(Tag, ",")
cINIFile.Key = TagSplit(2)
cINIFile.Value = "2"
'Now set the icon
For i = 1 To frmMain.lvFirewall.ListItems.Count
If frmMain.lvFirewall.ListItems(i).SubItems(2) = TagSplit(2) Then
frmMain.lvFirewall.ListItems(i).SmallIcon = 3
Exit For
End If
Next i

End If

Unload Me

End Sub

Private Sub Command2_Click()
Dim TagSplit() As String
Dim PID As Long
TagSplit = Split(Tag, ",")

If Check1.Value = vbChecked Then
'WRITE Rule BLOCKING program
cINIFile.Section = "RULES"
cINIFile.Key = TagSplit(2)
cINIFile.Value = "1"

For i = 1 To frmMain.lvFirewall.ListItems.Count
If frmMain.lvFirewall.ListItems(i).SubItems(2) = TagSplit(2) Then
frmMain.lvFirewall.ListItems(i).SmallIcon = 2
Exit For
End If
Next i

End If
PID = TagSplit(1)
CloseConnection PID, Connection(PID).Filename
Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim TagSplit() As String
Dim PID As Long
TagSplit = Split(Tag, ",")
PID = TagSplit(0)
mdlProcess1.ResumeThreads PID
AlertWindows = Replace(AlertWindows, "-=" & TagSplit(2) & "=-", "")
End Sub

