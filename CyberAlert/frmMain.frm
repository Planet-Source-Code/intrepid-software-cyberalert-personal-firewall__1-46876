VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CyberAlert Personal Firewall"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   2880
      Top             =   3720
   End
   Begin MSComctlLib.ImageList ilTray 
      Left            =   4080
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5874
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CyberAlert.cSysTray SysTray 
      Left            =   3240
      Top             =   3120
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmMain.frx":614E
      TrayTip         =   "CyberAlert Firewall - Normal"
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2880
      Top             =   3240
   End
   Begin MSComctlLib.ImageList ilListview 
      Left            =   3120
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7302
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   6600
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Security Status:"
            TextSave        =   "Security Status:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6959
            Text            =   "Normal"
            TextSave        =   "Normal"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "    View Console"
            TextSave        =   "    View Console"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvFirewall 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilListview"
      SmallIcons      =   "ilListview"
      ColHdrIcons     =   "ilListview"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Application"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Path"
         Object.Width           =   14888
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   1217
      ButtonWidth     =   1693
      ButtonHeight    =   1164
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilToolbar"
      DisabledImageList=   "ilToolbar"
      HotImageList    =   "ilToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Block All"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Applications"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logs"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Security"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Traffic"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "Packet"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Text            =   "System"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Test"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "About"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ilToolbar 
         Left            =   5880
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D936
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":DD88
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":14022
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B524
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B976
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Applications:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1770
   End
   Begin VB.Label lblBlocked 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5040
      TabIndex        =   9
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Label lblOutgoing 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   5040
      TabIndex        =   8
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Label lblIncoming 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   240
      Left            =   5040
      TabIndex        =   7
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blocked Traffic:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   6
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outgoing Traffic:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3360
      TabIndex        =   5
      Top             =   1440
      Width           =   1560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incoming Traffic:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FF0000&
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Firewall"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   345
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cyber"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   1575
      Left            =   120
      Top             =   840
      Width           =   6615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit Firewall"
      End
   End
   Begin VB.Menu mnuSecurity 
      Caption         =   "Security"
      Begin VB.Menu mnuSecurityAllow 
         Caption         =   "Allow all"
      End
      Begin VB.Menu mnuSecurityNormal 
         Caption         =   "Normal"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSecurityBlock 
         Caption         =   "Block all"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsApp 
         Caption         =   "Applications"
      End
      Begin VB.Menu mnuToolsLogs 
         Caption         =   "Logs"
         Begin VB.Menu mnuToolsLogsSecurity 
            Caption         =   "Security"
         End
         Begin VB.Menu mnuToolsLogsTraffic 
            Caption         =   "Traffic"
         End
         Begin VB.Menu mnuToolsLogsPacket 
            Caption         =   "Packet"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuToolsLogsSystem 
            Caption         =   "System"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "Options"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsStart 
         Caption         =   "Automatically Start with Windows"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsUpdate 
         Caption         =   "Check for Updates"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFirewall 
      Caption         =   "mnuFirewall"
      Visible         =   0   'False
      Begin VB.Menu mnuFirewallTrust 
         Caption         =   "Trust"
      End
      Begin VB.Menu mnuFirewallAsk 
         Caption         =   "Ask"
      End
      Begin VB.Menu mnuFirewallBlock 
         Caption         =   "Block"
      End
      Begin VB.Menu mnuFirewallSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFirewallTerminate 
         Caption         =   "Always Terminate Process"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
cINIFile.Path = App.Path & "\Firewall.dat"
cINIFile.Section = "ALREADY"
cINIFile.DeleteSection
cINFO.Path = App.Path & "\Guard.dat"
On Error Resume Next
MkDir App.Path & "\Logs\"
MkDir App.Path & "\Logs\Traffic"
MkDir App.Path & "\Logs\Security"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
Cancel = True
End Sub


Private Sub lvFirewall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim strTemp As String
strTemp = lvFirewall.SelectedItem.Text
mnuFirewall.Tag = lvFirewall.SelectedItem.Index
If strTemp = "(No Application Name)" Then strTemp = Mid(lvFirewall.SelectedItem.SubItems(2), InStrRev(lvFirewall.SelectedItem.SubItems(2), "\") + 1)
If Button = 2 And strTemp <> "" Then
mnuFirewallTrust.Caption = "Trust " & strTemp
mnuFirewallAsk.Caption = "Ask " & strTemp
mnuFirewallBlock.Caption = "Block " & strTemp
mnuFirewallTerminate.Caption = "Always Terminate " & strTemp
PopupMenu mnuFirewall

End If

End Sub


Private Sub mnuFileClose_Click()
Hide
End Sub

Private Sub mnuFileExit_Click()
End
End Sub


Private Sub mnuFirewallAsk_Click()
Dim idNum As Long
'Add rule to ask this
idNum = mnuFirewall.Tag
cINIFile.Section = "RULES"
cINIFile.Key = lvFirewall.ListItems.Item(idNum).SubItems(2)
cINIFile.Value = "0"
lvFirewall.ListItems(idNum).SmallIcon = 1
mdlFirewall.Execute True
End Sub


Private Sub mnuFirewallBlock_Click()
Dim idNum As Long
'Add rule to ask this
idNum = mnuFirewall.Tag
cINIFile.Section = "RULES"
cINIFile.Key = lvFirewall.ListItems.Item(idNum).SubItems(2)
cINIFile.Value = "1"
lvFirewall.ListItems(idNum).SmallIcon = 2
mdlFirewall.Execute True
End Sub


Private Sub mnuFirewallTerminate_Click()
Dim idNum As Long
Dim PID As Long
'Add rule to ask this
idNum = mnuFirewall.Tag
cINIFile.Section = "RULES"
cINIFile.Key = lvFirewall.ListItems.Item(idNum).SubItems(2)
cINIFile.Value = "3"
PID = lvFirewall.ListItems(idNum).Tag
KillProcessById Connection(PID).ProcessId
lvFirewall.ListItems.Remove (idNum)
End Sub

Private Sub mnuFirewallTrust_Click()
Dim idNum As Long
'Add rule to ask this
idNum = mnuFirewall.Tag
cINIFile.Section = "RULES"
cINIFile.Key = lvFirewall.ListItems.Item(idNum).SubItems(2)
cINIFile.Value = "2"
lvFirewall.ListItems(idNum).SmallIcon = 3
End Sub


Private Sub mnuSecurityAllow_Click()
StatusBar1.Panels(2).Text = "Allow all"
mnuSecurityNormal.Checked = False
mnuSecurityAllow.Checked = True
mnuSecurityBlock.Checked = False
SysTray.TrayTip = "CyberAlert Firewall - Allowing All"
End Sub

Private Sub mnuSecurityBlock_Click()
StatusBar1.Panels(2).Text = "Block all"
mnuSecurityNormal.Checked = False
mnuSecurityAllow.Checked = False
mnuSecurityBlock.Checked = True
SysTray.TrayTip = "CyberAlert Firewall - Block"
End Sub

Private Sub mnuSecurityNormal_Click()
StatusBar1.Panels(2).Text = "Normal"
mnuSecurityNormal.Checked = True
mnuSecurityAllow.Checked = False
mnuSecurityBlock.Checked = False
SysTray.TrayTip = "CyberAlert Firewall - Normal"
End Sub

Private Sub mnuToolsApp_Click()
frmApplications.Show , Me
End Sub

Private Sub mnuToolsLogsSecurity_Click()
frmLogSecurity.Show , Me
End Sub

Private Sub mnuToolsLogsTraffic_Click()
frmLogTraffic.Show , Me
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

If Panel.Index = 3 Then

End If

End Sub

Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
mdlFirewall.Execute True
frmMain.Show
End Sub

Private Sub Timer1_Timer()
mdlFirewall.Execute
End Sub


Private Sub Timer2_Timer()
Timer2.Enabled = False
Set SysTray.TrayIcon = ilTray.ListImages(10).ExtractIcon
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index
    Case 1 'Block ALL
    frmBlockAll.Show
    Ontop.MakeTopMost frmBlockAll.hwnd
    Case 2 'Applications
    frmApplications.Show , Me
    Case 3 'Logs
    frmLogSecurity.Show , Me
    Case 4
End Select

End Sub


Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Text
Case "Security"
frmLogSecurity.Show , Me
Case "Traffic"
frmLogTraffic.Show , Me
End Select
End Sub


