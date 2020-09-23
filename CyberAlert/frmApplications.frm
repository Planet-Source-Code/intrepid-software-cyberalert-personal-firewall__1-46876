VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmApplications 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Applications"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmApplications.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Modify"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvApplications 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Application Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Version"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Access"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   9596
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   $"frmApplications.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Menu mnuModify 
      Caption         =   "mnuModify"
      Visible         =   0   'False
      Begin VB.Menu mnuModifyTrust 
         Caption         =   "Trust"
      End
      Begin VB.Menu mnuModifyAsk 
         Caption         =   "Ask"
      End
      Begin VB.Menu mnuModifyBlock 
         Caption         =   "Block"
      End
      Begin VB.Menu mnuModifySep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuModifyDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmApplications"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function GetRule(Value As String) As String
Select Case Value
Case "0" 'ask
GetRule = "Ask"
Case "1" 'block
GetRule = "Block"
Case "2" 'trust
GetRule = "Trust"
End Select

End Function

Private Sub Command2_Click()

Unload Me
End Sub


Private Sub Command3_Click()
On Error Resume Next
Dim idNum As Integer
Dim strTemp As String
idNum = lvApplications.SelectedItem.Index
If idNum <> 0 Then
strTemp = lvApplications.SelectedItem.Text
If strTemp = "" Then strTemp = Mid(lvApplications.SelectedItem.SubItems(3), InStrRev(lvApplications.SelectedItem.SubItems(3), "\") + 1)
mnuModifyTrust.Caption = "Always allow " & strTemp
mnuModifyBlock.Caption = "Always block " & strTemp
mnuModifyAsk.Caption = "Ask for " & strTemp
mnuModifyDelete.Caption = "Remove " & strTemp & " from this list."
PopupMenu mnuModify
End If

End Sub

Private Sub Form_Load()
Dim strParse() As String
Dim i As Integer
Dim item As ListItem
cINIFile.Section = "RULES"
strParse = Split(cINIFile.INISection, Chr(0))
For i = 0 To UBound(strParse) - 1
cINIFile.Key = strParse(i)

If cINIFile.Value <> "" Then
Set item = lvApplications.ListItems.Add()
item.Text = GetFileDescription(strParse(i))
item.SubItems(2) = GetRule(cINIFile.Value)
item.SubItems(1) = GetFileVersion(strParse(i))
item.SubItems(3) = strParse(i)
End If

Next i


End Sub

Private Sub Form_Unload(Cancel As Integer)
mdlFirewall.Execute True
End Sub

Private Sub mnuModifyAsk_Click()
'Add rule to ask this
cINIFile.Section = "RULES"
cINIFile.Key = lvApplications.SelectedItem.SubItems(3)
cINIFile.Value = "0"
lvApplications.SelectedItem.SubItems(2) = "Ask"
End Sub


Private Sub mnuModifyBlock_Click()
'Add rule to block this
cINIFile.Section = "RULES"
cINIFile.Key = lvApplications.SelectedItem.SubItems(3)
cINIFile.Value = "1"
lvApplications.SelectedItem.SubItems(2) = "Block"
End Sub


Private Sub mnuModifyDelete_Click()
'delete this
cINIFile.Section = "RULES"
cINIFile.Key = lvApplications.SelectedItem.SubItems(3)
cINIFile.DeleteValue
lvApplications.ListItems.Remove lvApplications.SelectedItem.Index
End Sub


Private Sub mnuModifyTrust_Click()
'Add rule to trust this
cINIFile.Section = "RULES"
cINIFile.Key = lvApplications.SelectedItem.SubItems(3)
cINIFile.Value = "2"
lvApplications.SelectedItem.SubItems(2) = "Trust"
End Sub


