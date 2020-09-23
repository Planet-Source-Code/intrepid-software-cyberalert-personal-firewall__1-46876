VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogSecurity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security Log"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   5640
      Width           =   735
   End
   Begin VB.TextBox txtDesc 
      Height          =   975
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5040
      Width           =   5295
   End
   Begin MSComctlLib.ListView lvLog 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8705
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time Stamp"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Severity"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   10583
      EndProperty
   End
End
Attribute VB_Name = "frmLogSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Kill App.Path & "\Logs\Security" & AscDate & ".txt"
txtDesc.Text = ""
lvLog.ListItems.Clear
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
'Load Log File
Dim ff As Integer
Dim strInput As String
Dim strTemp As String
Dim num1 As Integer, num2 As Long

Dim Timestamp As String
Dim Severity As String
Dim Description As String

Dim Item As ListItem

On Error GoTo errer
ff = FreeFile
Open App.Path & "\Logs\Security\" & AscDate & ".txt" For Input As #ff

Do While EOF(ff) = False
    Line Input #ff, strInput
    num1 = InStr(1, strInput, "]") + 1
    strTemp = Mid(strInput, 2, num1 - 3)
    Timestamp = strTemp
    num2 = InStr(1, strInput, " -")
    strTemp = Mid(strInput, num1 + 1, num2 - num1)
    Severity = strTemp
    num1 = InStr(1, strInput, "- ") + 2
    num2 = Len(strInput)
    strTemp = Mid(strInput, num1, num2 - num1)
    Description = strTemp
    Set Item = lvLog.ListItems.Add()
    Item.Text = Timestamp
    Item.SubItems(1) = Severity
    Item.SubItems(2) = Description
    
    
DoEvents
Loop
Close #ff
errer:

End Sub


Private Sub lvLog_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtDesc.Text = Item.SubItems(2)
End Sub

