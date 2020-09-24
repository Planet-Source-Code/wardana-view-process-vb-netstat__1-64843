VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Netstat"
   ClientHeight    =   5580
   ClientLeft      =   240
   ClientTop       =   1530
   ClientWidth     =   11415
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   11415
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   465
      Left            =   10185
      TabIndex        =   1
      Top             =   4995
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   270
      Top             =   4290
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proto"
         Object.Width           =   1060
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Local IP"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "L. Port"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote IP"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DNS Lookup"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "R. Port"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "State"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "PID"
         Object.Width           =   1236
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Processs"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Path"
         Object.Width           =   7832
      EndProperty
   End
   Begin VB.Menu cMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu cRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu batas 
         Caption         =   "-"
      End
      Begin VB.Menu cmTelnet 
         Caption         =   "Telnet Client"
      End
      Begin VB.Menu cStop 
         Caption         =   "Stop Process"
      End
      Begin VB.Menu batas2 
         Caption         =   "-"
      End
      Begin VB.Menu cProperties 
         Caption         =   "File Properties"
      End
   End
   Begin VB.Menu cTools 
      Caption         =   "Tools"
      Begin VB.Menu cTelnet 
         Caption         =   "Telnet Client"
      End
      Begin VB.Menu cServer 
         Caption         =   "Telnet Server"
      End
      Begin VB.Menu batas3 
         Caption         =   "-"
      End
      Begin VB.Menu cStatistic 
         Caption         =   "Statistic"
      End
      Begin VB.Menu cProcess 
         Caption         =   "All  Proccesses"
      End
   End
   Begin VB.Menu cAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indexitem As Integer

Private Sub cAbout_Click()
    MsgBox "Created by Wardana" & vbCrLf & "Bogor, Indonesia" & vbCrLf & "March 29 2006"
End Sub

Private Sub cmTelnet_Click()
    With ListView1.ListItems(indexitem)
        Form2.Text1.Text = .SubItems(3)
        Form2.Text2.Text = .SubItems(5)
    End With
    Form2.Show
End Sub

Private Sub Command1_Click()
    MsgBox "If you like my code, Please vote me", vbOKOnly, "Don't forget"
    End
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form4.Show
End Sub

Private Sub cProcess_Click()
    Form4.Show
End Sub

Private Sub cProperties_Click()
    ShowProps ListView1.ListItems(indexitem).SubItems(9), Me.hwnd
End Sub

Private Sub cRefresh_Click()
    OnRefresh
End Sub

Private Sub cServer_Click()
    Form5.Show
End Sub

Private Sub cStatistic_Click()
    Form3.Show
End Sub

Private Sub cStop_Click()
    Dim cek As Boolean
    If MsgBox("Do you realy want to terminate this process ?", vbOKCancel) = vbOK Then
        cek = Terminate(ListView1.ListItems(indexitem).SubItems(7))
        If cek = False Then Exit Sub
        OnRefresh
        MsgBox "Done"
    End If
End Sub

Private Sub cTelnet_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    mheap = GetProcessHeap()
    OnRefresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MsgBox "If you like my code, Please vote me"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Timer1.Enabled = False
    indexitem = Item.Index
    If ListView1.ListItems(Item.Index).SubItems(3) = "" Then
        cmTelnet.Visible = False
    Else
        cmTelnet.Visible = True
    End If
    PopupMenu cMenu
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If Second(Now) Mod 5 = 0 Then OnRefresh
End Sub


