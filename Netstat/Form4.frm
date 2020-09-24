VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proccesses Manager"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   435
      Left            =   2610
      TabIndex        =   8
      ToolTipText     =   "Clear List Manager"
      Top             =   5865
      Width           =   1125
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   435
      Left            =   1335
      TabIndex        =   7
      ToolTipText     =   "Delete File Process in List Manager"
      Top             =   5865
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   435
      Left            =   75
      TabIndex        =   6
      ToolTipText     =   "Stop all process in List Manager"
      Top             =   5865
      Width           =   1170
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1725
      Left            =   45
      TabIndex        =   5
      ToolTipText     =   "List Manager"
      Top             =   4065
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3043
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process Address"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size (kb)"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date & Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pid"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   6510
      TabIndex        =   4
      Top             =   6240
      Width           =   930
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   435
      Top             =   5370
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7275
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3690
      Width           =   2925
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   3675
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   7260
   End
   Begin VB.Menu cMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu cKill 
         Caption         =   "Kill"
      End
      Begin VB.Menu cList 
         Caption         =   "Add to List Manager"
      End
      Begin VB.Menu cRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu cProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tmp As Integer 'for List1Index value when we click the List1

Private Sub cKill_Click()
    On Error Resume Next
    Dim ya As String
    Dim hProcess As Long
    ya = MsgBox(List1.List(tmp), 1, "To be Closed?")
    If ya = 1 Then
        Terminate List1.ItemData(tmp)
        List1.RemoveItem tmp
    End If
    List1.Refresh
End Sub

Private Sub cList_Click()
    On Error Resume Next
    Dim utama As ListItem
    
    Set utama = ListView1.ListItems.Add(, , List1.List(tmp))
    utama.SubItems(1) = Str(Round(Val(FileLen(List1.List(tmp)) / 1024), 4))
    utama.SubItems(2) = FileDateTime(List1.List(tmp))
    utama.SubItems(3) = List1.ItemData(tmp)
End Sub

Private Sub Command1_Click()
    Me.Hide
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Dim i As Integer
    Dim cek As Boolean
    If ListView1.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Do you want to stop all of this processes in List Manager ?", vbOKCancel) = vbCancel Then Exit Sub
    For i = 1 To ListView1.ListItems.Count
        cek = Terminate(ListView1.ListItems.Item(i).SubItems(3))
    Next
    Update
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim i As Integer
    If ListView1.ListItems.Count = 0 Then Exit Sub
    Call Command2_Click
    For i = 1 To ListView1.ListItems.Count
        RecycleBin ListView1.ListItems.Item(i).Text
    Next
End Sub

Private Sub Command4_Click()
    ListView1.ListItems.Clear
End Sub

Private Sub cProperties_Click()
    'to show properties-dialogbox of file
    ShowProps List1.List(tmp), Me.hwnd
End Sub

Private Sub cRefresh_Click()
    Update
End Sub

Private Sub Form_Load()
    Update
    Keterangan 0
End Sub

Private Sub List1_Click()
    tmp = List1.ListIndex
    Keterangan tmp
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu cMenu
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If MsgBox("Do you want to remove this item ?" & vbCrLf & Item, vbOKCancel, "Confirmation") = vbOK Then
        ListView1.ListItems.Remove Item.Index
    End If
End Sub

Private Sub Timer1_Timer()
    'update our list every 5 second
    Timer1.Enabled = False
    If Second(Now) Mod 5 = 0 Then Update
    Timer1.Enabled = True
End Sub

Private Sub Keterangan(nilai As Integer)
    On Error Resume Next
    'To give size and time information the file
    If List1.List(nilai) = "" Then Exit Sub
    Label1.Caption = List1.List(nilai)
    Label2.Caption = "Size : " & Round(FileLen(List1.List(nilai)) / 1024, 4) & " kb"
    Label3.Caption = "Date Time : " & FileDateTime(List1.List(nilai))
End Sub
