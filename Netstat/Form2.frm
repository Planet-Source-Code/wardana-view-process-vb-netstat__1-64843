VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telnet Client"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   390
      Left            =   5205
      TabIndex        =   7
      Top             =   3750
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   420
      Left            =   5190
      TabIndex        =   6
      Top             =   3285
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5190
      TabIndex        =   5
      Top             =   4320
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   435
      Left            =   5175
      TabIndex        =   4
      Top             =   720
      Width           =   1020
   End
   Begin VB.TextBox Text4 
      Height          =   435
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "Text to send"
      Top             =   4275
      Width           =   4995
   End
   Begin VB.TextBox Text3 
      Height          =   3420
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   720
      Width           =   4995
   End
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   5175
      TabIndex        =   1
      Text            =   "280"
      ToolTipText     =   "Port Number"
      Top             =   150
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Text            =   "127.0.0.1"
      ToolTipText     =   "Adress "
      Top             =   150
      Width           =   4950
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents klien As CSocketMaster
Attribute klien.VB_VarHelpID = -1

Private Sub Command1_Click()
    If Text1.Text = "" Or Text2.Text = "" Then
        MsgBox "Please, Fill the address and port number"
        Exit Sub
    End If
    'There is only one button for Stop and Connect for this application
    If Command1.Caption = "Connect" Then
        klien.Connect Text1.Text, Text2.Text
        Command1.Caption = "Stop"
        Command2.Enabled = True
    Else
        klien.CloseSck
        Command1.Caption = "Connect"
        Command2.Enabled = False
    End If
End Sub

Private Sub Command2_Click()
    Dim data As String
    data = Text4.Text
    'to send our word to server
    klien.SendData data
    'Sign (X) is the paket which I have send
    Text3.Text = Text3.Text + "(X) : " & data & vbCrLf
    Text4.Text = ""
End Sub

Private Sub Command3_Click()
    If Text3.Text = "" Then Exit Sub
    If MsgBox("Do you want to Delete all ?", vbOKCancel) = vbOK Then Text3.Text = ""
End Sub

Private Sub Command4_Click()
    If Text3.Text = "" Then Exit Sub
    'to make file from text in text3
    Open App.Path & "\TalkClient.txt" For Output As 1
        Print #1, Now
        Print #1, "------------------"
        Print #1, Text3.Text; ""
    Close #1
    MsgBox "Done"
End Sub

Private Sub Form_Load()
    Set klien = New CSocketMaster
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    klien.CloseSck
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    'if you prees enter button
    If Command2.Enabled = True And KeyAscii = 13 Then Call Command2_Click
End Sub

Private Sub Text4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Text4.SetFocus
End Sub

Private Sub klien_Connect()
    MsgBox "You Connect to " & Text1.Text
End Sub

Private Sub klien_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    klien.GetData data
    'Sign arrival paket is (O)
    Text3.Text = Text3.Text + "(O) : " & data & vbCrLf
End Sub

Private Sub klien_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    MsgBox Description, vbCritical, "Winsock Error"
    klien.CloseSck
    Command1.Caption = "Connect"
    Command2.Enabled = False
End Sub


