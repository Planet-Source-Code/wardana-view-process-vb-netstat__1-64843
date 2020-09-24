VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telnet Server"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   5220
      TabIndex        =   6
      Text            =   "280"
      ToolTipText     =   "Port Number"
      Top             =   135
      Width           =   1005
   End
   Begin VB.TextBox Text3 
      Height          =   3885
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   105
      Width           =   4995
   End
   Begin VB.TextBox Text4 
      Height          =   435
      Left            =   135
      TabIndex        =   4
      ToolTipText     =   "Text to send"
      Top             =   4125
      Width           =   4995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listen"
      Height          =   435
      Left            =   5205
      TabIndex        =   3
      Top             =   750
      Width           =   1020
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   435
      Left            =   5220
      TabIndex        =   2
      Top             =   4170
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   420
      Left            =   5220
      TabIndex        =   1
      Top             =   3135
      Width           =   990
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   390
      Left            =   5235
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Server As CSocketMaster
Attribute Server.VB_VarHelpID = -1

Private Sub Command1_Click()
    'There is only one button for Stop and Connect for this application
    If Command1.Caption = "Listen" Then
        Server.Bind Text2.Text
        Server.Listen
        Command1.Caption = "Stop"
        Command2.Enabled = True
    Else
        Server.CloseSck
        Command1.Caption = "Listen"
        Command2.Enabled = False
    End If
End Sub

Private Sub Command2_Click()
    Dim data As String
    data = Text4.Text
    'to send our word to server
    Server.SendData data
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
    Open App.Path & "\TalkServer.txt" For Output As 1
        Print #1, Now
        Print #1, "------------------"
        Print #1, Text3.Text; ""
    Close #1
    MsgBox "Done"
End Sub

Private Sub Form_Load()
    Set Server = New CSocketMaster
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Server.CloseSck
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    'if you press enter button
    If Command2.Enabled = True And KeyAscii = 13 Then Call Command2_Click
End Sub

Private Sub Text4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Text4.SetFocus
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)
    Server.CloseSck
    Server.Accept requestID
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
    Dim data As String
    Server.GetData data
    'Sign arrival paket is (O)
    Text3.Text = Text3.Text + "(O) : " & data & vbCrLf
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    MsgBox Description, vbCritical, "Winsock Error"
    Server.CloseSck
    Command1.Caption = "Connect"
    Command2.Enabled = False
End Sub



