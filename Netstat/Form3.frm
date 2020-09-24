VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Statistic"
   ClientHeight    =   7260
   ClientLeft      =   135
   ClientTop       =   960
   ClientWidth     =   8220
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8220
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1305
      Top             =   3210
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   3015
      Left            =   105
      TabIndex        =   3
      Top             =   4080
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Input"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Output"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView ListView4 
      Height          =   4035
      Left            =   4185
      TabIndex        =   2
      Top             =   3075
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   7117
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2025
      Left            =   4170
      TabIndex        =   1
      Top             =   420
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3572
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   90
      TabIndex        =   0
      Top             =   435
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Parameter"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ICMP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   7
      Top             =   3690
      Width           =   3900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   4200
      TabIndex        =   6
      Top             =   2715
      Width           =   3900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UDP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   4170
      TabIndex        =   5
      Top             =   60
      Width           =   3990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TCP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   75
      Width           =   3900
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IP As MIB_IPSTATS
Dim tcp As MIB_TCPSTATS
Dim udp As MIB_UDPSTATS
Dim icmp As MIBICMPINFO
Dim tStats As MIB_TCPSTATS

Private Sub Form_Load()
    'TCP
    With ListView1.ListItems
        .Add , , "Timeout algorithm"
        .Add , , "Minimum timeout"
        .Add , , "Maximum timeout"
        .Add , , "Maximum connections"
        .Add , , "Active opens"
        .Add , , "Passive opens"
        .Add , , "Failed attempts"
        .Add , , "Establised connections reset"
        .Add , , "Established connections"
        .Add , , "Segments received"
        .Add , , "Segment sent"
        .Add , , "Segments retransmitted"
        .Add , , "Incoming errors"
        .Add , , "Outgoing resets"
        .Add , , "Cumulative connections"
    End With
    
    'UDP
    With ListView2.ListItems
        .Add , , "received datagrams"
        .Add , , "datagrams for which no port"
        .Add , , "errors on received datagrams"
        .Add , , "sent datagrams"
        .Add , , "number of entries in UDP listener table"
    End With
    
     'IP
    With ListView3.ListItems
        .Add , , "number of messages"
        .Add , , "number of errors"
        .Add , , "destination unreachable messages"
        .Add , , "time-to-live exceeded messages"
        .Add , , "parameter problem messages"
        .Add , , "source quench messages"
        .Add , , "redirection messages"
        .Add , , "echo requests"
        .Add , , "echo replies"
        .Add , , "timestamp requests"
        .Add , , "timestamp replies"
        .Add , , "address mask requests"
        .Add , , "address mask replies"
    End With
    
    With ListView4.ListItems
        .Add , , "IP forwarding enabled or disabled"
        .Add , , "Default time-to-live"
        .Add , , "Datagrams received"
        .Add , , "Received header errors"
        .Add , , "Received address errors"
        .Add , , "datagrams forwarded"
        .Add , , "datagrams with unknown protocol"
        .Add , , "received datagrams discarded"
        .Add , , "received datagrams delivered"
        .Add , , "outgoing datagrams requested"
        .Add , , "outgoing datagrams discarded"
        .Add , , "sent datagrams discarded"
        .Add , , "datagrams for which no route"
        .Add , , "datagrams for which all frags didn't arrive"
        .Add , , "datagrams requiring reassembly"
        .Add , , "successful reassemblies"
        .Add , , "failed reassemblies"
        .Add , , "successful fragmentations"
        .Add , , "failed fragmentations"
        .Add , , "datagrams fragmented"
        .Add , , "number of interfaces on computer"
        .Add , , "number of IP address on computer"
        .Add , , "number of routes in routing table"
    End With
    UpdateStats1
    UpdateStats2
    UpdateStats3a
    UpdateStats3b
    UpdateStats4
End Sub

Private Sub UpdateStats1()
    On Error Resume Next
    Dim tStats          As MIB_TCPSTATS
    Static tStaticStats As MIB_TCPSTATS
    Dim lRetValue       As Long
    
    lRetValue = GetTcpStatistics(tStats)
    With tStats
        If Not tStaticStats.dwRtoAlgorithm = .dwRtoAlgorithm Then _
            ListView1.ListItems(1).SubItems(1) = .dwRtoAlgorithm
        If Not tStaticStats.dwRtoMin = .dwRtoMin Then _
            ListView1.ListItems(2).SubItems(1) = .dwRtoMin
        If Not tStaticStats.dwRtoMax = .dwRtoMax Then _
            ListView1.ListItems(3).SubItems(1) = .dwRtoMax
        If Not tStaticStats.dwMaxConn = .dwMaxConn Then _
            ListView1.ListItems(4).SubItems(1) = .dwMaxConn
        If Not tStaticStats.dwActiveOpens = .dwActiveOpens Then _
            ListView1.ListItems(5).SubItems(1) = .dwActiveOpens
        If Not tStaticStats.dwPassiveOpens = .dwPassiveOpens Then _
            ListView1.ListItems(6).SubItems(1) = .dwPassiveOpens
        If Not tStaticStats.dwAttemptFails = .dwAttemptFails Then _
            ListView1.ListItems(7).SubItems(1) = .dwAttemptFails
        If Not tStaticStats.dwEstabResets = .dwEstabResets Then _
            ListView1.ListItems(8).SubItems(1) = .dwEstabResets
        If Not tStaticStats.dwCurrEstab = .dwCurrEstab Then _
            ListView1.ListItems(9).SubItems(1) = .dwCurrEstab
        If Not tStaticStats.dwInSegs = .dwInSegs Then _
            ListView1.ListItems(10).SubItems(1) = .dwInSegs
        If Not tStaticStats.dwOutSegs = .dwOutSegs Then _
            ListView1.ListItems(11).SubItems(1) = .dwOutSegs
        If Not tStaticStats.dwRetransSegs = .dwRetransSegs Then _
            ListView1.ListItems(12).SubItems(1) = .dwRetransSegs
        If Not tStaticStats.dwInErrs = .dwInErrs Then _
            ListView1.ListItems(13).SubItems(1) = .dwInErrs
        If Not tStaticStats.dwOutRsts = .dwOutRsts Then _
            ListView1.ListItems(14).SubItems(1) = .dwOutRsts
        If Not tStaticStats.dwNumConns = .dwNumConns Then _
            ListView1.ListItems(15).SubItems(1) = .dwNumConns
    End With
    Repair ListView1, 1
End Sub


Private Sub UpdateStats2()
    On Error Resume Next
    Dim lRetValue       As Long
    Static udp2 As MIB_UDPSTATS
    
    lRetValue = GetUdpStatistics(udp)
    
    With udp
        If Not udp2.dwInDatagrams = .dwInDatagrams Then _
        ListView2.ListItems(1).SubItems(1) = .dwInDatagrams
        If Not udp2.dwNoPorts = .dwNoPorts Then _
        ListView2.ListItems(2).SubItems(1) = .dwNoPorts
        If Not udp2.dwInErrors = .dwInErrors Then _
        ListView2.ListItems(3).SubItems(1) = .dwInErrors
        If Not udp2.dwOutDatagrams = .dwOutDatagrams Then _
        ListView2.ListItems(4).SubItems(1) = .dwOutDatagrams
        If Not udp2.dwNumAddrs = .dwNumAddrs Then _
        ListView2.ListItems(5).SubItems(1) = .dwNumAddrs
    End With
    Repair ListView2, 1
End Sub
Private Sub UpdateStats3a()
    On Error Resume Next
    Dim lRetValue       As Long
    Static icmp2 As MIBICMPINFO
    
    lRetValue = GetIcmpStatistics(icmp)
    
    With icmp
        If Not icmp2.icmpOutStats.dwMsgs = .icmpOutStats.dwMsgs Then _
        ListView3.ListItems(1).SubItems(1) = .icmpOutStats.dwMsgs
        If Not icmp2.icmpOutStats.dwErrors = .icmpOutStats.dwErrors Then _
        ListView3.ListItems(2).SubItems(1) = .icmpOutStats.dwErrors
        If Not icmp2.icmpOutStats.dwDestUnreachs = .icmpOutStats.dwDestUnreachs Then _
        ListView3.ListItems(3).SubItems(1) = .icmpOutStats.dwDestUnreachs
        If Not icmp2.icmpOutStats.dwTimeExcds = .icmpOutStats.dwTimeExcds Then _
        ListView3.ListItems(4).SubItems(1) = .icmpOutStats.dwTimeExcds
        If Not icmp2.icmpOutStats.dwParmProbs = .icmpOutStats.dwParmProbs Then _
        ListView3.ListItems(5).SubItems(1) = .icmpOutStats.dwParmProbs
        If Not icmp2.icmpOutStats.dwSrcQuenchs = .icmpOutStats.dwSrcQuenchs Then _
        ListView3.ListItems(6).SubItems(1) = .icmpOutStats.dwSrcQuenchs
        If Not icmp2.icmpOutStats.dwRedirects = .icmpOutStats.dwRedirects Then _
        ListView3.ListItems(7).SubItems(1) = .icmpOutStats.dwRedirects
        If Not icmp2.icmpOutStats.dwEchos = .icmpOutStats.dwEchos Then _
        ListView3.ListItems(8).SubItems(1) = .icmpOutStats.dwEchos
        If Not icmp2.icmpOutStats.dwEchoReps = .icmpOutStats.dwEchoReps Then _
        ListView3.ListItems(9).SubItems(1) = .icmpOutStats.dwEchoReps
        If Not icmp2.icmpOutStats.dwTimestamps = .icmpOutStats.dwTimestamps Then _
        ListView3.ListItems(10).SubItems(1) = .icmpOutStats.dwTimestamps
        If Not icmp2.icmpOutStats.dwTimestampReps = .icmpOutStats.dwTimestampReps Then _
        ListView3.ListItems(11).SubItems(1) = .icmpOutStats.dwTimestampReps
        If Not icmp2.icmpOutStats.dwAddrMasks = .icmpOutStats.dwAddrMasks Then _
        ListView3.ListItems(12).SubItems(1) = .icmpOutStats.dwAddrMasks
        If Not icmp2.icmpOutStats.dwAddrMaskReps = .icmpOutStats.dwAddrMaskReps Then _
        ListView3.ListItems(13).SubItems(1) = .icmpOutStats.dwAddrMaskReps
    End With
    Repair ListView3, 1
End Sub
Private Sub UpdateStats3b()

    On Error Resume Next
    Dim lRetValue       As Long
    Static icmp2 As MIBICMPINFO
    
    lRetValue = GetIcmpStatistics(icmp)
    
    With icmp
        If Not icmp2.icmpInStats.dwMsgs = .icmpInStats.dwMsgs Then _
        ListView3.ListItems(1).SubItems(2) = .icmpInStats.dwMsgs
        If Not icmp2.icmpInStats.dwErrors = .icmpInStats.dwErrors Then _
        ListView3.ListItems(2).SubItems(2) = .icmpInStats.dwErrors
        If Not icmp2.icmpInStats.dwDestUnreachs = .icmpInStats.dwDestUnreachs Then _
        ListView3.ListItems(3).SubItems(2) = .icmpInStats.dwDestUnreachs
        If Not icmp2.icmpInStats.dwTimeExcds = .icmpInStats.dwTimeExcds Then _
        ListView3.ListItems(4).SubItems(2) = .icmpInStats.dwTimeExcds
        If Not icmp2.icmpInStats.dwParmProbs = .icmpInStats.dwParmProbs Then _
        ListView3.ListItems(5).SubItems(2) = .icmpInStats.dwParmProbs
        If Not icmp2.icmpInStats.dwSrcQuenchs = .icmpInStats.dwSrcQuenchs Then _
        ListView3.ListItems(6).SubItems(2) = .icmpInStats.dwSrcQuenchs
        If Not icmp2.icmpInStats.dwRedirects = .icmpInStats.dwRedirects Then _
        ListView3.ListItems(7).SubItems(2) = .icmpInStats.dwRedirects
        If Not icmp2.icmpInStats.dwEchos = .icmpInStats.dwEchos Then _
        ListView3.ListItems(8).SubItems(2) = .icmpInStats.dwEchos
        If Not icmp2.icmpInStats.dwEchoReps = .icmpInStats.dwEchoReps Then _
        ListView3.ListItems(9).SubItems(2) = .icmpInStats.dwEchoReps
        If Not icmp2.icmpInStats.dwTimestamps = .icmpInStats.dwTimestamps Then _
        ListView3.ListItems(10).SubItems(2) = .icmpInStats.dwTimestamps
        If Not icmp2.icmpInStats.dwTimestampReps = .icmpInStats.dwTimestampReps Then _
        ListView3.ListItems(11).SubItems(2) = .icmpInStats.dwTimestampReps
        If Not icmp2.icmpInStats.dwAddrMasks = .icmpInStats.dwAddrMasks Then _
        ListView3.ListItems(12).SubItems(2) = .icmpInStats.dwAddrMasks
        If Not icmp2.icmpInStats.dwAddrMaskReps = .icmpInStats.dwAddrMaskReps Then _
        ListView3.ListItems(13).SubItems(2) = .icmpInStats.dwAddrMaskReps
    End With
    Repair ListView3, 2
End Sub

Private Sub UpdateStats4()

    On Error Resume Next
    Static ip2 As MIB_IPSTATS
    Dim lRetValue       As Long
    
    lRetValue = GetIpStatistics(IP)
    
    With IP
        If Not ip2.dwForwarding = .dwForwarding Then _
        ListView4.ListItems(1).SubItems(1) = .dwForwarding
        If Not ip2.dwDefaultTTL = .dwDefaultTTL Then _
        ListView4.ListItems(2).SubItems(1) = .dwDefaultTTL
        If Not ip2.dwInReceives = .dwInReceives Then _
        ListView4.ListItems(3).SubItems(1) = .dwInReceives
        If Not ip2.dwInHdrErrors = .dwInHdrErrors Then _
        ListView4.ListItems(4).SubItems(1) = .dwInHdrErrors
        If Not ip2.dwInAddrErrors = .dwInAddrErrors Then _
        ListView4.ListItems(5).SubItems(1) = .dwInAddrErrors
        If Not ip2.dwForwDatagrams = .dwForwDatagrams Then _
        ListView4.ListItems(6).SubItems(1) = .dwForwDatagrams
        If Not ip2.dwInUnknownProtos = .dwInUnknownProtos Then _
        ListView4.ListItems(7).SubItems(1) = .dwInUnknownProtos
        If Not ip2.dwInDiscards = .dwInDiscards Then _
        ListView4.ListItems(8).SubItems(1) = .dwInDiscards
        If Not ip2.dwInDelivers = .dwInDelivers Then _
        ListView4.ListItems(9).SubItems(1) = .dwInDelivers
        If Not ip2.dwOutRequests = .dwOutRequests Then _
        ListView4.ListItems(10).SubItems(1) = .dwOutRequests
        If Not ip2.dwRoutingDiscards = .dwRoutingDiscards Then _
        ListView4.ListItems(11).SubItems(1) = .dwRoutingDiscards
        If Not ip2.dwOutDiscards = .dwOutDiscards Then _
        ListView4.ListItems(12).SubItems(1) = .dwOutDiscards
        If Not ip2.dwOutNoRoutes = .dwOutNoRoutes Then _
        ListView4.ListItems(13).SubItems(1) = .dwOutNoRoutes
        If Not ip2.dwReasmTimeout = .dwReasmTimeout Then _
        ListView4.ListItems(14).SubItems(1) = .dwReasmTimeout
        If Not ip2.dwReasmReqds = .dwReasmReqds Then _
        ListView4.ListItems(15).SubItems(1) = .dwReasmReqds
        If Not ip2.dwReasmOks = .dwReasmOks Then _
        ListView4.ListItems(16).SubItems(1) = .dwReasmOks
        If Not ip2.dwReasmFails = .dwReasmFails Then _
        ListView4.ListItems(17).SubItems(1) = .dwReasmFails
        If Not ip2.dwFragOks = .dwFragOks Then _
        ListView4.ListItems(18).SubItems(1) = .dwFragOks
        If Not ip2.dwFragFails = .dwFragFails Then _
        ListView4.ListItems(19).SubItems(1) = .dwFragFails
        If Not ip2.dwFragCreates = .dwFragCreates Then _
        ListView4.ListItems(20).SubItems(1) = .dwFragCreates
        If Not ip2.dwNumIf = .dwNumIf Then _
        ListView4.ListItems(21).SubItems(1) = .dwNumIf
        If Not ip2.dwNumAddr = .dwNumAddr Then _
        ListView4.ListItems(22).SubItems(1) = .dwNumAddr
        If Not ip2.dwNumRoutes = .dwNumRoutes Then _
        ListView4.ListItems(23).SubItems(1) = .dwNumRoutes
    End With
    Repair ListView4, 1
End Sub

Private Sub Repair(Lsv As ListView, Item As Integer)
    Dim i As Integer
    For i = 1 To Lsv.ListItems.Count
        If Lsv.ListItems(i).SubItems(Item) = "" Then
            Lsv.ListItems(i).SubItems(Item) = "0"
        End If
    Next
End Sub

Private Sub Timer1_Timer()
    If Second(Now) Mod 5 = 0 Then
        Timer1.Enabled = False
        UpdateStats1
        UpdateStats2
        UpdateStats3a
        UpdateStats3b
        UpdateStats4
        Timer1.Enabled = True
    End If
End Sub
