Attribute VB_Name = "Module3"
'to get statistic of pakets
'From CS Bandwidth Monitor By Shane M. Croft Of Crofts Software.

Type MIB_TCPROW
  dwState As Long        'state of the connection
  dwLocalAddr As String * 4    'address on local computer
  dwLocalPort As String * 4    'port number on local computer
  dwRemoteAddr As String * 4   'address on remote computer
  dwRemotePort As String * 4   'port number on remote computer
End Type

Type MIB_TCPTABLE
  dwNumEntries As Long    'number of entries in the table
  table(100) As MIB_TCPROW   'array of TCP connections
End Type

Public Declare Function GetTcpTable Lib "IPhlpAPI" (pTcpTable As MIB_TCPTABLE, pdwSize As Long, bOrder As Long) As Long

Type MIB_UDPROW
  dwLocalAddr As String * 4 'address on local computer
  dwLocalPort As String * 4 'port number on local computer
End Type

Type MIB_UDPTABLE
  dwNumEntries As Long    'number of entries in the table
  table(100) As MIB_UDPROW   'table of MIB_UDPROW structs
End Type

Public Declare Function GetUdpTable Lib "IPhlpAPI" (pUdpTable As MIB_UDPTABLE, pdwSize As Long, bOrder As Long) As Long

Type MIB_IPSTATS
  dwForwarding As Long       ' IP forwarding enabled or disabled
  dwDefaultTTL As Long       ' default time-to-live
  dwInReceives As Long       ' datagrams received
  dwInHdrErrors As Long      ' received header errors
  dwInAddrErrors As Long     ' received address errors
  dwForwDatagrams As Long    ' datagrams forwarded
  dwInUnknownProtos As Long  ' datagrams with unknown protocol
  dwInDiscards As Long       ' received datagrams discarded
  dwInDelivers As Long       ' received datagrams delivered
  dwOutRequests As Long      '
  dwRoutingDiscards As Long  '
  dwOutDiscards As Long      ' sent datagrams discarded
  dwOutNoRoutes As Long      ' datagrams for which no route
  dwReasmTimeout As Long     ' datagrams for which all frags didn't arrive
  dwReasmReqds As Long       ' datagrams requiring reassembly
  dwReasmOks As Long         ' successful reassemblies
  dwReasmFails As Long       ' failed reassemblies
  dwFragOks As Long          ' successful fragmentations
  dwFragFails As Long        ' failed fragmentations
  dwFragCreates As Long      ' datagrams fragmented
  dwNumIf As Long           ' number of interfaces on computer
  dwNumAddr As Long         ' number of IP address on computer
  dwNumRoutes As Long       ' number of routes in routing table
End Type

Public Declare Function GetIpStatistics Lib "IPhlpAPI" (pStats As MIB_IPSTATS) As Long

Type MIBICMPSTATS
  dwMsgs As Long            ' number of messages
  dwErrors As Long          ' number of errors
  dwDestUnreachs As Long    ' destination unreachable messages
  dwTimeExcds As Long       ' time-to-live exceeded messages
  dwParmProbs As Long       ' parameter problem messages
  dwSrcQuenchs As Long      ' source quench messages
  dwRedirects As Long       ' redirection messages
  dwEchos As Long           ' echo requests
  dwEchoReps As Long        ' echo replies
  dwTimestamps As Long      ' timestamp requests
  dwTimestampReps As Long   ' timestamp replies
  dwAddrMasks As Long       ' address mask requests
  dwAddrMaskReps As Long    ' address mask replies
End Type

 Type MIBICMPINFO
  icmpInStats As MIBICMPSTATS        ' stats for incoming messages
  icmpOutStats As MIBICMPSTATS       ' stats for outgoing messages
End Type

Public Declare Function GetIcmpStatistics Lib "IPhlpAPI" (pStats As MIBICMPINFO) As Long

Type MIB_TCPSTATS
  dwRtoAlgorithm As Long    ' timeout algorithm
  dwRtoMin As Long          ' minimum timeout
  dwRtoMax As Long          ' maximum timeout
  dwMaxConn As Long         ' maximum connections
  dwActiveOpens As Long     ' active opens
  dwPassiveOpens As Long    ' passive opens
  dwAttemptFails As Long    ' failed attempts
  dwEstabResets As Long     ' establised connections reset
  dwCurrEstab As Long       ' established connections
  dwInSegs As Long          ' segments received
  dwOutSegs As Long         ' segment sent
  dwRetransSegs As Long     ' segments retransmitted
  dwInErrs As Long          ' incoming errors
  dwOutRsts As Long         ' outgoing resets
  dwNumConns As Long        ' cumulative connections
End Type

Public Declare Function GetTcpStatistics Lib "IPhlpAPI" (pStats As MIB_TCPSTATS) As Long

Type MIB_UDPSTATS
  dwInDatagrams As Long    ' received datagrams
  dwNoPorts As Long        ' datagrams for which no port
  dwInErrors As Long       ' errors on received datagrams
  dwOutDatagrams As Long   ' sent datagrams
  dwNumAddrs As Long       ' number of entries in UDP listener table
End Type

Public Declare Function GetUdpStatistics Lib "IPhlpAPI" (pStats As MIB_UDPSTATS) As Long


