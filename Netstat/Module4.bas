Attribute VB_Name = "Module4"
'To resolve IP Address to Domain Name

Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, addrlen As Long, addrType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Private Type WSADATA
   wversion As Integer
   wHighVersion As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets As Long
   iMaxUdpDg As Long
   lpVendorInfo As Long
End Type

Private Type IP_OPTION_INFORMATION
   Ttl             As Byte
   Tos             As Byte
   flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type

Private Type ICMP_ECHO_REPLY
   address         As Long
   Status          As Long
   RoundTripTime   As Long
   DataSize        As Long
   Reserved        As Integer
   ptrData         As Long
   Options         As IP_OPTION_INFORMATION
   DATA            As String * 250
End Type

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Private Function SocketsInitialize() As Boolean

    Dim WSAD As WSADATA
    Dim X As Integer
    Dim szLoByte As String
    Dim szHiByte As String
    Dim szBuf As String
    
    X = WSAStartup(WS_VERSION_REQD, WSAD)
    If X <> 0 Then Exit Function 'Sockets isn't responding
    
   'check that the version of sockets is supported
    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or _
       (lobyte(WSAD.wversion) = WS_VERSION_MAJOR And _
        hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
        
        szHiByte = Trim$(Str$(hibyte(WSAD.wversion)))
        szLoByte = Trim$(Str$(lobyte(WSAD.wversion)))
        szBuf = "Windows Sockets Version " & szLoByte & "." & szHiByte
        szBuf = szBuf & " is not supported by Windows " & _
                          "Sockets for 32 bit Windows environments."
        MsgBox szBuf, vbExclamation
        Exit Function
    End If
    
   'check that there are available sockets
    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        'Not have enough sockets
        Exit Function
    End If
    
    SocketsInitialize = True
    
End Function

Private Function hibyte(ByVal wParam As Long) As Integer
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Long) As Integer
    lobyte = wParam And &HFF&
End Function

Public Function ResolveHostname(ByVal ipaddress As String) As String
    Dim hostip_addr As Long
    Dim hostent_addr As Long
    Dim newAddr As Long
    Dim Host As HOSTENT
    Dim strTemp As String
    Dim strHost As String * 255

    If SocketsInitialize() Then
        newAddr = inet_addr(ipaddress)
        hostent_addr = gethostbyaddr(newAddr, Len(newAddr), AF_INET)
        If hostent_addr <> 0 Then
            RtlMoveMemory Host, hostent_addr, Len(Host)
            RtlMoveMemory ByVal strHost, Host.hName, 255
            strTemp = strHost
            
            If InStr(strTemp, Chr(0)) <> 0 Then
                strTemp = Left(strTemp, InStr(strTemp, Chr(0)) - 1)
            End If
            
            ResolveHostname = Trim(strTemp)
        End If
        
        WSACleanup
    End If
End Function

