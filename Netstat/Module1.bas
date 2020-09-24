Attribute VB_Name = "Module1"
'For netstat
Private Const PROCESS_VM_READ           As Long = &H10
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_TERMINATE As Long = (&H1)

Private Const MIB_TCP_STATE_CLOSED      As Long = 1
Private Const MIB_TCP_STATE_LISTEN      As Long = 2
Private Const MIB_TCP_STATE_SYN_SENT    As Long = 3
Private Const MIB_TCP_STATE_SYN_RCVD    As Long = 4
Private Const MIB_TCP_STATE_ESTAB       As Long = 5
Private Const MIB_TCP_STATE_FIN_WAIT1   As Long = 6
Private Const MIB_TCP_STATE_FIN_WAIT2   As Long = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT  As Long = 8
Private Const MIB_TCP_STATE_CLOSING     As Long = 9
Private Const MIB_TCP_STATE_LAST_ACK    As Long = 10
Private Const MIB_TCP_STATE_TIME_WAIT   As Long = 11
Private Const MIB_TCP_STATE_DELETE_TCB  As Long = 12
    
Private Type PMIB_UDPEXROW
    dwLocalAddr     As Long
    dwLocalPort     As Long
    dwProcessId     As Long
End Type

Private Type PMIB_TCPEXROW
    dwStats         As Long
    dwLocalAddr     As Long
    dwLocalPort     As Long
    dwRemoteAddr    As Long
    dwRemotePort    As Long
    dwProcessId     As Long
End Type

Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef bOrder As Boolean, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long
Private Declare Function AllocateAndGetUdpExTableFromStack Lib "iphlpapi.dll" (ByRef pTcpTable As Any, ByRef bOrder As Boolean, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function GetProcessHeap Lib "kernel32.dll" () As Long

Private Declare Function EnumProcesses Lib "psapi.dll" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long

Public mheap As Long

'to know all of processes in form4
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Dim AppActive(100) As String
Dim RemoveApp(16) As String
Dim NumberOfProg As String
Dim counter As Integer
Dim BlinkThread(16) As Boolean

Public Sub OnRefresh()
    'to refresh list of all port process
    Dim TcpExTable() As PMIB_TCPEXROW
    Dim UdpExTable() As PMIB_UDPEXROW
    Dim Distant      As String
    Dim Pointer      As Long
    Dim Number       As Long
    Dim Size         As Long
    Dim i            As Long
    
    Form1.ListView1.ListItems.Clear
    'for TCP
    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim TcpExTable(Number - 1)
            Size = Number * Len(TcpExTable(0))
            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(TcpExTable)
                With Form1.ListView1.ListItems.Add
                    .Text = "TCP"
                    .SubItems(1) = GetIpString(TcpExTable(i).dwLocalAddr)
                    .SubItems(2) = GetPortNumber(TcpExTable(i).dwLocalPort)
                    If GetIpString(TcpExTable(i).dwRemoteAddr) = "0.0.0.0" Then
                        .SubItems(3) = ""
                        .SubItems(4) = ""
                        .SubItems(5) = ""
                    Else
                        .SubItems(3) = GetIpString(TcpExTable(i).dwRemoteAddr)
                        .SubItems(4) = ResolveHostname(GetIpString(TcpExTable(i).dwRemoteAddr))
                        .SubItems(5) = GetPortNumber(TcpExTable(i).dwRemotePort)
                    End If
                    .SubItems(6) = GetState(TcpExTable(i).dwStats)
                    .SubItems(7) = TcpExTable(i).dwProcessId
                    .SubItems(8) = GetProcessName(TcpExTable(i).dwProcessId)
                    .SubItems(9) = ProcessPathByPID(TcpExTable(i).dwProcessId)
                End With
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    Else
        MsgBox "Can't get TCP table", vbExclamation
    End If
    
    'For UDP
    If AllocateAndGetUdpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim UdpExTable(Number - 1)
            Size = Number * Len(UdpExTable(0))
            CopyMemory UdpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(UdpExTable)
                With Form1.ListView1.ListItems.Add
                    .Text = "UDP"
                    .SubItems(1) = GetIpString(UdpExTable(i).dwLocalAddr)
                    .SubItems(2) = GetPortNumber(UdpExTable(i).dwLocalPort)
                    .SubItems(3) = ""
                    .SubItems(4) = ""
                    .SubItems(5) = ""
                    .SubItems(6) = "LISTEN"
                    .SubItems(7) = UdpExTable(i).dwProcessId
                    .SubItems(8) = GetProcessName(UdpExTable(i).dwProcessId)
                    .SubItems(9) = ProcessPathByPID(UdpExTable(i).dwProcessId)
                End With
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    Else
        MsgBox "Can't get UDP table", vbExclamation
    End If
End Sub

Private Function GetProcessName(ByVal ProcessID As Long) As String
    Dim strName     As String * 1024
    Dim hProcess    As Long
    Dim cbNeeded    As Long
    Dim hMod        As Long
    Select Case ProcessID
        Case 0:    GetProcessName = "Proccess Inactive"
        Case 4:    GetProcessName = "System"
        Case Else: GetProcessName = "Unknown"
    End Select
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
    If hProcess Then
        If EnumProcessModules(hProcess, hMod, Len(hMod), cbNeeded) Then
            GetModuleBaseName hProcess, hMod, strName, Len(strName)
            GetProcessName = Left$(strName, lstrlen(strName))
        End If
        CloseHandle hProcess
    End If
End Function

Private Function GetIpString(ByVal Value As Long) As String
    Dim table(3) As Byte
    CopyMemory table(0), Value, 4
    GetIpString = table(0) & "." & table(1) & "." & table(2) & "." & table(3)
End Function

Private Function GetPortNumber(ByVal Value As Long) As Long
    GetPortNumber = (Value / 256) + (Value Mod 256) * 256
End Function

Private Function GetState(ByVal Value As Long) As String
    Select Case Value
        Case MIB_TCP_STATE_ESTAB: GetState = "ESTABLISH"
        Case MIB_TCP_STATE_CLOSED: GetState = "CLOSED"
        Case MIB_TCP_STATE_LISTEN: GetState = "LISTEN"
        Case MIB_TCP_STATE_CLOSING: GetState = "CLOSING"
        Case MIB_TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
        Case MIB_TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
        Case MIB_TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
        Case MIB_TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
        Case MIB_TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
        Case MIB_TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
        Case MIB_TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
    End Select
End Function

Public Function ProcessPathByPID(PID As Long) As String
    'to know process from its PID
    Dim cbNeeded As Long
    Dim Modules(1 To 200) As Long
    Dim Ret As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, PID)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded)
        If Ret <> 0 Then
            ModuleName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
            ProcessPathByPID = Left(ModuleName, Ret)
        End If
    End If
              
    Ret = CloseHandle(hProcess)
    If ProcessPathByPID = "" Then ProcessPathByPID = "SYSTEM"
End Function

Public Function Terminate(PID As Long) As Boolean
    'to terminate application by using its PID
    On Error Resume Next
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, PID)
    If GetExitCodeProcess(hProcess, lExitCode) = False Then
        Terminate = False
    Else
        If TerminateProcess(hProcess, lExitCode) = False Then
            Terminate = False
        Else
            Terminate = True
        End If
    End If
End Function

Public Function Update()
    'to update List1 in form4. List1 contain list of all process
    Dim Count, i As Integer
    Dim hSnapShot As Long, nProcess As Long
    Dim nPid As Long
    Dim uProcess As PROCESSENTRY32
    Dim DATA As String
    
    Form4.List1.Clear
    hSnapShot = CreateToolhelpSnapshot(2, 0)
    uProcess.dwSize = LenB(uProcess)
    nProcess = Process32First(hSnapShot, uProcess)
    
    Do While nProcess
        DATA = ProcessPathByPID(uProcess.th32ProcessID)
        If LCase(DATA) <> "system" Then
            If Atur(ProcessPathByPID(uProcess.th32ProcessID)) <> "" Then
                Form4.List1.AddItem Atur(ProcessPathByPID(uProcess.th32ProcessID))
                Form4.List1.ItemData(Form4.List1.NewIndex) = uProcess.th32ProcessID
            End If
        End If
        AppActive(Count) = UCase(uProcess.szExeFile)
        nProcess = Process32Next(hSnapShot, uProcess)
        For i = 1 To Val(NumberOfProg)
            If UCase(AppActive(Count)) = UCase(RemoveApp(i)) Then
                BlinkThread(i) = True
            End If
        Next i
        Count = Count + 1
    Loop
    counter = Count
    CloseHandle hSnapShot
End Function

Private Function Atur(awal As String) As String
    'to clear useless character in name of file
    Dim i As Integer
    Dim Ukuran As Integer
    If Mid(awal, 2, 1) = ":" Then
        Atur = awal
        Exit Function
    Else
        For i = 1 To Len(awal)
            If Mid(awal, i, 1) = ":" Then
                Atur = Mid(awal, i - 1, Len(awal) - i + 2)
                Exit Function
            End If
        Next
    End If
End Function



