VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" _
   (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Integer, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal Timeout As Long) As Long
Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type
Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Integer
    Reserved        As Integer
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String
End Type
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type
Private Declare Function WSAStartup Lib "WSOCK32.DLL" _
   (ByVal wVersionRequired As Long, _
    lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" _
   (ByVal szHost As String, _
    ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
   (ByVal szHost As String) As Long
Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
Private Declare Sub RtlMoveMemory Lib "kernel32" _
   (hpvDest As Any, _
    ByVal hpvSource As Long, _
    ByVal cbCopy As Long)
Private Declare Function inet_addr Lib "WSOCK32.DLL" _
        (ByVal cp As String) As Long


Public Function Ping(ByVal szAddress As String) As Boolean
   Dim hPort        As Long
   Dim dwAddress    As Long
   Dim sDataToSend  As String
   Dim iOpt         As Long
    Dim ECHO_Int As ICMP_ECHO_REPLY

   dwAddress = inet_addr(szAddress)

   hPort = IcmpCreateFile()
   Ping = IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), 0, ECHO_Int, Len(ECHO_Int), 200)
End Function

Public Function GetIP(ByVal hostname As String) As String
Dim hostent_addr As Long
Dim host As HOSTENT
Dim hostip_addr As Long
Dim temp_ip_address() As Byte
Dim i As Integer
Dim ip_address As String

    hostent_addr = gethostbyname(hostname)
    If hostent_addr = 0 Then
        MsgBox "Tidak Ketemu"
        Exit Function
    End If

    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4
    ReDim temp_ip_address(1 To host.hLength)
    RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

    For i = 1 To host.hLength
        ip_address = ip_address & "." & temp_ip_address(i)
    Next
    ip_address = Mid$(ip_address, 2)
    GetIP = ip_address
End Function

Private Sub Class_Initialize()
   Dim WSAD As WSADATA
    If WSAStartup(&H101, WSAD) <> 0 Then
        MsgBox "Komponen WinSock32.dll Gagal"
        Exit Sub
   End If
End Sub

Private Sub Class_Terminate()
    WSACleanup
End Sub
