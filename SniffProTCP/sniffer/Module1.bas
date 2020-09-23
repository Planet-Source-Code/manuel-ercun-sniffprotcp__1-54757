Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function WSAstartup Lib "wsock32.dll" Alias "WSAStartup" (ByVal wVersionRequired As Integer, ByRef lpWSAData As WSAdata) As Long
Private Declare Function WsACleanup Lib "wsock32.dll" Alias "WSACleanup" () As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function inet_ntoa Lib "wsock32.dll" (ByVal addr As Long) As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal name As String, ByVal namelen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal name As String) As Long
Private Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Private Declare Function recv Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Private Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function bind Lib "wsock32.dll" (ByVal s As Integer, addr As sockaddr, ByVal namelen As Integer) As Integer
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer


Private Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 255
    szSystemStatus As String * 128
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Private Type ipheader
     lenver As Byte
      tos As Byte
      len As Integer
     ident As Integer
     flags As Integer
     ttl As Byte
     proto As Byte
    checksum As Integer
    sourceIP As Long
    destIP As Long
End Type

'typedef struct tcp_hdr
'{
'USHORT th_sport
'USHORT th_dport
'unsigned int th_seq
'unsigned int th_ack
'unsigned char th_lenres
'unsigned char th_flag
'USHORT th_win
'USHORT th_sum
'USHORT th_urp
'}


Private Type tcp_hdr
    th_sport As Integer
    th_dport As Integer
    th_seq As Long
    th_ack As Long
    th_lenres As Byte
    th_flag As Byte
    th_win As Integer
    th_sum As Integer
    th_urp As Integer
End Type

'typedef struct udp_hdr
'{
'USHORT th_sport
'USHORT th_dport
'USHORT th_len
'USHORT th_sum
'}


Private Type udp_hdr
    th_sport As Integer
    th_dport As Integer
    th_len As Integer
    th_sum As Integer
End Type

'typedef struct _ihdr {
 ' BYTE i_type;
 ' BYTE i_code;
 ' USHORT i_cksum;
 ' USHORT i_id;
 ' USHORT i_seq;
 ' ULONG timestamp;
'};

Private Type icmp_hdr
    th_type As Byte
    th_code As Byte
    th_sum As Integer
    th_id As Integer
    th_seq As Integer
    th_time As Long
End Type


Private Const PF_INET = 2
Private Const SOCK_RAW = 3
Private Const AF_INET = 2
Public Const FD_READ = &H1
Private Const SIO_RCVALL = &H98000001
Private Const EM_REPLACESEL = &HC2

Dim host As HOSTENT
Public s As Long
Dim sock As sockaddr

Dim Header As ipheader
Dim tcpHead As tcp_hdr
Dim udpHead As udp_hdr
Dim icmpHead As icmp_hdr


Public tamaño() As Long, str As String
Public i As Long, z As Long
Dim protocol As String
Dim buffer() As Byte
Dim res As Long
Public salir As Boolean


Public Sub Wstartup()
Dim data As WSAdata
Call WSAstartup(&H202, data)

End Sub

Public Sub WCleanup(s As Long)
Call WsACleanup
closesocket s
End Sub



Public Function ip(ByRef address As String) As String

Dim pip As Long
Dim uip As Long
Dim s As Long
Dim ss As String
Dim cul As Long
CopyMemory host, ByVal gethostbyname(address), Len(host)
CopyMemory pip, ByVal host.h_addr_list, 4
CopyMemory uip, ByVal pip, 4
s = inet_ntoa(uip)
ss = Space(lstrlen(s))
cul = lstrcpy(ss, s)
ip = ss
End Function

Public Function hostname() As String
Dim r As Long
Dim s As String
Dim host As String

Wstartup
host = String(255, 0)
r = gethostname(host, 255)

If r = 0 Then hostname = Left(host, InStr(1, host, vbNullChar) - 1)



End Function


Public Sub Connecting(ByRef ip As String, pic As PictureBox)
Dim res As Long, buf As Long, bufb As Long
buf = 1

Wstartup

s = socket(AF_INET, SOCK_RAW, 0)
If s < 1 Then WCleanup s: Exit Sub
sock.sin_family = AF_INET
sock.sin_addr = inet_addr(ip)
res = bind(s, sock, Len(sock))
If res <> 0 Then WCleanup s: Exit Sub
res = WSAIoctl(s, SIO_RCVALL, buf, Len(buf), 0, 0, bufb, ByVal 0, ByVal 0)
If res <> 0 Then WCleanup s: Exit Sub
res = WSAAsyncSelect(s, pic.hWnd, &H202, ByVal FD_READ)
If res <> 0 Then WCleanup s: Exit Sub

End Sub

Public Sub Recibir(s As Long, ByVal wsafd As Long)



If wsafd = FD_READ Then
ReDim buffer(2000)
Do
res = recv(s, buffer(0), 2000, 0&)

If res > 0 Then

ReDim Preserve tamaño(z)
'Call SendMessage(Form1.Text1.hwnd, EM_REPLACESEL, 0, buffer(0))
str = buffer()
tamaño(z) = res

CopyMemory Header, buffer(0), Len(Header)

 
Debug.Print Header.proto

If Header.proto = 1 Then protocol = "ICMP": proticmp inversaip(Hex(Header.destIP)), inversaip(Hex(Header.sourceIP))
If Header.proto = 6 Then protocol = "TCP": protcp inversaip(Hex(Header.destIP)), inversaip(Hex(Header.sourceIP))
If Header.proto = 17 Then protocol = "UDP": proudp inversaip(Hex(Header.destIP)), inversaip(Hex(Header.sourceIP))



End If

Loop Until res <> 2000

End If

End Sub

Private Function inversaip(ByRef lng As String) As String
Dim sos As String
For i = 1 To Len(lng) Step 2
sos = Asc(Chr("&h" & Mid(lng, i, 2))) & "." & sos
Next i
inversaip = Mid(sos, 1, Len(sos) - 1)

End Function


Private Function proticmp(saa As String, soc As String) As String
Dim b
Dim t
Set t = Form1.ListView1.ListItems.Add(, , soc)
t.SubItems(2) = saa
t.SubItems(4) = protocol
t.SubItems(5) = Time
CopyMemory icmpHead, buffer(0 + 20), Len(icmpHead)
With Form1.TreeView1.Nodes
Set b = .Add(, , , soc & "->" & saa, 5)
.Add b, tvwChild, , "Protocol:" & protocol, 6
.Add b, tvwChild, , "code:" & icmpHead.th_code, 6
.Add b, tvwChild, , "i_id:" & icmpHead.th_id, 6
.Add b, tvwChild, , "timestamp:" & icmpHead.th_time, 6
.Add b, tvwChild, , "checksum:" & icmpHead.th_sum, 6
.Add b, tvwChild, , "Lenght:" & res, 6
End With
End Function



Private Sub protcp(saa As String, soc As String)
Dim b, t
CopyMemory tcpHead, buffer(0 + 20), Len(tcpHead)

With Form1.TreeView1.Nodes
Set b = .Add(, , , soc & "->" & saa, 10)
.Add b, tvwChild, , "Version:" & Hex(Header.lenver), 6
.Add b, tvwChild, , "tos:" & Header.tos, 6
.Add b, tvwChild, , "tot_len:" & res, 6
.Add b, tvwChild, , "id:" & Header.ident, 6
.Add b, tvwChild, , "frag_off:" & Header.flags, 6
.Add b, tvwChild, , "ttl:" & Header.ttl, 6
.Add b, tvwChild, , "Protocol:" & protocol, 6
.Add b, tvwChild, , "checksum:" & Header.checksum, 6
.Add b, tvwChild, , "TCP HEADER", 5
.Add b, tvwChild, , "Sport:" & ntohs(tcpHead.th_sport), 7
.Add b, tvwChild, , "Dport:" & ntohs(tcpHead.th_dport), 7
.Add b, tvwChild, , "Seq:" & tcpHead.th_seq, 7
.Add b, tvwChild, , "ack:" & tcpHead.th_ack, 7
.Add b, tvwChild, , "Off:" & tcpHead.th_lenres, 7
.Add b, tvwChild, , "Flag:" & tcpHead.th_flag, 7
.Add b, tvwChild, , "Windows:" & tcpHead.th_win, 7
.Add b, tvwChild, , "Checkum:" & tcpHead.th_sum, 7
.Add b, tvwChild, , "Urp:" & tcpHead.th_urp, 7




End With

Set t = Form1.ListView1.ListItems.Add(, , soc)
t.SubItems(1) = ntohs(tcpHead.th_sport)
t.SubItems(2) = saa
t.SubItems(3) = ntohs(tcpHead.th_dport)
t.SubItems(4) = protocol
t.SubItems(5) = Time
End Sub

Private Sub proudp(saa As String, soc As String)
Dim b, t
CopyMemory udpHead, buffer(0 + 20), Len(udpHead)

With Form1.TreeView1.Nodes
Set b = .Add(, , , soc & "->" & saa, 9)
.Add b, tvwChild, , "Version:" & Hex(Header.lenver), 6
.Add b, tvwChild, , "tos:" & Header.tos, 6
.Add b, tvwChild, , "tot_len:" & res, 6
.Add b, tvwChild, , "id:" & Header.ident, 6
.Add b, tvwChild, , "frag_off:" & Header.flags, 6
.Add b, tvwChild, , "ttl:" & Header.ttl, 6
.Add b, tvwChild, , "Protocol:" & protocol, 6
.Add b, tvwChild, , "checksum:" & Header.checksum, 6
.Add b, tvwChild, , "UDP HEADER", 5
.Add b, tvwChild, , "Sport:" & ntohs(udpHead.th_sport), 8
.Add b, tvwChild, , "Dport:" & ntohs(udpHead.th_dport), 8
.Add b, tvwChild, , "Ulen:" & res, 8
.Add b, tvwChild, , "Checksum:" & udpHead.th_sum, 8
End With
Set t = Form1.ListView1.ListItems.Add(, , soc)
t.SubItems(1) = ntohs(udpHead.th_sport)
t.SubItems(2) = saa
t.SubItems(3) = ntohs(udpHead.th_dport)
t.SubItems(4) = protocol
t.SubItems(5) = Time
End Sub

