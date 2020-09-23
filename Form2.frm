VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "yellowhead.com"
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton cmdGetIP 
      Caption         =   "Get IP from Host"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This code only works on the newer versions of Winsock V2 that comes with Windows
'Vista and Win7. It supports both IPv4 and IPv6. GetAddrInfo() is an API call that
'replaces the gethostbyname() call, and inet_ntop() replaces inet_ntoa(). inet_pton()
'replaces inet_aton() and inet_addr(), but it is not demonstrated here.

Private Const AF_INET As Long = 2
Private Const AF_INET6 As Long = 23
Private Const AF_UNSPEC As Long = 0
Private Const SOCK_STREAM As Long = 1
Private Const WSADESCRIPTION_LEN = 257
Private Const WSASYS_STATUS_LEN = 129

' Length of string fields for IPv4 and IPv6
Private Const INET_ADDRSTRLEN As Long = 16
Private Const INET6_ADDRSTRLEN As Long = 46

Private Type WSAData
   wVersion       As Integer
   wHighVersion   As Integer
   szDescription  As String * WSADESCRIPTION_LEN
   szSystemStatus As String * WSASYS_STATUS_LEN
   iMaxSockets    As Integer
   iMaxUdpDg      As Integer
   lpVendorInfo   As Long
End Type

Private Type AddrInfo
    ai_flags As Long
    ai_family As Long
    ai_socktype As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonname As Long 'strptr
    ai_addr As Long 'p sockaddr
    ai_next As Long 'p addrinfo
End Type
'
' Basic IPv4 addressing structures.
'
Private Type in_addr
   s_addr As Long
End Type
'
Private Type sockaddr_in
    sin_family          As Integer
    sin_port            As Integer
    sin_addr            As in_addr
    sin_zero(0 To 7)    As Byte
End Type
'
' Basic IPv6 addressing structures.
'
Private Type in6_addr
    s6_addr(0 To 15)      As Byte
End Type
'
Private Type sockaddr_in6
    sin6_family         As Integer
    sin6_port           As Integer
    sin6_flowinfo       As Long
    sin6_addr           As in6_addr
    sin6_scope_id       As Long
End Type
'
'To facilitate both IPv4 and IPv6 values, the sockaddr structure must be extended
'from 16 bytes to 28 bytes
Private Type sockaddr
    sa_family           As Integer '2 bytes
    sa_data(0 To 25)    As Byte    '26 bytes
End Type

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function GetAddrInfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal NodeName As String, ByVal ServName As String, ByVal lpHints As Long, lpResult As Long) As Long
Private Declare Function freeaddrinfo Lib "ws2_32.dll" (ByVal Res As Long) As Long
Private Declare Function inet_pton Lib "ws2_32.dll" (ByVal af As Long, ByVal pszAddrString As String, ByRef pAddrBuf As Any) As Long
Private Declare Function inet_ntop Lib "ws2_32.dll" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (pDestination As Any, ByVal lByteCount As Long)

Public Function IPtoString(ByVal lngIPAddress As Long, strLen As Long) As String
    Dim lpString   As Long
    Dim strBuffer  As String
    'Prepare a buffer, copy the IP into it, then trim and return.
    strBuffer = String$(strLen, 0)
    Call CopyMemory(ByVal strBuffer, ByVal lngIPAddress, Len(strBuffer))
    IPtoString = Mid$(strBuffer, 1, InStr(1, strBuffer, Chr$(0)))
    Debug.Print "IP " & IPtoString
End Function
Private Function PeekB(ByVal lpdwData As Long) As Byte
    CopyMemory PeekB, ByVal lpdwData, 1
End Function

Public Function GetIPFromHost(ByVal HostName As String) As String
'NOTE: When the source or the destination is a Visual Basic variable,
'it should be passed by reference, for example:
'    CopyMemory lngValue, intValue, 2
'When the source or the destination is a memory location
'it should be passed by value, for example:
'    CopyMemory ByVal address, lngValue, 4
'In this function, Hints is considered to be a variable, but Hints.ai_addr
'is considered to to be a memory location. If Hints is passed ByVal, the
'application works in the IDE, but will crash when compiled.
    'If contents of WSADATA are not needed, the following statement can be
    'used without a Type definition
    'Dim bWSAData(398) As Byte
    Dim udtWinsockData As WSAData
    Dim lpAi As Long
    Dim Hints As AddrInfo
    Dim resAi As AddrInfo
    'For reasons unknown, at least one definition for sockaddr or
    'sockaddr_in is required. Otherwise the GetAddrInfo routine fails.
    Dim Sa As sockaddr
    Dim Sa4 As sockaddr_in
    Dim Sa6 As sockaddr_in6
    Dim lRet As Long
    Dim aLen As Long
    'Used without Type definition
    'lRet = WSAStartup(2, bWSAData(0))
    lRet = WSAStartup(&H202, udtWinsockData)
    If lRet <> 0 Then MsgBox Err.LastDllError, vbInformation, "WSAStartup"
    lpAi = VarPtr(resAi)
    ZeroMemory Hints, Len(Hints)
    Hints.ai_family = AF_UNSPEC      'don't care IPv4 or IPv6
    Hints.ai_socktype = SOCK_STREAM  ' TCP stream sockets
    'Hints.ai_flags = AI_PASSIVE      ' fill in my IP for me
    lRet = GetAddrInfo(HostName, vbNullString, VarPtr(Hints), lpAi)
    If lRet <> 0 Then MsgBox Err.LastDllError, vbInformation, "GetAddrInfo"
    Hints.ai_next = lpAi
    While Hints.ai_next <> 0
        CopyMemory Hints, ByVal Hints.ai_next, LenB(Hints)
        If Hints.ai_family = AF_INET Then
            aLen = INET_ADDRSTRLEN
            CopyMemory Sa4, ByVal Hints.ai_addr, LenB(Sa4)
            ReDim bBuffer(0 To aLen - 1)
            lRet = inet_ntop(Hints.ai_family, Sa4.sin_addr, bBuffer(0), aLen)
        ElseIf Hints.ai_family = AF_INET6 Then
            aLen = INET6_ADDRSTRLEN
            CopyMemory Sa6, ByVal Hints.ai_addr, LenB(Sa6)
            ReDim bBuffer(0 To aLen - 1)
            lRet = inet_ntop(Hints.ai_family, Sa6.sin6_addr, bBuffer(0), aLen)
       End If
       If lRet Then GetIPFromHost = GetIPFromHost + IPtoString(lRet, aLen)
    Wend
    lRet = WSACleanup
    If lRet <> 0 Then MsgBox Err.LastDllError, vbInformation, "WSACleanup"
End Function




Private Sub cmdGetIP_Click()
    Dim M%, N%
    Dim IPList As String
    IPList = GetIPFromHost(Text1.Text)
    List1.Clear
    M% = 1
    N% = InStr(M%, IPList, Chr$(0))
    While N% <> 0
        List1.AddItem Mid$(IPList, M%, N% - M%)
        M% = N% + 1
        N% = InStr(M%, IPList, Chr$(0))
    Wend
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdGetIP_Click
End Sub


