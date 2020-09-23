VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl XMSNC 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   540
   ScaleHeight     =   525
   ScaleWidth      =   540
   Begin VB.PictureBox picMSN 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      Picture         =   "XMSNC.ctx":0000
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
   Begin MSWinsockLib.Winsock sckMain 
      Index           =   0
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "XMSNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'#############################################################################################
'#
'# VB MSN Client
'# Author: Joseph Smugeresky
'# Date January 24, 2002
'# E-Mail jsmugeresky@hotmail.com, jsmugeresky@aol.com, ilpre@aol.com
'# Version 1.0
'# Comments:
'#
'#
'# This client is a demonstartion of the MSN protocol in Visual Basic.
'# This is a very basic implementation of what the protocol can do.
'# I have not yet added support for file transfer, voice and video conferencing.
'# Also, the MD5 Algorithm was based off of some VB and C++ Code
'# I did not write all of the MD5 code but I did optomized most of what is here.
'# Please email any questions to any of the email addresses provided
'# VIEW THE README FOR MORE DETAILS!!!
'#
'# ###########################################################################################

'Connection Collection
Private ConnColl As New Collection

'Registry class var
Private reg As New clsReg

'Sound Class
Private sound As clsPlaySound

'MD5 Algorithm vars
Private lngTrack As Long
Private arrLongConversion(4) As Long
Private arrSplit64(63) As Byte

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

'Property vars
Private strLName As String
Private strFName As String
Private strPW As String
Private RLSetting As RLSET
Private intState As CSTATE
Private PLSetting As PLSET

'Event Enums

'Reverse List Enum
Public Enum RLSET
    AUTO = 0
    MANUAL = 1
End Enum

'Allow List Enum
Public Enum PLSET
    AL = 0
    BL = 1
End Enum

'State Enum
Public Enum CSTATE
    ONLINE = 0
    OFFLINE = 1
    HIDN = 2
    BUSY = 3
    IDLE = 4
    BRB = 5
    AWAY = 6
    PHONE = 7
    LUNCH = 8
End Enum

'List Retrieve Enum
Public Enum LRSET
    LFL = 0
    LAL = 1
    LBL = 2
    LRL = 3
End Enum

'List Add Enum
Public Enum ALSET
    AFL = 0
    AAL = 1
    ABL = 2
    ARL_READONLY = 3
End Enum

'List Change Enum
Public Enum LSTCHG
    ADDED = 0
    REMOVED = 1
End Enum

'For signon
Private blnSignOn As Boolean

'Events
Public Event MSNError(ErrNumber As Long)
Public Event MSNConnected(LogonName As String, FriendlyName As String)
Public Event MSNContactStateChange(State As CSTATE, UserName As String, FriendlyName As String)
Public Event MSNDisconnect()
Public Event MSNListChange(LstAction As LSTCHG, Lst As ALSET, UserName As String, FriendlyName As String)
Public Event MSNMail(Unread As Integer, InboxURL As String, FolderURL As String, PostURL As String)
Public Event MSNMessage(Joining As Boolean, SessionIndex As Integer, UserName As String, FriendlyName As String, Font As String, Message As String)
Public Event MSNMessageReady(SessionIndex As Integer)
Public Event MSNMessageTyping(UserName As String, FriendlyName As String)
Public Event MSNOnline()
Public Event MSNPermitList(Setting As PLSET)
Public Event MSNRawIncomingData(Data As String)
Public Event MSNReverseListSetting(Setting As RLSET)
Public Event MSNSessionJoin(SessionIndex As Integer, UserName As String, FriendlyName As String)
Public Event MSNSessionLeave(SessionIndex As Integer, UserName As String)
Public Event MSNUserList(Lst As LRSET, UserName As String, FriendlyName As String)
Private Function MD5Round(strRound As String, a As Long, b As Long, C As Long, d As Long, X As Long, S As Long, ac As Long) As Long

    Select Case strRound
    
        Case Is = "FF"
            a = MD5LongAdd4(a, (b And C) Or (Not (b) And d), X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
        
        Case Is = "GG"
            a = MD5LongAdd4(a, (b And d) Or (C And Not (d)), X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
            
        Case Is = "HH"
            a = MD5LongAdd4(a, b Xor C Xor d, X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
            
        Case Is = "II"
            a = MD5LongAdd4(a, C Xor (b Or Not (d)), X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
            
    End Select
    
End Function

Private Function MD5Rotate(lngValue As Long, lngBits As Long) As Long
    
    Dim lngSign As Long
    Dim lngI As Long
    
    lngBits = (lngBits Mod 32)
    
    If lngBits = 0 Then MD5Rotate = lngValue: Exit Function
    
    For lngI = 1 To lngBits
        lngSign = lngValue And &HC0000000
        lngValue = (lngValue And &H3FFFFFFF) * 2
        lngValue = lngValue Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next
    
    MD5Rotate = lngValue

End Function
Private Function TRID() As String

    Dim sngNum As Single, lngnum As Long
    Dim strResult As String
   
    sngNum = Rnd(2147483648#)
    strResult = CStr(sngNum)
    
    strResult = Replace(strResult, "0.", "")
    strResult = Replace(strResult, ".", "")
    strResult = Replace(strResult, "E-", "")
    
    TRID = strResult

End Function


Private Function MD564Split(lngLength As Long, bytBuffer() As Byte) As String

    Dim lngBytesTotal As Long, lngBytesToAdd As Long
    Dim intLoop As Integer, intLoop2 As Integer, lngTrace As Long
    Dim intInnerLoop As Integer, intLoop3 As Integer
    
    lngBytesTotal = lngTrack Mod 64
    lngBytesToAdd = 64 - lngBytesTotal
    lngTrack = (lngTrack + lngLength)
    
    If lngLength >= lngBytesToAdd Then
        For intLoop = 0 To lngBytesToAdd - 1
            arrSplit64(lngBytesTotal + intLoop) = bytBuffer(intLoop)
        Next intLoop
        
        MD5Conversion arrSplit64
        
        lngTrace = (lngLength) Mod 64

        For intLoop2 = lngBytesToAdd To lngLength - intLoop - lngTrace Step 64
            For intInnerLoop = 0 To 63
                arrSplit64(intInnerLoop) = bytBuffer(intLoop2 + intInnerLoop)
            Next intInnerLoop
            
            MD5Conversion arrSplit64
        
        Next intLoop2
        
        lngBytesTotal = 0
    Else
    
      intLoop2 = 0
    
    End If
    
    For intLoop3 = 0 To lngLength - intLoop2 - 1
        
        arrSplit64(lngBytesTotal + intLoop3) = bytBuffer(intLoop2 + intLoop3)
    
    Next intLoop3
     
End Function

Private Function MD5StringArray(strInput As String) As Byte()
    
    Dim intLoop As Integer
    Dim bytBuffer() As Byte
    ReDim bytBuffer(Len(strInput))
    
    For intLoop = 0 To Len(strInput) - 1
        bytBuffer(intLoop) = Asc(Mid(strInput, intLoop + 1, 1))
    Next intLoop
    
    MD5StringArray = bytBuffer
    
End Function
Private Sub MD5Conversion(bytBuffer() As Byte)

    Dim X(16) As Long, a As Long
    Dim b As Long, C As Long
    Dim d As Long
    
    a = arrLongConversion(1)
    b = arrLongConversion(2)
    C = arrLongConversion(3)
    d = arrLongConversion(4)
    
    MD5Decode 64, X, bytBuffer
    
    MD5Round "FF", a, b, C, d, X(0), S11, -680876936
    MD5Round "FF", d, a, b, C, X(1), S12, -389564586
    MD5Round "FF", C, d, a, b, X(2), S13, 606105819
    MD5Round "FF", b, C, d, a, X(3), S14, -1044525330
    MD5Round "FF", a, b, C, d, X(4), S11, -176418897
    MD5Round "FF", d, a, b, C, X(5), S12, 1200080426
    MD5Round "FF", C, d, a, b, X(6), S13, -1473231341
    MD5Round "FF", b, C, d, a, X(7), S14, -45705983
    MD5Round "FF", a, b, C, d, X(8), S11, 1770035416
    MD5Round "FF", d, a, b, C, X(9), S12, -1958414417
    MD5Round "FF", C, d, a, b, X(10), S13, -42063
    MD5Round "FF", b, C, d, a, X(11), S14, -1990404162
    MD5Round "FF", a, b, C, d, X(12), S11, 1804603682
    MD5Round "FF", d, a, b, C, X(13), S12, -40341101
    MD5Round "FF", C, d, a, b, X(14), S13, -1502002290
    MD5Round "FF", b, C, d, a, X(15), S14, 1236535329

    MD5Round "GG", a, b, C, d, X(1), S21, -165796510
    MD5Round "GG", d, a, b, C, X(6), S22, -1069501632
    MD5Round "GG", C, d, a, b, X(11), S23, 643717713
    MD5Round "GG", b, C, d, a, X(0), S24, -373897302
    MD5Round "GG", a, b, C, d, X(5), S21, -701558691
    MD5Round "GG", d, a, b, C, X(10), S22, 38016083
    MD5Round "GG", C, d, a, b, X(15), S23, -660478335
    MD5Round "GG", b, C, d, a, X(4), S24, -405537848
    MD5Round "GG", a, b, C, d, X(9), S21, 568446438
    MD5Round "GG", d, a, b, C, X(14), S22, -1019803690
    MD5Round "GG", C, d, a, b, X(3), S23, -187363961
    MD5Round "GG", b, C, d, a, X(8), S24, 1163531501
    MD5Round "GG", a, b, C, d, X(13), S21, -1444681467
    MD5Round "GG", d, a, b, C, X(2), S22, -51403784
    MD5Round "GG", C, d, a, b, X(7), S23, 1735328473
    MD5Round "GG", b, C, d, a, X(12), S24, -1926607734
  
    MD5Round "HH", a, b, C, d, X(5), S31, -378558
    MD5Round "HH", d, a, b, C, X(8), S32, -2022574463
    MD5Round "HH", C, d, a, b, X(11), S33, 1839030562
    MD5Round "HH", b, C, d, a, X(14), S34, -35309556
    MD5Round "HH", a, b, C, d, X(1), S31, -1530992060
    MD5Round "HH", d, a, b, C, X(4), S32, 1272893353
    MD5Round "HH", C, d, a, b, X(7), S33, -155497632
    MD5Round "HH", b, C, d, a, X(10), S34, -1094730640
    MD5Round "HH", a, b, C, d, X(13), S31, 681279174
    MD5Round "HH", d, a, b, C, X(0), S32, -358537222
    MD5Round "HH", C, d, a, b, X(3), S33, -722521979
    MD5Round "HH", b, C, d, a, X(6), S34, 76029189
    MD5Round "HH", a, b, C, d, X(9), S31, -640364487
    MD5Round "HH", d, a, b, C, X(12), S32, -421815835
    MD5Round "HH", C, d, a, b, X(15), S33, 530742520
    MD5Round "HH", b, C, d, a, X(2), S34, -995338651
 
    MD5Round "II", a, b, C, d, X(0), S41, -198630844
    MD5Round "II", d, a, b, C, X(7), S42, 1126891415
    MD5Round "II", C, d, a, b, X(14), S43, -1416354905
    MD5Round "II", b, C, d, a, X(5), S44, -57434055
    MD5Round "II", a, b, C, d, X(12), S41, 1700485571
    MD5Round "II", d, a, b, C, X(3), S42, -1894986606
    MD5Round "II", C, d, a, b, X(10), S43, -1051523
    MD5Round "II", b, C, d, a, X(1), S44, -2054922799
    MD5Round "II", a, b, C, d, X(8), S41, 1873313359
    MD5Round "II", d, a, b, C, X(15), S42, -30611744
    MD5Round "II", C, d, a, b, X(6), S43, -1560198380
    MD5Round "II", b, C, d, a, X(13), S44, 1309151649
    MD5Round "II", a, b, C, d, X(4), S41, -145523070
    MD5Round "II", d, a, b, C, X(11), S42, -1120210379
    MD5Round "II", C, d, a, b, X(2), S43, 718787259
    MD5Round "II", b, C, d, a, X(9), S44, -343485551
    
    arrLongConversion(1) = MD5LongAdd(arrLongConversion(1), a)
    arrLongConversion(2) = MD5LongAdd(arrLongConversion(2), b)
    arrLongConversion(3) = MD5LongAdd(arrLongConversion(3), C)
    arrLongConversion(4) = MD5LongAdd(arrLongConversion(4), d)
    
End Sub
Private Function MD5LongAdd(lngVal1 As Long, lngVal2 As Long) As Long
    
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    
    MD5LongAdd = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function
Private Function MD5LongAdd4(lngVal1 As Long, lngVal2 As Long, lngVal3 As Long, lngVal4 As Long) As Long
    
    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&) + (lngVal3 And &HFFFF&) + (lngVal4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + ((lngVal3 And &HFFFF0000) \ 65536) + ((lngVal4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    MD5LongAdd4 = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function

Private Sub MD5Decode(intLength As Integer, lngOutBuffer() As Long, bytInBuffer() As Byte)
    
    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double
    
    intDblIndex = 0
    
    For intByteIndex = 0 To intLength - 1 Step 4
        
        dblSum = bytInBuffer(intByteIndex) + bytInBuffer(intByteIndex + 1) * 256# + bytInBuffer(intByteIndex + 2) * 65536# + bytInBuffer(intByteIndex + 3) * 16777216#
        lngOutBuffer(intDblIndex) = MD5LongConversion(dblSum)
        intDblIndex = (intDblIndex + 1)
    
    Next intByteIndex

End Sub

Private Function MD5LongConversion(dblValue As Double) As Long
    
    If dblValue < 0 Or dblValue >= OFFSET_4 Then Error 6
        
    If dblValue <= MAXINT_4 Then
        MD5LongConversion = dblValue
    Else
        MD5LongConversion = dblValue - OFFSET_4
    End If
        
End Function

Private Sub MD5Finish()
    
    Dim dblBits As Double
    Dim arrPadding(72) As Byte
    Dim lngBytesBuffered As Long
    
    arrPadding(0) = &H80
    
    dblBits = lngTrack * 8
    
    lngBytesBuffered = lngTrack Mod 64
    
    If lngBytesBuffered <= 56 Then
        MD564Split (56 - lngBytesBuffered), arrPadding
    Else
        MD564Split (120 - lngTrack), arrPadding
    End If
    
    
    arrPadding(0) = MD5LongConversion(dblBits) And &HFF&
    arrPadding(1) = MD5LongConversion(dblBits) \ 256 And &HFF&
    arrPadding(2) = MD5LongConversion(dblBits) \ 65536 And &HFF&
    arrPadding(3) = MD5LongConversion(dblBits) \ 16777216 And &HFF&
    arrPadding(4) = 0
    arrPadding(5) = 0
    arrPadding(6) = 0
    arrPadding(7) = 0
    
    MD564Split 8, arrPadding
    
End Sub
Private Function MD5StringChange(lngnum As Long) As String
        
        Dim bytA As Byte
        Dim bytB As Byte
        Dim bytC As Byte
        Dim bytD As Byte
        
        bytA = lngnum And &HFF&
        If bytA < 16 Then
            MD5StringChange = "0" & Hex(bytA)
        Else
            MD5StringChange = Hex(bytA)
        End If
               
        bytB = (lngnum And &HFF00&) \ 256
        If bytB < 16 Then
            MD5StringChange = MD5StringChange & "0" & Hex(bytB)
        Else
            MD5StringChange = MD5StringChange & Hex(bytB)
        End If
        
        bytC = (lngnum And &HFF0000) \ 65536
        If bytC < 16 Then
            MD5StringChange = MD5StringChange & "0" & Hex(bytC)
        Else
            MD5StringChange = MD5StringChange & Hex(bytC)
        End If
       
        If lngnum < 0 Then
            bytD = ((lngnum And &H7F000000) \ 16777216) Or &H80&
        Else
            bytD = (lngnum And &HFF000000) \ 16777216
        End If
        
        If bytD < 16 Then
            MD5StringChange = MD5StringChange & "0" & Hex(bytD)
        Else
            MD5StringChange = MD5StringChange & Hex(bytD)
        End If

End Function

Private Function MD5Value() As String

    MD5Value = LCase(MD5StringChange(arrLongConversion(1)) & MD5StringChange(arrLongConversion(2)) & MD5StringChange(arrLongConversion(3)) & MD5StringChange(arrLongConversion(4)))

End Function

Private Function MSNEncryptPw(strPassword As String) As String

    Dim bytBuffer() As Byte
    
    bytBuffer = MD5StringArray(strPassword)
    
    MD5Start
    MD564Split Len(strPassword), bytBuffer
    MD5Finish
    
    MSNEncryptPw = MD5Value
    
End Function



Private Sub MD5Start()

    lngTrack = 0
    arrLongConversion(1) = MD5LongConversion(1732584193#)
    arrLongConversion(2) = MD5LongConversion(4023233417#)
    arrLongConversion(3) = MD5LongConversion(2562383102#)
    arrLongConversion(4) = MD5LongConversion(271733878#)
    
End Sub

Private Sub picMSN_Click()

End Sub

Private Sub sckMain_Close(Index As Integer)

    If Index <> 0 Then
        
        ConnColl.Remove Index
        Unload sckMain(Index)

    Else
        
        If blnSignOn = False Then
            RaiseEvent MSNDisconnect
        End If
    
    End If
    
End Sub

Private Sub sckMain_Connect(Index As Integer)
    
    Dim strSend As String
    
    If Index = 0 Then
    
        strSend = "VER " & TRID & " MSNP4" & Chr(13) & Chr(10)
        sckMain(0).SendData strSend
    
    Else
        
        Dim clsConnClass As clsConnect
    
        Set clsConnClass = ConnColl.Item(Index)
    
        If clsConnClass.CType = CMSG Then
            With clsConnClass
                strSend = "ANS " & TRID & " " & MSNLogonName & " " & .CAuthenticate & " " & .CSessionID & Chr(13) & Chr(10)
            End With
            
            sckMain(Index).SendData strSend
          
        ElseIf clsConnClass.CType = CSB Then
            With clsConnClass
                strSend = "USR " & TRID & " " & MSNLogonName & " " & clsConnClass.CAuthenticate & Chr(13) & Chr(10)
            End With
            
            clsConnClass.CType = CMSGOUT

            sckMain(Index).SendData strSend
    
        End If
    
    End If

End Sub

Private Sub sckMain_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim varData As Variant, strData As String
      
    sckMain(Index).GetData varData
    strData = StrConv(CStr(varData), vbUnicode)
    FilterServerMessage strData, Index
    
End Sub

Public Sub MSNConnect(Optional IPAddr As String = "64.4.12.123", Optional Port As Long = 1863)

    If IPAddr = "" Then
        RaiseEvent MSNError(100)
        Exit Sub
    End If
    
    If Port = 0 Or Port > 65535 Then
        RaiseEvent MSNError(101)
        Exit Sub
    End If
    
    blnSignOn = True
    
    If sckMain(0).State = 0 Then
        
        sckMain(0).Connect IPAddr, Port
    
    Else
        Do
            sckMain(0).Close
        Loop Until sckMain(0).State = 0
        
        sckMain(0).Connect IPAddr, Port
      
    End If
    
End Sub

Private Sub FilterServerMessage(strData As String, Index As Integer)

    'This is the main engine of the control.
    'Here is where we signon and process all messages.
    'Each connection is added to a collection of the clsConnect type
    'except for the initial connection of zero.

    Dim strSend As String, intChar As Integer, strTrID As String, arrNames() As String
    Dim strTemp As String, arrTemp() As String, intLoop As Integer
    Dim clsConnection As New clsConnect, arrVar() As Variant
        
    RaiseEvent MSNRawIncomingData(strData)

    If IsNumeric(Left(strData, 3)) Then
        RaiseEvent MSNError(CLng(Left(strData, 3)))
        Exit Sub
    End If
        
    Select Case Left(strData, 3)
    
        'Telling us to transfer to adifferent NS server.
        Case Is = "XFR"
            arrTemp = Split(strData, " ")
            
            If UBound(arrTemp) <> 5 Then
                strData = Replace(strData, " 0" & vbCrLf, "")
                intChar = InStrRev(strData, " ")
                intChar = intChar + 1
                strTemp = Mid(strData, intChar)
                strTemp = Replace(strTemp, vbCrLf, "")
                arrNames() = Split(strTemp, ":")
                DoEvents
                    sckMain(0).Close
                DoEvents
                sckMain(0).Connect arrNames(0), CLng(arrNames(1))
                
            ElseIf UBound(arrTemp) = 5 Then
                
                'This is a transfer to a switchboard server.
                'The switchboard server is how we message people.
                intLoop = (sckMain.UBound + 1)
                Load sckMain(intLoop)
            
                clsConnection.CIndex = intLoop
                clsConnection.CAddress = arrTemp(3)
                clsConnection.CAuthenticate = arrTemp(5)
                clsConnection.CType = CSB
                clsConnection.CAddress = MSNLogonName & ":" & MSNFriendlyName
                
                ConnColl.Add clsConnection, CStr(intLoop)
                
                arrNames() = Split(arrTemp(3), ":")
                sckMain(intLoop).Connect arrNames(0), CLng(arrNames(1))
                
            End If
        
        'Version of the protocol
        Case Is = "VER"
            strSend = "INF " & TRID & Chr(13) & Chr(10)
            sckMain(0).SendData strSend
    
        'Signon process, information on the user, and the security package used, which is MD5 for signon
        Case Is = "INF"
            strSend = "USR " & TRID & " MD5 I " & MSNLogonName & Chr(13) & Chr(10)
            sckMain(0).SendData strSend
        
        'Response from INF asking for the encryption of out password and the cookie sent with USR
        Case Is = "USR"
            If InStr(strData, "OK") = 0 Then
                intChar = InStr(4, strData, "S")
                intChar = intChar + 2
                strTrID = Mid(strData, intChar)
                strTrID = Replace(strTrID, vbCrLf, "")
                strSend = "USR " & TRID & " MD5 S " & MSNEncryptPw(strTrID & MSNPassword) & Chr(13) & Chr(10)
                sckMain(0).SendData strSend
            
            'We are connected
            ElseIf InStr(strData, "OK") > 0 Then
                If Index = 0 Then
                    intChar = InStr(4, strData, "OK")
                    intChar = intChar + 3
                    strTemp = Mid(strData, intChar)
                    strTemp = Replace(strTemp, vbCrLf, "")
                    arrNames() = Split(strTemp, " ")
                    arrNames(0) = Replace(arrNames(0), "%20", " ")
                    arrNames(1) = Replace(arrNames(1), "%20", " ")
                    strFName = arrNames(1)
                    RaiseEvent MSNConnected(arrNames(0), arrNames(1))
                    MSNLogonName = arrNames(0)
                    MSNFriendlyName = arrNames(1)
                    sckMain(0).SendData "SYN " & TRID & " " & GetLatestVersion & Chr(13) & Chr(10)
                Else
                    'A message is ready for further processing
                    RaiseEvent MSNMessageReady(Index)
                    
                End If
                
            End If
            
        'Synchronization of our contact list, this is where we get our Service Number from.
        'Each time the list changes, the Service number is increased by one.
        'If the client cache of the service number does not match that of the server
        'a list synchronization is sent.
        Case Is = "SYN"
        
            arrNames() = Split(strData, vbCrLf)
            intChar = InStrRev(arrNames(0), " ")
            strTemp = Mid(arrNames(0), intChar)
            If Val(strTemp) <> GetLatestVersion Then
                reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Ser#", strTemp
                ExtractList strData
            End If
            
            strSend = "CHG " & TRID & " NLN" & Chr(13) & Chr(10)
            sckMain(0).SendData strSend
            blnSignOn = False
            RaiseEvent MSNOnline
        
        Case Is = "GTC"
            ExtractList strData
        
        'Our contact list
        Case Is = "LST"
      
            ExtractList strData
            
        'Adding or have been added to a contact list
        Case Is = "ADD"
        
            strData = Replace(strData, vbCrLf, "")
        
            arrNames() = Split(strData, " ")
        
            If Val(arrNames(3)) <> GetLatestVersion Then
                reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Ser#", arrNames(3)
            End If
            
            Select Case arrNames(2)
            
                Case Is = "FL"
                    RaiseEvent MSNListChange(ADDED, AFL, arrNames(4), arrNames(5))
                
                Case Is = "AL"
                    RaiseEvent MSNListChange(ADDED, AAL, arrNames(4), arrNames(5))
                
                Case Is = "BL"
                    RaiseEvent MSNListChange(ADDED, ABL, arrNames(4), arrNames(5))
                
                Case Is = "RL"
                    If MSNReverseListSetting = AUTO Then
                        strSend = "ADD " & TRID & " FL " & arrNames(4) & " " & arrNames(5) & Chr(13) & Chr(10)
                        sckMain(0).SendData strSend
                        strSend = "ADD " & TRID & " AL " & arrNames(4) & " " & arrNames(5) & Chr(13) & Chr(10)
                        sckMain(0).SendData strSend
                        MSNSaveContact arrNames(4), arrNames(5)
                    End If
                    
                    RaiseEvent MSNListChange(ADDED, ARL_READONLY, arrNames(4), arrNames(5))
            
            End Select
            
        'Remove or have been removed from contact list
        Case Is = "REM"
        
            strData = Replace(strData, vbCrLf, "")
            
            arrNames() = Split(strData, " ")
        
            If Val(arrNames(3)) <> GetLatestVersion Then
                reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Ser#", arrNames(3)
            End If
            
            Select Case arrNames(2)
            
                Case Is = "FL"
                    RaiseEvent MSNListChange(REMOVED, AFL, arrNames(4), "Unknown")
                
                Case Is = "AL"
                    RaiseEvent MSNListChange(REMOVED, AAL, arrNames(4), "Unknown")
                
                Case Is = "BL"
                    RaiseEvent MSNListChange(REMOVED, ABL, arrNames(4), "Unknown")
                
                Case Is = "RL"
                    RaiseEvent MSNListChange(REMOVED, ARL_READONLY, arrNames(4), "Unknown")
            
            End Select

        'Online status change, also uses ILN to do this
        Case Is = "NLN"
            
            strData = Replace(strData, vbCrLf, "")
        
            arrNames() = Split(strData, " ")
            
            Select Case arrNames(1)
                
                Case Is = "NLN"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(ONLINE, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(ONLINE, arrNames(2), arrNames(3))
                    End If
                
                Case Is = "FLN"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(OFFLINE, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(OFFLINE, arrNames(2), arrNames(3))
                    End If
                
                Case Is = "BSY"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(BUSY, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(BUSY, arrNames(2), arrNames(3))
                    End If
                    
                Case Is = "IDL"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(IDLE, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(IDLE, arrNames(2), arrNames(3))
                    End If
                        
                Case Is = "BRB"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(BRB, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(BRB, arrNames(2), arrNames(3))
                    End If
                        
                Case Is = "AWY"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(AWAY, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(AWAY, arrNames(2), arrNames(3))
                    End If
                        
                Case Is = "PHN"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(PHONE, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(PHONE, arrNames(2), arrNames(3))
                    End If
                        
                Case Is = "LUN"
                    If UBound(arrNames()) <> 3 Then
                        RaiseEvent MSNContactStateChange(LUNCH, arrNames(2), "Unknown")
                    Else
                        RaiseEvent MSNContactStateChange(LUNCH, arrNames(2), arrNames(3))
                    End If
                        
            End Select
    
        Case Is = "ILN"
            
            arrTemp() = Split(strData, vbCrLf)
            
            For intLoop = 0 To (UBound(arrTemp) - 1)
            
                If arrTemp(intLoop) = "" Then: GoTo NXT
            
                arrNames() = Split(arrTemp(intLoop), " ")
                
                On Error Resume Next
                
                If Left(arrTemp(intLoop), 3) = "MSG" Then
                    Select Case Right(arrTemp(intLoop), 3)
                        Case Is = "221"
                            arrVar() = Array(arrTemp(intLoop + 4), arrTemp(intLoop + 5), arrTemp(intLoop + 6), arrTemp(intLoop + 7), arrTemp(intLoop + 8))
                            ExtractMailbox arrVar()
                    End Select
                End If
                
                Select Case arrNames(2)
                    
                    Case Is = "NLN"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(ONLINE, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(ONLINE, arrNames(3), arrNames(4))
                        End If
                        
                    Case Is = "FLN"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(OFFLINE, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(OFFLINE, arrNames(3), arrNames(4))
                        End If
                    
                    Case Is = "BSY"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(BUSY, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(BUSY, arrNames(3), arrNames(4))
                        End If
                    
                    Case Is = "IDL"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(IDLE, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(IDLE, arrNames(3), arrNames(4))
                        End If
                        
                    Case Is = "BRB"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(BRB, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(BRB, arrNames(3), arrNames(4))
                        End If
                        
                    Case Is = "AWY"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(AWAY, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(AWAY, arrNames(3), arrNames(4))
                        End If
                        
                    Case Is = "PHN"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(PHONE, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(PHONE, arrNames(3), arrNames(4))
                        End If
                        
                    Case Is = "LUN"
                        If UBound(arrNames()) <> 4 Then
                            RaiseEvent MSNContactStateChange(LUNCH, arrNames(3), "Unknown")
                        Else
                            RaiseEvent MSNContactStateChange(LUNCH, arrNames(3), arrNames(4))
                        End If
                              
                End Select
                
NXT:
             Next intLoop
                          
        'Contact offline
        Case Is = "FLN"
                    
            strData = Replace(strData, vbCrLf, "")
                
            arrNames() = Split(strData, " ")
            RaiseEvent MSNContactStateChange(OFFLINE, arrNames(1), "Unknown")
            
        'Message coming to us, here we take the params and connect to a switchboard server.
        'We then send our ANS
        Case Is = "RNG"
            
            arrNames() = Split(strData, " ")
            intLoop = (sckMain.UBound + 1)
            Load sckMain(intLoop)
            
            clsConnection.CIndex = intLoop
            
            clsConnection.CSessionID = CLng(arrNames(1))
            clsConnection.CAddress = arrNames(2)
            clsConnection.CAuthenticate = arrNames(4)
            clsConnection.CType = CMSG
            
            If UBound(arrNames) <> 6 Then
                clsConnection.CName = arrNames(5) & ":Unknown"
            Else
                clsConnection.CName = arrNames(5) & ":" & arrNames(6)
            End If
            
            clsConnection.CName = arrNames(5)
            ConnColl.Add clsConnection, CStr(intLoop)
    
            arrTemp = Split(arrNames(2), ":")
            sckMain(intLoop).Connect arrTemp(0), CLng(arrTemp(1))
        
        'Someone has joined our session
        Case Is = "JOI"
             
            strData = Replace(strData, vbCrLf, "")
            arrNames() = Split(strData, " ")
            If arrNames(1) <> MSNLogonName Then
                RaiseEvent MSNSessionJoin(Index, arrNames(1), arrNames(2))
            End If
            
        'Can be one of many things, check the content type
        Case Is = "MSG"
            
            arrTemp() = Split(strData, vbCrLf)
            If InStr(arrTemp(2), "x-msmsgsprofile") = 0 And InStr(arrTemp(2), "x-msmsgsinitialemailnotification") = 0 Then
                FilterMSNMessage Index, strData
            End If
            
        'We have joined a session
        Case Is = "IRO"
            
            strData = Replace(strData, vbCrLf, " ")
            arrNames() = Split(strData, " ")
            RaiseEvent MSNMessage(True, Index, arrNames(4), arrNames(5), "", "")

        'We have left or someone has left a session
        Case Is = "BYE"
        
            strData = Replace(strData, vbCrLf, "")
            arrNames() = Split(strData, " ")
            RaiseEvent MSNSessionLeave(Index, arrNames(1))

    End Select
    
End Sub

Public Property Get MSNLogonName() As String

    MSNLogonName = strLName

End Property

Public Property Let MSNLogonName(ByVal LogonName As String)

    strLName = LogonName

End Property

Public Property Get MSNPassword() As String

    MSNPassword = strPW

End Property

Public Property Let MSNPassword(ByVal Password As String)

    strPW = Password

End Property

Private Function GetLatestVersion() As Long

    'Gets the client cache of the Service Number

    Dim strCheck As String
    
    strCheck = reg.GetRegSetting(HKEY_CURRENT_USER, "XMSG\OCXINFO", "Info")
    
    If strCheck = "" Then
        reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO", "Info", "Do not change/modify/delete any of these values.  They are for internal use only by the control."
    End If
    
    strCheck = reg.GetRegSetting(HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Ser#")
    
    If strCheck = "" Then
        reg.CreateRegFolder HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName
        reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Info", "Do not change/modify/delete any of these values.  They are for internal use only by the control."
        reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Ser#", "0"
        GetLatestVersion = 0
        Exit Function
        
    Else
        
        GetLatestVersion = CLng(strCheck)
    
    End If

End Function



Public Property Get MSNFriendlyName() As String

    MSNFriendlyName = strFName

End Property


Public Property Let MSNFriendlyName(ByVal FriendlyName As String)

    strFName = FriendlyName

End Property

Private Sub ExtractList(strData As String)

    'Here we extract the current list settings, we only recieve some of these commands
    'if the Service numbers do not match.

    Dim arrData() As String, arrFields() As String
    Dim intLoop As Integer, strType As String
    
    arrData = Split(strData, vbCrLf)
    
    For intLoop = 0 To UBound(arrData)
        strType = Left(arrData(intLoop), 3)
        
        Select Case strType
        
            Case Is = "GTC"
            
                arrFields() = Split(arrData(intLoop), " ")
                If arrFields(3) = "A" Then
                    RaiseEvent MSNReverseListSetting(MANUAL)
                ElseIf arrFields(3) = "N" Then
                    RaiseEvent MSNReverseListSetting(AUTO)
                End If
                
            Case Is = "BLP"
                
                arrFields() = Split(arrData(intLoop), " ")
                If arrFields(3) = "AL" Then
                    RaiseEvent MSNPermitList(AL)
                ElseIf arrFields(3) = "BL" Then
                    RaiseEvent MSNPermitList(BL)
                End If
                
            Case Is = "LST"
                
                arrFields() = Split(arrData(intLoop), " ")
                
                Select Case arrFields(2)
                    
                    Case Is = "FL"
                        If arrFields(4) = "0" Then
                            GoTo NOLIST
                        End If
                        
                        If UBound(arrFields) <> 7 Then
                            RaiseEvent MSNUserList(LFL, arrFields(6), "Unknown")
                        Else
                            RaiseEvent MSNUserList(LFL, arrFields(6), arrFields(7))
                        End If
                        
                    Case Is = "AL"
                        If arrFields(4) = "0" Then
                            GoTo NOLIST
                        End If
                        
                        If UBound(arrFields) <> 7 Then
                            RaiseEvent MSNUserList(LAL, arrFields(6), "Unknown")
                        Else
                            RaiseEvent MSNUserList(LAL, arrFields(6), arrFields(7))
                        End If
                    
                    Case Is = "BL"
                        If arrFields(4) = "0" Then
                            GoTo NOLIST
                        End If
                        
                        If UBound(arrFields) <> 7 Then
                            RaiseEvent MSNUserList(LBL, arrFields(6), "Unknown")
                        Else
                            RaiseEvent MSNUserList(LBL, arrFields(6), arrFields(7))
                        End If
                        
                    Case Is = "RL"
                        If arrFields(4) = "0" Then
                            GoTo NOLIST
                        End If
                        
                        If UBound(arrFields) <> 7 Then
                            RaiseEvent MSNUserList(LRL, arrFields(6), "Unknown")
                        Else
                            RaiseEvent MSNUserList(LRL, arrFields(6), arrFields(7))
                        End If
            
NOLIST:
            
                End Select
            
            Case Else
        
        End Select

    Next intLoop


End Sub

Public Property Get MSNCurrentState() As CSTATE

    MSNCurrentState = intState

End Property

Public Property Let MSNCurrentState(ByVal State As CSTATE)

    Dim strSend As String

    intState = State
    
    If MSNCurrentConnectionState = 7 Then
        Select Case State
        
            Case Is = ONLINE
                strSend = "CHG " & TRID & " NLN" & Chr(13) & Chr(10)
            Case Is = AWAY
                strSend = "CHG " & TRID & " AWY" & Chr(13) & Chr(10)
            Case Is = BRB
                strSend = "CHG " & TRID & " BRB" & Chr(13) & Chr(10)
            Case Is = BUSY
                strSend = "CHG " & TRID & " BSY" & Chr(13) & Chr(10)
            Case Is = IDLE
                strSend = "CHG " & TRID & " IDL" & Chr(13) & Chr(10)
            Case Is = LUNCH
                strSend = "CHG " & TRID & " LUN" & Chr(13) & Chr(10)
            Case Is = PHONE
                strSend = "CHG " & TRID & " PHN" & Chr(13) & Chr(10)
            Case Is = HIDN
                strSend = "CHG " & TRID & " HDN" & Chr(13) & Chr(10)
            
        End Select
        
        sckMain(0).SendData strSend
        
    End If
    
End Property

Public Sub MSNAddUserToList(Lst As ALSET, UserName As String, Optional FriendlyName As String = "Unknown")

    Dim strSend As String

    If Lst = AFL Then
        
        If FriendlyName <> "" Then
            strSend = "ADD " & TRID & " FL " & UserName & " " & FriendlyName & Chr(13) & Chr(10)
        Else
            strSend = "ADD " & TRID & " FL " & UserName & " Unknown" & Chr(13) & Chr(10)
        End If
        
    ElseIf Lst = AAL Then
        
        If FriendlyName <> "" Then
            strSend = "ADD " & TRID & " AL " & UserName & " " & FriendlyName & Chr(13) & Chr(10)
        Else
            strSend = "ADD " & TRID & " AL " & UserName & " Unknown" & Chr(13) & Chr(10)
        End If
    
    ElseIf Lst = ABL Then
        
        If FriendlyName <> "" Then
            strSend = "ADD " & TRID & " BL " & UserName & " " & FriendlyName & Chr(13) & Chr(10)
        Else
            strSend = "ADD " & TRID & " BL " & UserName & " Unknown" & Chr(13) & Chr(10)
        End If
        
    ElseIf Lst = ARL_READONLY Then
        
        RaiseEvent MSNError(103)
        Exit Sub
    
    End If
    
    If MSNCurrentConnectionState = 7 Then
        sckMain(0).SendData strSend
    Else
        RaiseEvent MSNError(913)
        Exit Sub
    End If

End Sub


Public Sub MSNRemoveUserFromList(Lst As ALSET, UserName As String)

    Dim strSend As String

    If Lst = AFL Then
        
        strSend = "REM " & TRID & " FL " & UserName & Chr(13) & Chr(10)
        
    ElseIf Lst = AAL Then
        
        strSend = "REM " & TRID & " AL " & UserName & Chr(13) & Chr(10)

    ElseIf Lst = ABL Then
        
        strSend = "REM " & TRID & " BL " & UserName & Chr(13) & Chr(10)
        
    End If
    
    If MSNCurrentConnectionState = 7 Then
        sckMain(0).SendData strSend
    Else
        RaiseEvent MSNError(913)
        Exit Sub
    End If
    
End Sub

Public Sub MSNRequestList(Request As LRSET)

    Dim strSend As String
 
    If Request = LFL Then
        strSend = "LST " & TRID & " FL" & Chr(13) & Chr(10)
    
    ElseIf Request = LAL Then
        strSend = "LST " & TRID & " AL" & Chr(13) & Chr(10)
    
    ElseIf Request = LBL Then
        strSend = "LST " & TRID & " BL" & Chr(13) & Chr(10)
    
    ElseIf Request = LRL Then
        strSend = "LST " & TRID & " RL" & Chr(13) & Chr(10)
    
    End If
  
    If MSNCurrentConnectionState = 7 Then
        sckMain(0).SendData strSend
    Else
        RaiseEvent MSNError(913)
        Exit Sub
    End If

End Sub


Public Sub MSNSendMessage()

    Dim strSend As String
    
    strSend = "XFR " & TRID & " SB" & Chr(13) & Chr(10)

    If MSNCurrentConnectionState = 7 Then
        sckMain(0).SendData strSend
    Else
        RaiseEvent MSNError(913)
        Exit Sub
    End If

End Sub

Public Sub MSNSendMessageEx(SessionIndex As Integer, UserName As String)

    Dim clsConnClass As clsConnect, strSend As String
                    
    Set clsConnClass = ConnColl.Item(SessionIndex)
    clsConnClass.CName = UserName
    clsConnClass.CMsgTo = UserName
    strSend = "CAL " & TRID & " " & UserName & Chr(13) & Chr(10)
    sckMain(clsConnClass.CIndex).SendData strSend

End Sub


Private Sub ExtractMailbox(arrVar() As Variant)

    'This sub takes the number of items unread in your mailbox
    'It also gets the current path for your inbox

    Dim intChar As Integer
    
    intChar = InStr(arrVar(0), ":")
    arrVar(0) = Mid(arrVar(0), (intChar + 1))
    
    intChar = InStr(arrVar(2), ":")
    arrVar(2) = Mid(arrVar(2), (intChar + 1))
    
    intChar = InStr(arrVar(3), ":")
    arrVar(3) = Mid(arrVar(3), (intChar + 1))

    intChar = InStr(arrVar(4), ":")
    arrVar(4) = Mid(arrVar(4), (intChar + 1))
    
    RaiseEvent MSNMail(CInt(arrVar(0)), CStr(arrVar(2)), CStr(arrVar(3)), CStr(arrVar(4)))
    
End Sub

Private Sub FilterMSNMessage(intIndex As Integer, strData As String)

    'This sub filters through the mime messages sent
    'It first checks to see whether this is a typing message(TypingUser)
    'or if it is a regular text message.
    'File transfers also work in this manner.

    Dim arrNames() As String, arrTemp() As String, arrMsgSplit() As String
    Dim intChar As Integer
    
    arrTemp() = Split(strData, vbCrLf)
    arrNames() = Split(arrTemp(0), " ")
    
    If InStr(arrTemp(3), "TypingUser") > 0 Then
        RaiseEvent MSNMessageTyping(arrNames(1), arrNames(2))

    ElseIf InStr(arrTemp(3), "X-MMS-IM-Format") > 0 Then
        arrTemp() = Split(strData, vbCrLf, 5)
        arrMsgSplit() = Split(arrTemp(3), ";")
        intChar = InStr(arrMsgSplit(0), ":")
        arrMsgSplit(0) = Mid(arrMsgSplit(0), (intChar + 5))
        arrMsgSplit(0) = Replace(arrMsgSplit(0), "%20", " ")
        arrTemp(4) = Mid(arrTemp(4), 3)
        
        RaiseEvent MSNMessage(False, intIndex, arrNames(1), arrNames(2), arrMsgSplit(0), arrTemp(4))
    
    End If
    
End Sub


Public Function MSNGetErrorDesc(ErrNumber As Long) As String

    Select Case ErrNumber
    
        Case Is = 100
            MSNGetErrorDesc = "Invalid Server Address"
            
        Case Is = 101
            MSNGetErrorDesc = "Invalid Server Port"
            
        Case Is = 102
            MSNGetErrorDesc = "Invalid Client State"
    
        Case Is = 103
            MSNGetErrorDesc = "Cannot Add To Reverse List"
    
        Case Is = 200
            MSNGetErrorDesc = "Syntax Error"
            
        Case Is = 201
            MSNGetErrorDesc = "Invalid Parameter"
            
        Case Is = 205
            MSNGetErrorDesc = "Invalid User"
        
        Case Is = 206
            MSNGetErrorDesc = "FQDN Missing"
            
        Case Is = 207
            MSNGetErrorDesc = "Already Logged In"
            
        Case Is = 208
            MSNGetErrorDesc = "Invalid User Name"
            
        Case Is = 209
            MSNGetErrorDesc = "Invalid Friendly Name"
            
        Case Is = 210
            MSNGetErrorDesc = "List Full"

        Case Is = 215
            MSNGetErrorDesc = "Already There"
            
        Case Is = 216
            MSNGetErrorDesc = "Not On List"
    
        Case Is = 218
            MSNGetErrorDesc = "Already In The Mode"
            
        Case Is = 219
            MSNGetErrorDesc = "Already In Opposite List"
            
        Case Is = 280
            MSNGetErrorDesc = "Switchboard Failed"
        
        Case Is = 281
            MSNGetErrorDesc = "Notify Transfer Failed"
            
        Case Is = 300
            MSNGetErrorDesc = "Required Fields Missing"
            
        Case Is = 302
            MSNGetErrorDesc = "Not Logged In"
            
        Case Is = 500
            MSNGetErrorDesc = "Internal Server Error"
        
        Case Is = 501
            MSNGetErrorDesc = "Database Server Error"
            
        Case Is = 510
            MSNGetErrorDesc = "File Operation Error"
    
        Case Is = 520
            MSNGetErrorDesc = "Memory Allocation Error"
            
        Case Is = 600
            MSNGetErrorDesc = "Server Busy"
            
        Case Is = 601
            MSNGetErrorDesc = "Server Unavailable"
            
        Case Is = 602
            MSNGetErrorDesc = "Peer Notification Server Down"

        Case Is = 603
            MSNGetErrorDesc = "Database Connection Error"
            
        Case Is = 604
            MSNGetErrorDesc = "Server Going Down"

        Case Is = 707
            MSNGetErrorDesc = "Connection Creation Error"

        Case Is = 711
            MSNGetErrorDesc = "Write Blocking Error"
            
        Case Is = 712
            MSNGetErrorDesc = "Session Overload"
            
        Case Is = 713
            MSNGetErrorDesc = "User Too Active"
            
        Case Is = 714
            MSNGetErrorDesc = "Too Many Sessions"
            
        Case Is = 715
            MSNGetErrorDesc = "Unexpected Error"
            
        Case Is = 717
            MSNGetErrorDesc = "Bad Friend File"
            
        Case Is = 911
            MSNGetErrorDesc = "Authentication Failed"
        
        Case Is = 913
            MSNGetErrorDesc = "Not Allowed When Offline"
            
        Case Is = 920
            MSNGetErrorDesc = "Not Accepting New Users"
        
        Case Else
            MSNGetErrorDesc = "Unknown"
        
    End Select

End Function

Public Property Get MSNCurrentConnectionState() As Integer

    MSNCurrentConnectionState = sckMain(0).State

End Property


Public Sub MSNMessage(SessionIndex As Integer, Message As String, Font As String)

    Dim strSetup As String, strSend As String, clsConnectClass As clsConnect
    
    strSetup = "MIME-Version: 1.0" & Chr(13) & Chr(10)
    strSetup = strSetup & "Content-Type: text/plain" & Chr(13) & Chr(10)
    strSetup = strSetup & "X-MMS-IM-Format: FN=" & Font & "; EF=; CO=0; CS=0; PF=22" & Chr(13) & Chr(10)
    strSetup = strSetup & Chr(13) & Chr(10)
    strSetup = strSetup & Message
    
    strSend = "MSG " & TRID & " U " & Len(strSetup) & Chr(13) & Chr(10) & strSetup
    
    If sckMain(SessionIndex).State = 7 Then
        sckMain(SessionIndex).SendData strSend
    Else
        RaiseEvent MSNError(913)
        Exit Sub
    End If

End Sub




Public Sub MSNLeaveSession(SessionIndex As Integer)

    If sckMain(SessionIndex).State = 7 Then
        sckMain(SessionIndex).SendData "BYE " & MSNLogonName & Chr(13) & Chr(10)
        sckMain(SessionIndex).Close
    Else
        RaiseEvent MSNError(913)
        Exit Sub
    End If
    
End Sub



Public Sub MSNDisconnect()

    Dim intLoop As Integer
    
    On Error Resume Next
    
    blnSignOn = False
    sckMain(0).Close
    
    RaiseEvent MSNDisconnect

End Sub


Public Sub MSNInviteUserToSession(SessionIndex As Integer, UserName As String)

    If sckMain(SessionIndex).State = 7 Then
        sckMain(SessionIndex).SendData "CAL " & TRID & " " & UserName & Chr(13) & Chr(10)
    Else
        RaiseEvent MSNError(913)
        Exit Sub
    End If
End Sub


Public Sub MSNSaveContact(UserName As String, FriendlyName As String)

    Dim strContacts As String

    If MSNLogonName = "" Then
        RaiseEvent MSNError(205)
        Exit Sub
    End If

    If UserName = "" Then
        Exit Sub
    ElseIf FriendlyName = "" Then
        Exit Sub
    End If
    
    strContacts = reg.GetRegSetting(HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Contacts")
    If InStr(strContacts, UserName) > 0 Then
        Exit Sub
    ElseIf strContacts = "" Then
        reg.CreateRegFolder HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName
    End If
    
    strContacts = strContacts & UserName & "," & FriendlyName & ";"
    
    reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Contacts", strContacts

End Sub

Public Function MSNRetrieveContacts() As String

    If MSNLogonName = "" Then
        RaiseEvent MSNError(205)
        Exit Function
    End If

    MSNRetrieveContacts = reg.GetRegSetting(HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Contacts")

End Function

Public Sub MSNRemoveSavedContact(UserName As String)

    Dim strContacts As String, arrContacts() As String, intLoop As Integer
    Dim arrSplitContacts() As String, strNewContacts As String
    
    strContacts = MSNRetrieveContacts
    
    If strContacts = "" Or UserName = "" Then
        RaiseEvent MSNError(205)
        Exit Sub
    End If

    arrContacts() = Split(strContacts, ";")
    
    For intLoop = 0 To UBound(arrContacts)
        If arrContacts(intLoop) <> "" Then
            arrSplitContacts() = Split(arrContacts(intLoop), ",")
            If arrSplitContacts(0) <> LCase(UserName) Then
                strNewContacts = strNewContacts & arrContacts(intLoop) & ";"
            End If
        End If
    Next intLoop
    
    MSNRefreshSavedContacts strNewContacts

End Sub

Private Sub MSNRefreshSavedContacts(Contacts As String)

    reg.SetRegValue HKEY_CURRENT_USER, "XMSG\OCXINFO\" & MSNLogonName, "Contacts", Contacts

End Sub

Public Property Get MSNReverseListSetting() As RLSET

    MSNReverseListSetting = RLSetting

End Property

Public Property Let MSNReverseListSetting(ByVal Setting As RLSET)

    RLSetting = Setting

    If MSNCurrentConnectionState = 7 Then
        Select Case Setting
        
            Case Is = AUTO
            
                sckMain(0).SendData "GTC " & TRID & " N" & Chr(13) & Chr(10)
                
            Case Is = MANUAL
            
                sckMain(0).SendData "GTC " & TRID & " A" & Chr(13) & Chr(10)
                
        End Select
        
    End If

End Property

Public Property Get MSNGeneralPrivacy() As PLSET

    MSNGeneralPrivacy = PLSetting

End Property

Public Property Let MSNGeneralPrivacy(ByVal Setting As PLSET)

    PLSetting = Setting

    If MSNCurrentConnectionState = 7 Then
        Select Case Setting
        
            Case Is = AL
                
                sckMain(0).SendData "BLP " & TRID & " AL" & Chr(13) & Chr(10)

            Case Is = BL
            
                sckMain(0).SendData "BLP " & TRID & " BL" & Chr(13) & Chr(10)
            
        End Select
    
    End If
    
End Property

Private Sub UserControl_Initialize()

    Height = 540
    Width = 540

End Sub

Private Sub UserControl_Resize()

    Height = 540
    Width = 540

End Sub



Public Sub MSNPlayResSound(ResID As Variant, ResName As String)

    Set sound = New clsPlaySound

    sound.PlaySound ResID, ResName

End Sub
