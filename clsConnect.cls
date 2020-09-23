VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'User reference for identifying collection, always the users logon(email address)
Private strCollectionName As String

'Sending Message To
Private strMsgTo As String

'Index of the control
Private intIndex As Integer

'SessionID
Private lngSessionID As Long

'NS Server Address
Private strAddress As String

'Cookie
Private strAuthenticationInfo As String

'Type of class
Public Enum CITYPE
    CMSG = 0
    CSB = 1
    CMSGOUT = 2
End Enum

'Type of class
Private intType As CITYPE

Public Property Get CName() As String

    CName = strCollectionName

End Property

Public Property Let CName(ByVal strName As String)

    strCollectionName = strName

End Property

Public Property Get CIndex() As Integer

    CIndex = intIndex

End Property

Public Property Let CIndex(ByVal intCIndex As Integer)

    intIndex = intCIndex

End Property

Public Property Get CSessionID() As Long

    CSessionID = lngSessionID

End Property

Public Property Let CSessionID(ByVal lngID As Long)

    lngSessionID = lngID

End Property

Public Property Get CAddress() As String

    CAddress = strAddress

End Property

Public Property Let CAddress(ByVal strAddr As String)

    strAddress = strAddr

End Property

Public Property Get CAuthenticate() As String

    CAuthenticate = strAuthenticationInfo

End Property

Public Property Let CAuthenticate(ByVal strAuthInfo As String)

    strAuthenticationInfo = strAuthInfo

End Property

Public Property Get CType() As CITYPE

    CType = intType

End Property

Public Property Let CType(ByVal intIType As CITYPE)

    intType = intIType

End Property

Public Property Get CMsgTo() As String

    CMsgTo = strMsgTo

End Property

Public Property Let CMsgTo(ByVal strTo As String)

    strMsgTo = strTo

End Property



