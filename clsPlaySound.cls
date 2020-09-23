VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlaySound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Call for playing sounds from a resource file

'Declare
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long

'Constants
Private Const SND_ASYNC = &H1
Private Const SND_MEMORY = &H4
Private Const SND_NODEFAULT = &H2


Public Sub PlaySound(ResID As Variant, ResName As String)

    Dim strConvert As String
    
    strConvert = StrConv(LoadResData(ResID, ResName), vbUnicode)

    sndPlaySound strConvert, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY

End Sub
