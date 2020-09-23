VERSION 5.00
Begin VB.Form frmReverse 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recent Reverse List Entries"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReverse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdBlock 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Block User"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdAllow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow User"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.ListBox lstReverse1 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox lstReverse 
      Height          =   4155
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmReverse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAllow_Click()

    If lstReverse.List(lstReverse.ListIndex) = "" Then
        Exit Sub
    End If
    
    frmOnline.XMSNC1.MSNAddUserToList AFL, lstReverse1.List(lstReverse.ListIndex), lstReverse.List(lstReverse.ListIndex)
    frmOnline.XMSNC1.MSNAddUserToList AAL, lstReverse1.List(lstReverse.ListIndex), lstReverse.List(lstReverse.ListIndex)
    frmOnline.XMSNC1.MSNSaveContact lstReverse1.List(lstReverse.ListIndex), lstReverse.List(lstReverse.ListIndex)

    frmContacts.lstAllow.AddItem lstReverse.List(lstReverse.ListIndex)
    frmContacts.lstAllow1.AddItem lstReverse1.List(lstReverse.ListIndex)
    lstReverse.RemoveItem lstReverse.ListIndex

End Sub

Private Sub cmdBlock_Click()

    If lstReverse.List(lstReverse.ListIndex) = "" Then
        Exit Sub
    End If
    
    frmOnline.XMSNC1.MSNAddUserToList ABL, lstReverse1.List(lstReverse.ListIndex), lstReverse.List(lstReverse.ListIndex)

    frmContacts.lstBlock.AddItem lstReverse.List(lstReverse.ListIndex)
    frmContacts.lstBlock1.AddItem lstReverse1.List(lstReverse.ListIndex)
    lstReverse.RemoveItem lstReverse.ListIndex

End Sub

Private Sub cmdOk_Click()

    frmReverse.Hide
    
End Sub

