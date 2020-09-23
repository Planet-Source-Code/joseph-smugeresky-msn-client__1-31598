VERSION 5.00
Begin VB.Form frmContacts 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Privacy Controls"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAsk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ask Me When Other Users Add Me To Their Contact List"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   5295
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame fraPPO 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Other Prefs"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   7095
      Begin VB.OptionButton opBlock 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Do Not Allow All Other Users To See My Online Status, Send Messages, etc."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   6855
      End
      Begin VB.OptionButton opAllow 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Allow All Other Users To See My Online Status, Send Messages, etc."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame fraPP 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Privacy Prefernces"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.CommandButton cmdRev 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reverse List"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ListBox lstBlock1 
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox lstAllow1 
         Height          =   255
         Left            =   2760
         TabIndex        =   10
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdBlock 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Block >>"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdAllow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "<< Allow"
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.ListBox lstBlock 
         Height          =   2400
         Left            =   4440
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.ListBox lstAllow 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAsk_Click()

    If chkAsk.Value = 1 Then
        frmOnline.XMSNC1.MSNReverseListSetting = MANUAL
    Else
        frmOnline.XMSNC1.MSNReverseListSetting = AUTO
    End If

End Sub

Private Sub cmdAllow_Click()

    If lstBlock.List(lstBlock.ListIndex) = "" Then
        Exit Sub
    End If

    frmOnline.XMSNC1.MSNRemoveUserFromList ABL, lstBlock1.List(lstBlock.ListIndex)
    frmOnline.XMSNC1.MSNSaveContact lstBlock1.List(lstBlock.ListIndex), lstBlock.List(lstBlock.ListIndex)
    frmOnline.XMSNC1.MSNAddUserToList AAL, lstBlock1.List(lstBlock.ListIndex), lstBlock.List(lstBlock.ListIndex)
    frmOnline.XMSNC1.MSNAddUserToList AAL, lstBlock1.List(lstBlock.ListIndex), lstBlock.List(lstBlock.ListIndex)

    lstAllow.AddItem lstBlock.List(lstBlock.ListIndex)
    lstAllow1.AddItem lstBlock1.List(lstBlock.ListIndex)
    
    lstBlock1.RemoveItem lstBlock.ListIndex
    lstBlock.RemoveItem lstBlock.ListIndex

End Sub

Private Sub cmdBlock_Click()

    If lstAllow.List(lstAllow.ListIndex) = "" Then
        Exit Sub
    End If

    frmOnline.XMSNC1.MSNRemoveUserFromList AAL, lstAllow1.List(lstAllow.ListIndex)
    frmOnline.XMSNC1.MSNRemoveUserFromList AFL, lstAllow1.List(lstAllow.ListIndex)
    frmOnline.XMSNC1.MSNRemoveSavedContact lstAllow1.List(lstAllow.ListIndex)
    frmOnline.XMSNC1.MSNAddUserToList ABL, lstAllow1.List(lstAllow.ListIndex), lstAllow.List(lstAllow.ListIndex)

    lstBlock.AddItem lstAllow.List(lstAllow.ListIndex)
    lstBlock1.AddItem lstAllow1.List(lstAllow.ListIndex)
    
    lstAllow1.RemoveItem lstAllow.ListIndex
    lstAllow.RemoveItem lstAllow.ListIndex

End Sub

Private Sub cmdOk_Click()

    Unload Me
    Me.Hide

End Sub

Private Sub cmdRev_Click()

    frmOnline.XMSNC1.MSNRequestList LRL
    frmReverse.lstReverse.Clear
    frmReverse.lstReverse1.Clear
    frmReverse.Show vbModal

End Sub

Private Sub Form_Load()
        
    If frmOnline.XMSNC1.MSNGeneralPrivacy = AL Then
        opAllow.Value = True
    Else
        opBlock.Value = True
    End If
    
    If frmOnline.XMSNC1.MSNReverseListSetting = MANUAL Then
        chkAsk.Value = 1
    End If

End Sub

Private Sub opAllow_Click()

    frmOnline.XMSNC1.MSNGeneralPrivacy = AL

End Sub

Private Sub opBlock_Click()

    frmOnline.XMSNC1.MSNGeneralPrivacy = BL

End Sub


