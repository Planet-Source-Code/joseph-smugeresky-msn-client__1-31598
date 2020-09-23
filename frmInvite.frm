VERSION 5.00
Begin VB.Form frmInvite 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invite To Conversation"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox lstInvite1 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton cmdInvite 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Invite"
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox lstInvite 
      Height          =   2985
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

    Unload Me
    Me.Hide

End Sub

Private Sub cmdInvite_Click()

    If lstInvite.List(lstInvite.ListIndex) = "" Then
        Unload Me
        Me.Hide
        Exit Sub
    End If
    
    frmOnline.XMSNC1.MSNInviteUserToSession CInt(Me.Tag), lstInvite1.List(lstInvite.ListIndex)
    
    Unload Me
    Me.Hide

End Sub
