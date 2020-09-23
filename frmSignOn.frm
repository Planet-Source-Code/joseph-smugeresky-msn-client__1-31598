VERSION 5.00
Begin VB.Form frmSignOn 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "frmSignOn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSignon 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sign On"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtLogon 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   885
      Width           =   885
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Address:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   405
      Width           =   1335
   End
End
Attribute VB_Name = "frmSignOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub LoadContacts()

    Dim strContacts As String, arrContacts() As String, arrSplit() As String
    Dim intLoop As Integer, rNode As Node, mNode As Node
    
    On Error Resume Next
    Set mNode = frmOnline.tvContacts.Nodes.Add(, , "Online", "Online")
    mNode.Expanded = True
    Set mNode = frmOnline.tvContacts.Nodes.Add(, , "Offline", "Offline")
    mNode.Expanded = True
    
    strContacts = frmOnline.XMSNC1.MSNRetrieveContacts
    
    If strContacts = "" Then
        Exit Sub
    End If
    
    arrContacts = Split(strContacts, ";")
    
    For intLoop = 0 To UBound(arrContacts)
    
        If arrContacts(intLoop) <> "" Then
            arrSplit() = Split(arrContacts(intLoop), ",")
            Set mNode = frmOnline.tvContacts.Nodes.Add("Offline", tvwChild, arrSplit(0), arrSplit(1))
        End If
        
    Next intLoop

End Sub



Private Sub cmdSignon_Click()
    
    Me.MousePointer = 1

    If txtLogon = "" Then
        MsgBox "Invalid Email Address"
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        MsgBox "Invalid Password"
        Exit Sub
    End If

    frmOnline.XMSNC1.MSNLogonName = txtLogon
    frmOnline.XMSNC1.MSNPassword = txtPassword
    frmOnline.XMSNC1.MSNReverseListSetting = MANUAL
    LoadContacts
    frmOnline.XMSNC1.MSNConnect
    
    frmSignOn.Caption = "Signing On"
    frmSignOn.MousePointer = 11

End Sub


Private Sub Form_Load()
    
    txtPassword.PasswordChar = "*"

End Sub



Private Sub txtPassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdSignon_Click
    End If

End Sub
