VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOnline 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOnline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin XMSN.XMSNC XMSNC1 
      Height          =   540
      Left            =   2040
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   5835
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4260
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:32 PM"
         EndProperty
      EndProperty
      MousePointer    =   2
   End
   Begin VB.Frame fraContacts 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contacts"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton cmdRem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Remove A Contact"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Add A Contact"
         Top             =   4920
         Width           =   375
      End
      Begin MSComctlLib.TreeView tvContacts 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   7858
         _Version        =   393217
         Style           =   2
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnnuOut 
         Caption         =   "&Sign Out"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuStatus 
         Caption         =   "&Change Status"
         Begin VB.Menu mnuOnline 
            Caption         =   "O&nline"
         End
         Begin VB.Menu mnuAway 
            Caption         =   "&Away"
         End
         Begin VB.Menu mnuBusy 
            Caption         =   "&Busy"
         End
         Begin VB.Menu mnuBrb 
            Caption         =   "B&e Right Back"
         End
         Begin VB.Menu mnuIdle 
            Caption         =   "&Idle"
         End
         Begin VB.Menu mnuPhone 
            Caption         =   "&On The Phone"
         End
         Begin VB.Menu mnuLunch 
            Caption         =   "Out To &Lunch"
         End
         Begin VB.Menu mnuspace 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHide 
            Caption         =   "&Hidden"
         End
      End
      Begin VB.Menu mnuSession 
         Caption         =   "&Send Message"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuAddContact 
         Caption         =   "&Add Contact"
      End
      Begin VB.Menu mnuDelectContact 
         Caption         =   "&Delete Contact"
      End
      Begin VB.Menu mnuPrivacy 
         Caption         =   "&Privacy..."
      End
      Begin VB.Menu mnuspace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVRaw 
         Caption         =   "&View Raw Data Window"
      End
   End
End
Attribute VB_Name = "frmOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    Dim strAdd As String
    
    strAdd = InputBox("Enter the E-Mail Address Of The Person You Wish To Add To Your Contact List", "AddContact")
    
    If strAdd = "" Then
        MsgBox "Invalid Address"
        Exit Sub
    End If
    
    'XMSNC1.MSNAddUserToList AFL, strAdd
    XMSNC1.MSNAddUserToList AAL, strAdd

End Sub

Private Sub cmdRem_Click()

    If tvContacts.SelectedItem.Key = "" Or tvContacts.SelectedItem.Key = "Online" Or tvContacts.SelectedItem.Key = "Offline" Then
        MsgBox "Please Select A Contact To Remove"
        Exit Sub
    End If

    XMSNC1.MSNRemoveUserFromList AAL, tvContacts.SelectedItem.Key
    XMSNC1.MSNRemoveUserFromList AFL, tvContacts.SelectedItem.Key
    tvContacts.Nodes.Remove (tvContacts.SelectedItem.Key)

End Sub



Private Sub mnnuOut_Click()

    XMSNC1.MSNDisconnect

End Sub

Private Sub mnuAddContact_Click()

    cmdAdd_Click

End Sub

Private Sub mnuAway_Click()

    XMSNC1.MSNCurrentState = AWAY

End Sub

Private Sub mnuBrb_Click()

    XMSNC1.MSNCurrentState = BRB

End Sub

Private Sub mnuBusy_Click()

    XMSNC1.MSNCurrentState = BUSY

End Sub


Private Sub mnuDelectContact_Click()

    cmdRem_Click

End Sub

Private Sub mnuExit_Click()

    XMSNC1.MSNDisconnect
    End

End Sub

Private Sub mnuHide_Click()

    XMSNC1.MSNCurrentState = HIDN

End Sub

Private Sub mnuIdle_Click()

    XMSNC1.MSNCurrentState = IDLE

End Sub

Private Sub mnuLunch_Click()

    XMSNC1.MSNCurrentState = LUNCH

End Sub

Private Sub mnuOnline_Click()

    XMSNC1.MSNCurrentState = ONLINE

End Sub

Private Sub mnuPhone_Click()

    XMSNC1.MSNCurrentState = PHONE

End Sub


Private Sub mnuPrivacy_Click()

    blnPrivacyEdit = True
    frmContacts.lstAllow.Clear
    frmContacts.lstAllow1.Clear
    frmContacts.lstBlock.Clear
    frmContacts.lstBlock1.Clear
    XMSNC1.MSNRequestList LAL
    XMSNC1.MSNRequestList LBL
    frmContacts.Show vbModal

End Sub




Private Sub mnuSession_Click()

    If tvContacts.SelectedItem.Key <> "" And tvContacts.SelectedItem.Key <> "Online" And tvContacts.SelectedItem.Key <> "Offline" Then
        
        If tvContacts.SelectedItem.Parent <> "Offline" Then
            XMSNC1.MSNSendMessage
        End If
        
    End If

End Sub


Private Sub mnuVRaw_Click()

    frmRaw.Show

End Sub

Private Sub tvContacts_DblClick()

    mnuSession_Click

End Sub


Private Sub XMSNC1_MSNConnected(LogonName As String, FriendlyName As String)

    frmOnline.Caption = LogonName
    frmOnline.fraContacts.Caption = FriendlyName & ": Contacts"
    
End Sub

Private Sub XMSNC1_MSNContactStateChange(State As CSTATE, UserName As String, FriendlyName As String)

    Dim intChar As Integer, mNode As Node

    On Error GoTo NEWCONTACTFOUND
  
    Set mNode = tvContacts.Nodes(UserName)
    
    Select Case State
        Case Is = AWAY
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Away)")
        Case Is = BRB
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Be Right Back)")
        Case Is = BUSY
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Busy)")
        Case Is = HIDN
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Offline", tvwChild, UserName, FriendlyName)
        Case Is = IDLE
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Idle)")
        Case Is = LUNCH
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Out To Lunch)")
        Case Is = OFFLINE
            If InStr(mNode.Text, "(") > 0 Then
                intChar = InStr(mNode.Text, "(")
                FriendlyName = Mid(mNode.Text, 1, (intChar - 1))
            Else
                FriendlyName = mNode.Text
            End If
            
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Offline", tvwChild, UserName, FriendlyName)
        Case Is = ONLINE
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName)
        Case Is = PHONE
            tvContacts.Nodes.Remove (UserName)
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (On The Phone)")
    End Select
        
    Exit Sub
    
NEWCONTACTFOUND:
    
    XMSNC1.MSNSaveContact UserName, FriendlyName
    
    Select Case State
        Case Is = AWAY
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Away)")
        Case Is = BRB
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Be Right Back)")
        Case Is = BUSY
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Busy)")
        Case Is = HIDN
            Set mNode = tvContacts.Nodes.Add("Offline", tvwChild, UserName, FriendlyName)
        Case Is = IDLE
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Idle)")
        Case Is = LUNCH
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (Out To Lunch)")
        Case Is = OFFLINE
            FriendlyName = mNode.Text
            Set mNode = tvContacts.Nodes.Add("Offline", tvwChild, UserName, FriendlyName)
        Case Is = ONLINE
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName)
        Case Is = PHONE
            Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName & " (On The Phone)")
    End Select
        
End Sub


Private Sub XMSNC1_MSNDisconnect()

    Dim frmCheck As Form
    
    For Each frmCheck In Forms
        Unload frmCheck
        frmCheck.Hide
    Next
    
    tvContacts.Nodes.Clear
    frmSignOn.Show
    frmSignOn.Caption = "Logon"
    frmSignOn.MousePointer = 1

End Sub

Private Sub XMSNC1_MSNError(ErrNumber As Long)

    MsgBox "Error: " & ErrNumber & vbCrLf & XMSNC1.MSNGetErrorDesc(ErrNumber)
    
End Sub

Private Sub XMSNC1_MSNListChange(LstAction As LSTCHG, Lst As ALSET, UserName As String, FriendlyName As String)

    Dim intAsk As Integer, mNode As Node

    If LstAction = REMOVED Then
        XMSNC1.MSNRemoveSavedContact UserName
    
    ElseIf LstAction = ADDED Then
        
        If Lst = ARL_READONLY Then
            
            If XMSNC1.MSNReverseListSetting = MANUAL Then
                intAsk = MsgBox(UserName & " Has Just Added You To Their Contact List." & vbCrLf & "Do You Want To Add Them To Your Allow List?", vbYesNo, "Contact Notice")
                
                If intAsk = vbYes Then
                    XMSNC1.MSNAddUserToList AFL, UserName, FriendlyName
                    XMSNC1.MSNAddUserToList AAL, UserName, FriendlyName
                    XMSNC1.MSNSaveContact UserName, FriendlyName
                    Set mNode = tvContacts.Nodes.Add("Online", tvwChild, UserName, FriendlyName)
                End If
                
            End If
            
        End If
        
    End If

End Sub

Private Sub XMSNC1_MSNMail(Unread As Integer, InboxURL As String, FolderURL As String, PostURL As String)

    staMain.Panels(1).Text = "[Inbox] " & Unread & " Unread"

End Sub

Private Sub XMSNC1_MSNMessage(Joining As Boolean, SessionIndex As Integer, UserName As String, FriendlyName As String, Font As String, Message As String)

    Dim frmMes As Form, strCaption As String
    
    If Joining = True Then
        For Each frmMes In Forms
            If frmMes.Tag <> "" Then
                If Val(frmMes.Tag) <> SessionIndex Then
                    GoTo NEXTFRM
                Else
                    frmMes.txtInfo.Text = frmMes.txtInfo.Text & UserName & ":" & FriendlyName & ";"
                    frmMes.staIM.Panels(1).Text = "Message From: " & FriendlyName
                    strCaption = frmMes.Caption
                    strCaption = Replace(strCaption, "Chatting: ", "")
                    If Len(strCaption) = 0 Then
                        frmMes.Caption = "Chatting: " & FriendlyName
                    Else
                        frmMes.Caption = frmMes.Caption & ", " & FriendlyName
                    End If
                    Exit Sub
                End If
            End If
NEXTFRM:
        Next
        
        Set frmMes = New frmIM
        frmMes.Tag = SessionIndex
        frmMes.txtInfo.Text = frmMes.txtInfo.Text & UserName & ":" & FriendlyName & ";"
        frmMes.staIM.Panels(1).Text = "Message From: " & FriendlyName
        strCaption = frmMes.Caption
        strCaption = Replace(strCaption, "Chatting: ", "")
        If Len(strCaption) = 0 Then
            frmMes.Caption = "Chatting: " & FriendlyName
        Else
            frmMes.Caption = frmMes.Caption & ", " & FriendlyName
        End If
        
        frmMes.Show
        Exit Sub
    
    Else
    
        For Each frmMes In Forms
            If frmMes.Tag <> "" Then
                If Val(frmMes.Tag) <> SessionIndex Then
                    GoTo NEXTFRMEXIST
                Else
                    frmMes.rtfIn.SelColor = vbBlack
                    frmMes.rtfIn.SelStart = Len(frmMes.rtfIn.Text)
                    frmMes.rtfIn.SelFontSize = 10
                    frmMes.rtfIn.SelFontName = "Tahoma"
                    frmMes.rtfIn.SelText = FriendlyName & " says: " & vbCrLf & "   "
                    frmMes.rtfIn.SelStart = Len(frmMes.rtfIn.Text)
                    frmMes.rtfIn.SelFontSize = 8
                    frmMes.rtfIn.SelFontName = Font
                    frmMes.rtfIn.SelText = Message & vbCrLf
                    frmMes.txtTime.Text = Time & " - " & Date
                    frmMes.staIM.Panels(1).Text = "Last Message Received: " & frmMes.txtTime.Text
                    XMSNC1.MSNPlayResSound 3, "SOUND"
                    Exit Sub
                End If
            End If
NEXTFRMEXIST:
        Next

        Set frmMes = New frmIM
        frmMes.Tag = SessionIndex
        frmMes.rtfIn.SelColor = vbBlack
        frmMes.rtfIn.SelStart = Len(frmMes.rtfIn.Text)
        frmMes.rtfIn.SelFontSize = 10
        frmMes.rtfIn.SelFontName = "Tahoma"
        frmMes.rtfIn.SelText = FriendlyName & " says: " & vbCrLf & "   "
        frmMes.rtfIn.SelStart = Len(frmMes.rtfIn.Text)
        frmMes.rtfIn.SelFontSize = 8
        frmMes.rtfIn.SelFontName = Font
        frmMes.rtfIn.SelText = Message & vbCrLf
        frmMes.Caption = "Chatting: " & FriendlyName
        frmMes.txtInfo.Text = frmMes.txtInfo.Text & UserName & ":" & FriendlyName & ";"
        frmMes.staIM.Panels(1).Text = "Last Message Received: " & frmMes.txtTime.Text
        frmMes.Show
        XMSNC1.MSNPlayResSound 3, "SOUND"
        
    End If
    
End Sub

Private Sub XMSNC1_MSNMessageReady(SessionIndex As Integer)

    Dim frmMes As Form

    XMSNC1.MSNSendMessageEx SessionIndex, tvContacts.SelectedItem.Key

End Sub

Private Sub XMSNC1_MSNMessageTyping(UserName As String, FriendlyName As String)

    Dim frmMes As Form
    
    On Error Resume Next
    
    For Each frmMes In Forms
        
        If InStr(frmMes.txtInfo.Text, UserName) > 0 Then
            frmMes.staIM.Panels(1).Text = FriendlyName & " is typing a message"
        End If
    
    Next

End Sub

Private Sub XMSNC1_MSNOnline()
    
    frmSignOn.Hide
    frmSignOn.MousePointer = 1
    frmOnline.Show
    XMSNC1.MSNPlayResSound 2, "SOUND"

End Sub

Private Sub XMSNC1_MSNRawIncomingData(Data As String)

    frmRaw.txtRaw.Text = frmRaw.txtRaw.Text & vbCrLf & Data

End Sub




Private Sub XMSNC1_MSNSessionJoin(SessionIndex As Integer, UserName As String, FriendlyName As String)

    Dim frmMes As Form
    MsgBox "hi"
    For Each frmMes In Forms
        If frmMes.Tag <> "" Then
            If Val(frmMes.Tag) = SessionIndex Then
                frmMes.txtInfo.Text = frmMes.txtInfo.Text & UserName & ":" & FriendlyName & ";"
                frmMes.rtfIn.SelStart = Len(frmMes.rtfIn.Text)
                frmMes.rtfIn.SelFontSize = 10
                frmMes.rtfIn.SelFontName = "Tahoma"
                frmMes.rtfIn.SelColor = vbBlue
                frmMes.rtfIn.SelText = "---------------" & vbCrLf & FriendlyName & " has joined the conversation." & vbCrLf & "---------------" & vbCrLf
                If InStr(frmMes.Caption, UserName) = 0 Then
                    frmMes.Caption = frmMes.Caption & ", " & FriendlyName
                End If
                Exit Sub
            End If
        End If
    Next
    
    Set frmMes = New frmIM
    frmMes.Tag = SessionIndex
    frmMes.txtInfo.Text = frmMes.txtInfo.Text & UserName & ":" & FriendlyName & ";"
    frmMes.staIM.Panels(1).Text = "Message To: " & FriendlyName
    frmMes.Caption = "Chatting: " & FriendlyName
    frmMes.Show

End Sub

Private Sub XMSNC1_MSNSessionLeave(SessionIndex As Integer, UserName As String)

    Dim frmMes As Form, strNewList As String, arrSplit() As String, arrSplit2() As String
    Dim intLoop As Integer, intLoop1 As Integer, strTemp As String
    Dim arrCaptionSplit() As String, strCaption As String
    
    For Each frmMes In Forms
        If frmMes.Tag <> "" Then
            If CInt(frmMes.Tag) = SessionIndex Then
                arrSplit() = Split(frmMes.txtInfo.Text, ";")
                For intLoop = 0 To UBound(arrSplit)
                    If arrSplit(intLoop) <> "" Then
                        arrSplit2() = Split(arrSplit(intLoop), ":")
                        If arrSplit2(0) = UserName Then
                            strTemp = frmMes.Caption
                            strTemp = Replace(strTemp, "Chatting: ", "")
                            strTemp = Replace(strTemp, " ", "")
                            arrCaptionSplit() = Split(strTemp, ",")
                            If UBound(arrCaptionSplit) <> 0 Then
                                For intLoop1 = 0 To UBound(arrCaptionSplit)
                                    If arrCaptionSplit(intLoop1) <> arrSplit2(1) Then
                                        If strCaption = "" Then
                                            strCaption = arrCaptionSplit(intLoop1)
                                        Else
                                            strCaption = strCaption & ", " & arrCaptionSplit(intLoop1)
                                        End If
                                        frmMes.Caption = ""
                                        frmMes.Caption = "Chatting: " & strCaption
                                    End If
                                Next intLoop1
                                frmMes.rtfIn.SelStart = Len(frmMes.rtfIn.Text)
                                frmMes.rtfIn.SelFontSize = 10
                                frmMes.rtfIn.SelFontName = "Tahoma"
                                frmMes.rtfIn.SelColor = vbRed
                                frmMes.rtfIn.SelText = "---------------" & vbCrLf & arrSplit2(1) & " has left the conversation." & vbCrLf & "---------------" & vbCrLf
                            Else
                                frmMes.Caption = "Chatting: "
                                frmMes.txtInfo.Text = ""
                                frmMes.rtfIn.SelStart = Len(frmMes.rtfIn.Text)
                                frmMes.rtfIn.SelFontSize = 10
                                frmMes.rtfIn.SelFontName = "Tahoma"
                                frmMes.rtfIn.SelColor = vbRed
                                frmMes.rtfIn.SelText = "---------------" & vbCrLf & strTemp & " has left the conversation." & vbCrLf & "---------------" & vbCrLf
                            End If
                        Else
                            strNewList = strNewList & arrSplit2(0) & ":" & arrSplit2(1) & ";"
                            frmMes.txtInfo.Text = strNewList
                        End If
                    End If
                Next intLoop
            End If
        End If
    Next
                           
End Sub

Private Sub XMSNC1_MSNUserList(Lst As LRSET, UserName As String, FriendlyName As String)

    Dim intLoop As Integer, mNode As Node

    If blnPrivacyEdit = True Then
        
        Select Case Lst
            Case Is = LAL
                frmContacts.lstAllow.AddItem FriendlyName
                frmContacts.lstAllow1.AddItem UserName
            Case Is = LBL
                frmContacts.lstBlock.AddItem FriendlyName
                frmContacts.lstBlock1.AddItem UserName
            Case Is = LRL
                
                For intLoop = 0 To (frmContacts.lstAllow.ListCount - 1)
                    If frmContacts.lstAllow.List(intLoop) = FriendlyName Then
                        Exit Sub
                    End If
                Next
                
                For intLoop = 0 To (frmContacts.lstBlock.ListCount - 1)
                    If frmContacts.lstBlock.List(intLoop) = FriendlyName Then
                        Exit Sub
                    End If
                Next
                frmReverse.lstReverse.AddItem FriendlyName
                frmReverse.lstReverse1.AddItem UserName
        
        End Select
        
    Else
    
        On Error GoTo NEWCONTACTFOUND
  
        Set mNode = tvContacts.Nodes(UserName)
        
        Exit Sub
    
NEWCONTACTFOUND:
    
        XMSNC1.MSNSaveContact UserName, FriendlyName
        Set mNode = tvContacts.Nodes.Add("Offline", tvwChild, UserName, FriendlyName)
    
    End If

End Sub


