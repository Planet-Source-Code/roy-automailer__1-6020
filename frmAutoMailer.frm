VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmAutoMailer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoMailer"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   ForeColor       =   &H00000000&
   Icon            =   "frmAutoMailer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
   End
   Begin MSMAPI.MAPIMessages mapMess 
      Left            =   5880
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession mapSess 
      Left            =   5280
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6840
      Top             =   4080
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
   End
End
Attribute VB_Name = "frmAutoMailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim bNewSession As Boolean

Private Sub cmdExit_Click()

    On Error GoTo ErrorHandler

    If cmdStart.Caption = "&Stop" Then
        cmdStart_Click
    End If

    DoEvents

    Unload Me

Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub cmdStart_Click()

    On Error GoTo ErrorHandler

    If cmdStart.Caption = "&Start" Then
        cmdStart.Caption = "&Stop"
        Log Now & vbTab & "---- Logging on to mail server..."
        LogOn
        Log Now & vbTab & "---- AutoMailer services started"
        Timer1.Enabled = True
    Else
        cmdStart.Caption = "&Start"
        Log Now & vbTab & "---- AutoMailer services stopped"
        Timer1.Enabled = False
    End If

Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub Form_Load()

    Dim sDBLocation As String
    Dim hSysMenu As Long

    On Error GoTo ErrorHandler

    ' disable the 'X':
    hSysMenu = GetSystemMenu(hwnd, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

    Log "AutoMailer version " & App.Major & "." & App.Minor & App.Revision
    Log "Press start to begin services"

    sDBLocation = App.Path & "\data.mdb"

    Set cn = New ADODB.Connection
    cn.Open "PROVIDER=MSDASQL;" & _
                "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                "DBQ=" & sDBLocation & ";" & _
                "UID=;PWD=;"

Exit Sub

ErrorHandler:
    Select Case Err.Number
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub Timer1_Timer()
    ' If the Flag field in the database is set to 1 the e-mail address for
    ' this record will be sent an e-mail.  The flag will then be changed to
    ' 0.  This is single mail mode.  The lines marked below may be removed to
    ' leave the flag a 1.  All addresses in the database marked with the 1 will
    ' sent an e-mail.

    ' PLEASE DO NOT USE THIS FOR EVIL.
    ' USE ONLY ON FRIENDS WHO DESERVE IT.

    ' more cool code at my website!!!
    ' www.geocities.com/royzda_one

    On Error GoTo ErrorHandler
    
    If LogOn = False Then
        Log Now & vbTab & "---- Error logging on to mail server!  Will try again."
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open "Select * from email where Flag = 1", cn, , , adCmdUnknown
    rs.MoveFirst

    If rs.RecordCount <> 0 Then
        While rs.EOF = False
            
            m_send_mail rs.Fields("EMail"), "You've been hit by the Spam-o-matic!"
            
            DoEvents

            Log Now & vbTab & "Mail sent to:  " & rs.Fields("EMail")
            
            ' *********comment the following two lines to begin the spam session *********
            rs.Fields("Flag").Value = 0
            rs.Update
            ' *****************************************************************************************
            
            rs.MoveNext
        Wend
    Else
        rs.Close
    End If

Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 3021
            Log Now & vbTab & "---- No requests"
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub m_send_mail(sRecipient As String, sMessage As String)

    On Error GoTo ErrorHandler
    
    Dim strMessage As String
    
    mapMess.Compose

    mapMess.RecipAddress = sRecipient

    mapMess.AddressResolveUI = True
    mapMess.ResolveName

    mapMess.MsgSubject = "I Love Spam®"
    mapMess.MsgNoteText = sMessage
    
    mapMess.Send False 'set this to true to view message and then send manually or cancel
    
Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select
End Sub

Private Function LogOn() As Boolean

    On Error GoTo ErrorHandler
    
    If mapSess.NewSession Then
        ' Session already established
        LogOn = True
        Exit Function
    End If
    

    With mapSess
        ' Set DownLoadMail to False to prevent immediate download.
        .DownLoadMail = False
        .LogonUI = True ' Use the underlying email system's logon UI.
        '-----------------------------------------------------------------------
        '.LogonUI = False
        '.UserName = "username"  ' Uncomment these lines, add your username and password
        '.Password = "password"   '    to eliminate the logon screen
        '-----------------------------------------------------------------------
        .SignOn
        ' If successful, return True
        LogOn = True
        ' Set NewSession to True and set
        ' variable flag to true
        .NewSession = True
        bNewSession = .NewSession
        mapMess.SessionID = .SessionID ' You must set this before continuing.
    End With
    
Exit Function

ErrorHandler:
    Select Case Err.Number
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Function

Private Sub Log(ByVal sText As String)

    ' this way it doesnt refresh the whole thing every time, no blinking
    With txtStatus
        .SelStart = Len(.Text)
        .SelText = sText & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
        .SelLength = 0
    End With

End Sub

Private Sub txtStatus_GotFocus()

    frmAutoMailer.SetFocus

End Sub
