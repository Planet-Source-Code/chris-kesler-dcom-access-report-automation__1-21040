VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MailSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mSession As MAPISession
Private mMessages As MAPIMessages

Private mSessionIsOpen As Boolean
Private mLoggedOnUser As String

Public Property Get LoggedOnUser() As String
   LoggedOnUser = mLoggedOnUser
End Property

Public Property Get SessionIsOpen() As Boolean
   SessionIsOpen = mSessionIsOpen
End Property

Public Sub SignOn(ProfileName As String, _
                  Password As String)
   
   On Error Resume Next

   With mSession
      .UserName = ProfileName
      .Password = Password
      .LogonUI = False
      .NewSession = True
      
      .SignOn
   End With

   Select Case Err.Number
      Case 0
         mSessionIsOpen = True
      Case 32050        'Already Logged In
         mSessionIsOpen = True
      Case Else
         mSessionIsOpen = False
   End Select

   If mSessionIsOpen Then
      mLoggedOnUser = ProfileName
   End If
End Sub

Public Sub SignOff()
   On Error Resume Next
   mSession.SignOff
End Sub

'Can only send messages to one person at a time.
'Also does not do attachments.
Public Sub Send(Recipient As String, _
                Subject As String, _
                Message As String)
                   On Error GoTo TrapSendErrors
   
   With mMessages
       .SessionID = mSession.SessionID
       .Compose
       .RecipDisplayName = Recipient
       .MsgSubject = Subject
       .MsgNoteText = Message
       .ResolveName
       .Send
   End With
   
   Exit Sub

TrapSendErrors:

   Select Case Err.Number
      Case 32014  'Invalid Recipient
         'MsgBox "Invalid Recipient - " & Recip
         Resume Next
      Case Else
         App.LogEvent "An error occurred while attempting to send an email: [" & Err.Number & "] " & Err.Description
         'Do Nothing
   End Select
End Sub

Private Sub Class_Initialize()
   mLoggedOnUser = ""
   mSessionIsOpen = False
   
   Load frmMail

   Set mSession = frmMail.MAPISession1
   Set mMessages = frmMail.MAPIMessages1
End Sub

Private Sub Class_Terminate()
   If mSessionIsOpen Then
      SignOff
   End If
   
   Set mSession = Nothing
   Set mMessages = Nothing

   Unload frmMail
End Sub
