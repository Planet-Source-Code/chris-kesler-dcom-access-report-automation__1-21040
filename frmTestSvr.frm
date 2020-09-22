VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "SSDW3B32.OCX"
Begin VB.Form frmTestSvr 
   Caption         =   "Mail Reports Remotely"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   Icon            =   "frmTestSvr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFileNames 
      Height          =   840
      ItemData        =   "frmTestSvr.frx":0ECA
      Left            =   1410
      List            =   "frmTestSvr.frx":0ECC
      TabIndex        =   0
      Top             =   30
      Width           =   3270
   End
   Begin VB.CommandButton cmdClearForm 
      Caption         =   "Clear Form"
      Height          =   390
      Left            =   4740
      TabIndex        =   7
      Top             =   30
      Width           =   1470
   End
   Begin SSDataWidgets_B.SSDBGrid lstCriteriaVals 
      Height          =   1800
      Left            =   1410
      TabIndex        =   2
      Top             =   1341
      Width           =   4155
      _Version        =   196616
      DataMode        =   2
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      FieldSeparator  =   ","
      Col.Count       =   3
      MultiLine       =   0   'False
      RowSelectionStyle=   1
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      ExtraHeight     =   106
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "Parameter"
      Columns(0).Name =   "CriteriaName"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   873
      Columns(1).Caption=   "="
      Columns(1).Name =   "CriteriaCon"
      Columns(1).Alignment=   2
      Columns(1).CaptionAlignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Value"
      Columns(2).Name =   "CriteriaValue"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   7329
      _ExtentY        =   3175
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtMessage 
      Height          =   960
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3990
      Width           =   4815
   End
   Begin VB.TextBox txtSubject 
      Height          =   330
      Left            =   1410
      TabIndex        =   4
      Top             =   3597
      Width           =   4800
   End
   Begin VB.TextBox txtTO 
      Height          =   330
      Left            =   1410
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3204
      Width           =   3270
   End
   Begin SSDataWidgets_B.SSDBCombo ssdbRptNames 
      Height          =   345
      Left            =   1410
      TabIndex        =   1
      Top             =   933
      Width           =   3270
      DataFieldList   =   "Column 0"
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns(0).Width=   7144
      Columns(0).Caption=   "ReportNames"
      Columns(0).Name =   "ReportNames"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   5768
      _ExtentY        =   609
      _StockProps     =   93
      BackColor       =   -2147483643
      DataFieldToDisplay=   "Column 0"
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "Send Mail"
      Enabled         =   0   'False
      Height          =   390
      Left            =   4740
      TabIndex        =   6
      Top             =   3180
      Width           =   1470
   End
   Begin VB.Label lblInstruction 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To Exclude a Parameter from the Report just add a # to the Value Side of that parameter."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   30
      TabIndex        =   13
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Choose Database:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   12
      Top             =   75
      Width           =   990
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   555
      TabIndex        =   11
      Top             =   3960
      Width           =   810
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   705
      TabIndex        =   10
      Top             =   3615
      Width           =   660
   End
   Begin VB.Label lblTo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1110
      TabIndex        =   9
      Top             =   3285
      Width           =   255
   End
   Begin VB.Label lblReportName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Report Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   1005
      Width           =   1110
   End
End
Attribute VB_Name = "frmTestSvr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'DCOM Access Mail Client
'Created on: 01/15/2001
'Created by: Chris Kesler
'Created for: Arch Wireless/Pagenet (non Exclusive)
'************************************************************

Dim rptSvr As Object
Dim dbs As DAO.Database
Dim rptNames As ADOR.Recordset
Dim rptParms As ADOR.Recordset
Dim rptSent As Boolean
Dim bmkMemory As Long

'*************************************************************
'Procedure:    Private Method cmdClearForm_Click
'Created on:   02/08/01
'Module:       frmTestSvr
'Module File:  D:\AccessServerEx\frmTestSvr.frm
'Project:      TestSvr
'Project File: D:\AccessServerEx\TestSvr.vbp
'Parameters:
'
'*************************************************************

Private Sub cmdClearForm_Click()
Dim dbArray()
Dim ctl As Control
    
    Set rptSvr = Nothing
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "TextBox"
                ctl.Text = ""
            Case "SSDBCombo"
                ctl.Text = ""
                ctl.Enabled = True
                If ctl.Name <> "ssCriteria" Then
                    ctl.RemoveAll
                End If
            Case "SSDBGrid"
                ctl.RemoveAll
            Case "CommandButton"
                If ctl.Caption = "Get Reports" Then
                    ctl.Enabled = True
                ElseIf ctl.Caption = "Clear Form" Then
                    ctl.Enabled = True
                Else
                    ctl.Enabled = False
                End If
            Case "ListBox"
                ctl.Enabled = True
                ctl.Clear
        End Select
    Next
    
    Set rptSvr = CreateObject("ReportServer.SendReport")
    ReDim Preserve dbArray(UBound(rptSvr.DBVals))
    dbArray = rptSvr.DBVals
    For i = 0 To UBound(dbArray)
        lstFileNames.AddItem dbArray(i)
    Next
    
End Sub


'*************************************************************
'Procedure:    Private Method cmdSendMail_Click
'Created on:   02/08/01
'Module:       frmTestSvr
'Module File:  D:\AccessServerEx\frmTestSvr.frm
'Project:      TestSvr
'Project File: D:\AccessServerEx\TestSvr.vbp
'Parameters:
'
'*************************************************************

Private Sub cmdSendMail_Click()
    Dim RecArray() As String, i
    Dim NameArray() As String
    Dim RealError As Long
    
On Error GoTo CommandErr
    
    ReDim RecArray(Int(lstCriteriaVals.Rows) - 1)
   lstCriteriaVals.MoveFirst
   For i = 0 To lstCriteriaVals.Rows - 1
        RecArray(i) = lstCriteriaVals.Columns(2).Value
        lstCriteriaVals.MoveNext
    Next
    ReDim NameArray(Int(lstCriteriaVals.Rows) - 1)
    i = 0
    lstCriteriaVals.MoveFirst
    For i = 0 To lstCriteriaVals.Rows - 1
        NameArray(i) = lstCriteriaVals.Columns(0).Value
        lstCriteriaVals.MoveNext
    Next
    
    MousePointer = vbHourglass
    If rptSvr.BuildAndSendRpt(txtTO.Text, txtSubject.Text, ssdbRptNames.Text, txtMessage.Text, RecArray(), NameArray()) = True Then
        MsgBox "The Report was successfully built and sent!", vbInformation, "Message Sent"
    End If
    MousePointer = vbDefault

CommandErr:
    RealError = Err.Number
    
    If RealError > 0 And RealError < 65535 Then
        Select Case RealError
            Case UserNotFoundErr
                MsgBox "User Not Found!!!  Contact your administrator.", vbOKOnly, "User Name Changed or not Present"
            Case DataBaseErr
                MsgBox "Database is busy, another user has the database opened exclusively." & vbCrLf & vbCrLf & _
                            "Please try again momentarily." & vbCrLf & vbTab & "Err: " & Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
            Case SQLErr
                MsgBox "There was an error building the SQL Statement on the Server." & vbCrLf & vbCrLf & _
                            "Please try again or contact your administrator if the problem continues." & vbCrLf & vbTab & "Err: " & _
                            Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
            Case UnknownErr
                MsgBox "An unknown error was detected.", vbOKOnly, "Customization Component Error"
            Case Else
                MsgBox "Unhandled or Unexpected Error", vbOKOnly, "Error Unhandled"
        End Select
    End If
    MousePointer = vbDefault
End Sub

'*************************************************************
'Procedure:    Private Method Form_Load
'Created on:   02/08/01
'Module:       frmTestSvr
'Module File:  D:\AccessServerEx\frmTestSvr.frm
'Project:      TestSvr
'Project File: D:\AccessServerEx\TestSvr.vbp
'Parameters:
'
'*************************************************************

Private Sub Form_Load()
Dim dbArray()

On Error GoTo CommandErr
    
    bmkMemory = -1
    Set rptSvr = CreateObject("ReportServer.SendReport")
    ReDim Preserve dbArray(UBound(rptSvr.DBVals))
    dbArray = rptSvr.DBVals
    For i = 0 To UBound(dbArray)
        lstFileNames.AddItem dbArray(i)
    Next
    
CommandErr:
    RealError = Err.Number
    
    If RealError > 0 And RealError < 65535 Then
        Select Case RealError
            Case UserNotFoundErr
                MsgBox "User Not Found!!!  Contact your administrator.", vbOKOnly, "User Name Changed or not Present"
            Case DataBaseErr
                MsgBox "Database is busy, another user has the database opened exclusively." & vbCrLf & vbCrLf & _
                            "Please try again momentarily." & vbCrLf & vbTab & "Err: " & Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
            Case SQLErr
                MsgBox "There was an error building the SQL Statement on the Server." & vbCrLf & vbCrLf & _
                            "Please try again or contact your administrator if the problem continues." & vbCrLf & vbTab & "Err: " & _
                            Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
            Case UnknownErr
                MsgBox "An unknown error was detected.", vbOKOnly, "Customization Component Error"
            Case Else
                MsgBox "There was an Error Creating the ActiveX Server, please try again or contact your administrator", vbOKOnly, "Problem Starting Server Component"
        End Select
    End If
    
End Sub

Private Sub GetReports(Filename As String)

    ssdbRptNames.RemoveAll
    MousePointer = vbHourglass
    Set rptNames = New ADOR.Recordset
    Set rptNames = rptSvr.GetRptNames(Filename)
    Do Until rptNames.EOF
        With ssdbRptNames
            .AddItem rptNames("rptName").Value
        End With
        rptNames.MoveNext
    Loop
    MousePointer = vbDefault
    Set rptNames = Nothing

End Sub
 
 Private Sub GetParameters(RptName As String)
     
    lstCriteriaVals.RemoveAll
    MousePointer = vbHourglass
    Set rptParms = New ADOR.Recordset
    Set rptParms = rptSvr.GetParms(RptName)
    If rptParms.EOF Then MsgBox "No values returned"
    Do Until rptParms.EOF
        With lstCriteriaVals
            .AddItem rptParms(0).Value & ", ="
        End With
        rptParms.MoveNext
    Loop
    MousePointer = vbDefault
    Set rptParms = Nothing
    ssdbRptNames.Enabled = False

 End Sub


Private Sub lstCriteriaVals_Click()
    cmdSendMail.Enabled = True
End Sub


Private Sub lstFileNames_Click()

On Error GoTo CommandErr:
    
    Call GetReports(lstFileNames.Text)
    ssdbRptNames.SetFocus
    lstFileNames.Enabled = False
CommandErr:
     RealError = Err.Number
    If RealError < 0 Then RealError = (RealError) * -1
    If RealError > 0 Then
        Select Case RealError
            Case UserNotFoundErr
                MsgBox "User Not Found!!!  Contact your administrator.", vbOKOnly, "User Name Changed or not Present"
            Case DataBaseErr
                MsgBox "Database is busy, another user has the database opened exclusively." & vbCrLf & vbCrLf & _
                            "Please try again momentarily." & vbCrLf & vbTab & "Err: " & Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
            Case SQLErr
                MsgBox "There was an error building the SQL Statement on the Server." & vbCrLf & vbCrLf & _
                            "Please try again or contact your administrator if the problem continues." & vbCrLf & vbTab & "Err: " & _
                            Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
            Case UnknownErr
                MsgBox "An unknown error was detected.", vbOKOnly, "Customization Component Error"
            Case Else
                MsgBox Err.Description, vbOKOnly, "There was an error on the server"
                Set rptSvr = Nothing
                Set rptSvr = CreateObject("ReportServer.SendReport")
        End Select
    End If
    MousePointer = vbDefault
End Sub


Private Sub ssdbRptNames_Validate(Cancel As Boolean)

On Error GoTo CommandErr
    txtTO.SetFocus
    GetParameters ssdbRptNames.Text

CommandErr:
    RealError = Err.Number
    
    If RealError <> 0 Then
        Select Case RealError
            Case UserNotFoundErr
                MsgBox "User Not Found!!!  Contact your administrator.", vbOKOnly, "User Name Changed or not Present"
                MousePointer = vbDefault
            Case DataBaseErr
                MsgBox "Database is busy, another user has the database opened exclusively." & vbCrLf & vbCrLf & _
                            "Please try again momentarily." & vbCrLf & vbTab & "Err: " & Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
                MousePointer = vbDefault
            Case SQLErr
                MsgBox "There was an error building the SQL Statement on the Server." & vbCrLf & vbCrLf & _
                            "Please try again or contact your administrator if the problem continues." & vbCrLf & vbTab & "Err: " & _
                            Err.Number & vbCrLf & vbTab & "Description: " & Err.Description _
                            , vbOKOnly, "Database in use"
                MousePointer = vbDefault
            Case UnknownErr
                MsgBox "An unknown error was detected.", vbOKOnly, "Customization Component Error"
                MousePointer = vbDefault
            Case Else
                MsgBox "There was an undetected Run-time error returned from the server" & vbCrLf & vbTab & "Err: " & _
                            Err.Number & vbCrLf & vbTab & "Description: " & Err.Description, vbOKOnly, "Server Run-Time Error"
                MousePointer = vbDefault
        End Select
    End If
End Sub
