VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SendReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************
'DCOM Access Server for Report Generation
'Created On: 01/15/2001
'Created By: Chris Kesler
'Created For: Arch Wireless/Pagenet (Non Exclusive)
'**************************************************

Option Explicit

Public Enum CustErrorType
    UserNotFoundErr = 1000
    DataBaseErr
    NoDataErr
    SQLErr
    UnknownErr
    DivideByZeroErr
End Enum

Private cnAccServ           As New adodb.Connection
Private rsAccServ           As New adodb.Recordset
Private AccSendRpt         As New Access.Application
Private dbs                 As DAO.Database
Private tdf                 As DAO.TableDef, fld As DAO.Field
Private rst                 As DAO.Recordset
Private rpt                 As New Access.Report
Private qryDef          As DAO.QueryDef
Private MailSession         As MAPISession
Private MailMessages        As MAPIMessages
Private mSessionIsOpen      As Boolean
Private mLoggedOnUser       As String
Private sDBName               As String
Private SQLTemp             As String

Public Property Get DBVals() As Variant
    Dim RealError As Long, i
    Dim mdbFileName As String
    Dim tmpStrg As String
    Dim DBArray()
    
    On Error GoTo CommandErr
    
        tmpStrg = Dir$(App.Path & "\*.mdb")
        i = 0
        
        'Look for mdb files (If you want to change to any other type,
        ' just replace *.mdb with another extension (E.g *.bmp)
        If tmpStrg <> "" Then 'have mdb's in the directory
            mdbFileName = Left$(tmpStrg, Len(tmpStrg) - 4)
            ReDim Preserve DBArray(i)
            DBArray(i) = mdbFileName  'Add the mdb's to the list box
            tmpStrg = Dir$ 'Go back to the directory to add more mdb's
            While Len(tmpStrg) > 0 'While there is still more unadded mdb's
                i = i + 1
                mdbFileName = Left$(tmpStrg, Len(tmpStrg) - 4)
                ReDim Preserve DBArray(i)
                DBArray(i) = mdbFileName  'Add the mdb
                tmpStrg = Dir$ 'Go back to the directory to add more mdb's
            Wend
        Else 'There isn't any mdb's in the directory
            DBArray(0) = "There were no mdb's available"
        End If
        DBVals = DBArray()
        
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
    
End Property
'*************************************************************
'Procedure:    Public Method BuildAndSendRpt
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'sTo
'sSubject
'strRptName
'sCC
'sBCC
'sMessageText
'arCriteriaVals(
'arCriteriaNames(
'*************************************************************

Public Function BuildAndSendRpt(sTo As String, _
                                              sSubject As String, _
                                              strRptName As String, _
                                              sMessageText As String, _
                                              arCriteriaVals() As String, _
                                              arCriteriaNames() As String) As Boolean
On Error GoTo SendErr
        
    '**************Start Build Report and Export to File
    ModifyQuery App.Path & "\" & sDBName, strRptName, False, arCriteriaVals(), arCriteriaNames
    timedPause 2
    AccSendRpt.DoCmd.OutputTo acSendReport, strRptName, acFormatRTF, App.Path & "\" & strRptName & ".rtf"
    timedPause 2
    '**************End Build Report and Export to File
    
    '**************Start Email build and send
    Call SignOn
    Call Send(sTo, sSubject, sMessageText, strRptName & ".rtf")
    Call SignOff
    '**************End Email build and send
    
    '**************Change Report Parameters back to original Parameter Values
    ModifyQuery App.Path & "\" & sDBName, strRptName, True, arCriteriaVals(), arCriteriaNames()
    timedPause 2
    AccSendRpt.CloseCurrentDatabase
    
SendErr:
    If Err.Number > 0 And Err.Number < 65535 Then
        Err.Description = "There was a problem opening the database, please wait a few moments and try again."
        Err.Number = DataBaseErr
        Err.Raise DataBaseErr
        BuildAndSendRpt = False
    Else
        BuildAndSendRpt = True
    End If
    
ExitSend:
    Exit Function
    
End Function

'*************************************************************
'Procedure:    Public Method GetRptNames
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'
'*************************************************************

Public Function GetRptNames(DBName As String) As adodb.Recordset
    Dim x, i
    Dim objDocs As Documents
    Dim tblDocs As Documents
    Dim RealError As Long
    
 On Error GoTo CompErr:
    sDBName = DBName & ".mdb"
    Call SetConnect(sDBName)
    Set tblDocs = dbs.Containers!Tables.Documents
    For i = 0 To tblDocs.Count - 1
        'If Table exists, delete table for fresh entry
        If tblDocs(i).Name = "ReportNames" Then
            dbs.TableDefs.Delete "ReportNames"
        End If
    Next
    'Create table and insert field to store temporary report names.
    Set tdf = dbs.CreateTableDef("ReportNames")
    Set fld = tdf.CreateField("rptName", dbText)
    tdf.Fields.Append fld
    dbs.TableDefs.Append tdf
    Set rst = dbs.OpenRecordset("ReportNames")
    Set objDocs = dbs.Containers!Reports.Documents
    
    For x = 0 To objDocs.Count - 1
        With rst
            .AddNew
            rst("rptName").Value = objDocs(x).Name
            .Update
        End With
    Next
    rst.Close
    
    'Open and retrieve temporary report names to return to client in an
    'ADO recordset.
    If cnAccServ.State = adStateOpen Then
        cnAccServ.Close
    End If

    cnAccServ.ConnectionString = "Provider = " _
                        & "Microsoft.Jet.OLEDB.4.0; " _
                        & "Data Source =" & App.Path & "\" & sDBName
    
    cnAccServ.Open
    timedPause 2
    Set GetRptNames = cnAccServ.Execute("SELECT * FROM REPORTNAMES")
    timedPause 2
    If GetRptNames.RecordCount = 0 Then
        Err.Number = NoDataErr
    End If

CompErr:
    RealError = Err.Number
    If RealError < 0 Then RealError = (RealError) * -1
    Select Case RealError
        Case DataBaseErr
            Err.Description = "There was a problem opening the database, please wait a few moments and try again."
            Err.Number = DataBaseErr
            Err.Raise DataBaseErr
            If cnAccServ.State = adStateOpen Then cnAccServ.Close
            Set GetRptNames = Nothing
            Exit Function
        Case NoDataErr
            Err.Description = "No Data was found to return"
            Err.Number = NoDataErr
            Err.Raise NoDataErr
            Exit Function
        Case 7867
            Err.Description = "Database is already open, if there was an error please click continue to reset connection."
            Err.Raise vbObjectError, App.EXEName
            Exit Function
        Case Is > 0
            Err.Description = "A database parameter or object name was not found or defined"
            Err.Raise vbObjectError, App.EXEName
            Exit Function
        Case Is = 0
            Exit Function
    End Select
    
End Function

'*************************************************************
'Procedure:    Public Method GetParms
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'ReportName
'*************************************************************

Public Function GetParms(ReportName As String) As adodb.Recordset
    Dim x, i
    Dim tblDocs As Documents
    
On Error GoTo CompErr:
    
    Set tblDocs = dbs.Containers!Tables.Documents
    For i = 0 To tblDocs.Count - 1
        'If Table exists, delete table for fresh entry
        If tblDocs(i).Name = "tblParameters" Then
            dbs.TableDefs.Delete "tblParameters"
        End If
    Next
    'Create table and insert field to store temporary report names.
    Set tdf = dbs.CreateTableDef("tblParameters")
    Set fld = tdf.CreateField("Parms", dbText)
    tdf.Fields.Append fld
    Set fld = tdf.CreateField("ParmsName", dbText)
    tdf.Fields.Append fld
    dbs.TableDefs.Append tdf
    Set qryDef = dbs.QueryDefs("qry" & ReportName)
    Set rst = dbs.OpenRecordset("tblParameters")
    For x = 0 To qryDef.Parameters.Count - 1
        With rst
            .AddNew
            timedPause 1
            rst("Parms").Value = qryDef.Parameters(x).Name
            .Update
        End With
    Next
    rst.Close
    
    'If connection is open close the connection
    If cnAccServ.State = adStateOpen Then
        cnAccServ.Close
    End If
    
    'Open and retrieve temporary report names to return to client in an
    'ADO recordset.
    cnAccServ.ConnectionString = "Provider = " _
                        & "Microsoft.Jet.OLEDB.4.0; " _
                        & "Data Source =" & App.Path & "\" & sDBName
    
    cnAccServ.Open
    timedPause 2
    Set GetParms = cnAccServ.Execute("SELECT * FROM tblPARAMETERS")
    If GetParms.EOF Then
        Err.Number = NoDataErr
    End If
    timedPause 2
        
CompErr:
    Select Case Err.Number
        Case NoDataErr
            Err.Description = "There was a problem opening the database, please wait a few moments and try again."
            Err.Number = NoDataErr
            Err.Raise NoDataErr
            If cnAccServ.State = adStateOpen Then cnAccServ.Close
            Set GetParms = Nothing
            Exit Function
        Case 0
            Exit Function
    End Select
        
End Function
Private Sub Class_Initialize()
'Open a database connection and set database.

End Sub

'*************************************************************
'Procedure:    Private Method ModifyQuery
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'strDBPath
'strQryName
'resetVals
'arParms(
'*************************************************************

Private Sub ModifyQuery(strDBPath As String, _
                                  strQryName As String, _
                                  resetVals As Boolean, _
                                  arParms() As String, _
                                  arnames() As String)
                                  
    Dim catDB As ADOX.Catalog
    Dim cmd As adodb.Command
    Dim CurSql As String
    Dim strWhereClause As String
    Dim intPos As Integer
    Dim newSQL As String
    Dim strSQL As String
    Dim ParmCount As Integer
    Dim i, x
On Error GoTo CompErr:
    
    Set catDB = New ADOX.Catalog
    catDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & strDBPath
    Set cmd = New adodb.Command
    If resetVals = True Then
        Set cmd = catDB.Views("qry" & strQryName).Command
    Else
        Set cmd = catDB.Procedures("qry" & strQryName).Command
    End If
    CurSql = cmd.CommandText
    intPos = InStr(1, CurSql, "WHERE")
    ParmCount = cmd.Parameters.Count
    
    If intPos > 0 Then
            strWhereClause = Right(CurSql, Len(CurSql) - (intPos - 1))
            newSQL = Left(CurSql, intPos - 1)
    End If
    If resetVals = False Then
        'String value that hold the original Where Clause
        SQLTemp = strWhereClause
        strWhereClause = CreateNewWhere(strWhereClause, arParms())
    Else
        'Set the Where Clause back to the original from the Global Temp Storage
        strWhereClause = SQLTemp
    End If
    'Re-paste the Where Clause back on to the SQL String
    strSQL = newSQL & " " & strWhereClause
    cmd.CommandText = strSQL
    If resetVals = True Then
        Set catDB.Views("qry" & strQryName).Command = cmd
    Else
        Set catDB.Procedures("qry" & strQryName).Command = cmd
    End If
    
    timedPause 2
    
CompErr:
    Select Case Err.Number
        Case Is > 0
            Err.Description = "There was a problem opening the database, please wait a few moments and try again."
            Err.Number = DataBaseErr
            Err.Raise DataBaseErr
            Set catDB = Nothing
            Set cmd = Nothing
            Exit Sub
        Case Is = 0
            Exit Sub
    End Select
    
End Sub

'*************************************************************
'Procedure:    Private Method timedPause
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'secs
'*************************************************************

Private Sub timedPause(secs As Long)
    Dim secStart   As Variant
    Dim secNow     As Variant
    Dim secDiff    As Variant
    Dim Temp%
   
   secStart = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'get the starting seconds

    Do While secDiff < secs
        secNow = Format(Now(), "mm/dd/yyyy hh:nn:ss AM/PM") 'this is the current time and Date at any itteration of the Loop
        secDiff = DateDiff("s", secStart, secNow) 'this compares the start time With the current time
        Temp% = DoEvents
    Loop

End Sub

'*************************************************************
'Procedure:    Private Method CreateNewWhere
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'strwhere
'arvals(
'*************************************************************

Private Function CreateNewWhere(strwhere As String, arvals() As String) As String
Dim strVal As String, i
Dim startFind As Integer
Dim endFind As Integer
Dim PlusVal As Integer
Dim startVal As Integer
Dim endVal As Integer
Dim critVal As String
Dim intMoreVal As Integer

On Error GoTo CompErr:
    startVal = InStr(1, strwhere, "(")
    endVal = InStr(startVal, strwhere, ")")
    startFind = InStr(endVal, strwhere, "[")
    endFind = InStr(startFind, strwhere, "]")
    
    For i = 0 To UBound(arvals)
        If startFind > 0 Then
            If arvals(i) = "#" Then
                critVal = Mid(strwhere, startVal + 1, endVal - startVal - 1)
                critVal = Replace(critVal, "(", "", 1)
                strVal = Mid(strwhere, startFind, endFind - startFind + 1)
                strwhere = Replace(strwhere, strVal, critVal, 1)
                If Len(strVal) > Len(critVal) Then
                    intMoreVal = (Len(strVal) - Len(critVal))
                Else
                    intMoreVal = (Len(critVal) - Len(strVal))
                End If
            Else
                strVal = Mid(strwhere, startFind + 1, endFind - startFind - 1)
                strwhere = Replace(strwhere, strVal, arvals(i), 1)
                If Len(strVal) > Len(arvals(i)) Then
                    PlusVal = (Len(arvals(i)) - Len(strVal))
                Else
                    PlusVal = (Len(strVal) - Len(arvals(i)))
                End If
            End If
        End If
        startVal = InStr(endVal + intMoreVal, strwhere, "(")
        If startVal = 0 Then Exit For
        endVal = InStr(startVal + intMoreVal, strwhere, ")")
        startFind = InStr(endFind + PlusVal, strwhere, "[")
        If startFind = 0 Then Exit For
        endFind = InStr(startFind + PlusVal, strwhere, "]")
    Next
    strwhere = Replace(strwhere, "[", "'")
    strwhere = Replace(strwhere, "]", "'")
    CreateNewWhere = strwhere

CompErr:
    Select Case Err.Number
        Case Is > 0
            Err.Description = "There was a problem building the SQL Statement, please wait a few moments and try again."
            Err.Number = SQLErr
            Err.Raise SQLErr
            Exit Function
        Case Is = 0
            Exit Function
    End Select

End Function

'*************************************************************
'Procedure:    Private Method ResetWhere
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'strwhere
'arvals(
'*************************************************************

Private Function ResetWhere(strwhere As String, arvals() As String) As String
Dim strVal As String, i
Dim startFind As Integer
Dim endFind As Integer
Dim startVal As Integer
Dim endVal As Integer
Dim PlusVal As Integer

On Error GoTo CompErr

    startFind = InStr(1, strwhere, "'")
    endFind = InStr(startFind + 1, strwhere, "'")
    startVal = InStr(1, strwhere, "=")
    endVal = InStr(startVal, strwhere, ")")
    
    If startVal > 0 Then
        For i = 0 To UBound(arvals) - 1
            strVal = Mid(strwhere, startVal + 1, (endVal - 1) - (startVal + 1))
            strwhere = Replace(strwhere, strVal, arvals(i), 1)
            If Len(strVal) > Len(arvals(i)) Then
                PlusVal = (Len(arvals(i)) - Len(strVal))
            Else
                PlusVal = (Len(strVal) - Len(arvals(i)))
            End If
            startVal = InStr(endVal + PlusVal + 1, strwhere, "=")
            If startVal = 0 Then Exit For
            endVal = InStr(startVal + PlusVal + 1, strwhere, ")")
        Next

    End If
    strwhere = Replace(strwhere, "'", "")
    ResetWhere = strwhere

CompErr:
    Select Case Err.Number
        Case Is > 0
            Err.Description = "There was a problem building the SQL Statement, please wait a few moments and try again."
            Err.Number = SQLErr
            Err.Raise SQLErr
            Exit Function
        Case Is = 0
            Exit Function
    End Select

End Function

'*************************************************************
'Procedure:    Public Method SignOn
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'
'*************************************************************

Public Sub SignOn()
   
   On Error Resume Next
   Set MailSession = New MAPISession
   With MailSession
      'This will need to be changed to the server's Exchange Settings or Sims Agent
      'settings.
      .UserName = "MS Exchange Settings"
      'Password may need to be added to this at a later date.
      '.Password = "cmkes1er01"
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
      mLoggedOnUser = "MS Exchange Settings"
   End If
End Sub

'*************************************************************
'Procedure:    Public Method Send
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'Recipient
'Subject
'Message
'Attachment
'*************************************************************

Public Sub Send(Recipient As String, _
                Subject As String, _
                Message As String, _
                Attachment As String)
                
On Error GoTo TrapSendErrors
   Set MailMessages = New MAPIMessages
   With MailMessages
        .SessionID = MailSession.SessionID
        .Compose
        .RecipDisplayName = Recipient
        .MsgSubject = Subject
        .MsgNoteText = Message
        .AttachmentPathName = App.Path & "\" & Attachment
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

'*************************************************************
'Procedure:    Public Method SignOff
'Created on:   02/08/01
'Module:       SendReport
'Module File:  D:\AccessServerEx\SendReport.cls
'Project:      ReportServer
'Project File: D:\AccessServerEx\ReportServer.vbp
'Parameters:
'
'*************************************************************

Public Sub SignOff()
   On Error Resume Next
   MailSession.SignOff
End Sub

Public Sub SetConnect(DBName As String)

    AccSendRpt.OpenCurrentDatabase App.Path & "\" & DBName, False
    Set dbs = AccSendRpt.CurrentDb

End Sub

