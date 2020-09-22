VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1050
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2925
      Width           =   6645
   End
   Begin VB.TextBox Text1 
      Height          =   1110
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   1350
      Width           =   6660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   900
      Left            =   1335
      TabIndex        =   0
      Top             =   180
      Width           =   4035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strVal As String, i, x
Dim startFind As Integer
Dim endFind As Integer
Dim strParm As String
Dim startParmFind As Integer
Dim endParmFind As Integer
Dim CheckAnd As Integer
Dim CheckOr As Integer
Dim arnames(1) As String
Dim ParmCount As Integer

On Error GoTo CompErr
    strwhere = Text1.Text
    arnames(0) = "[Zip Code]"
    ParmCount = 2
    
    startFind = InStr(1, strwhere, "(")
    startParmFind = InStr(1, strwhere, "[")
    endParmFind = InStr(startParmFind + 1, strwhere, "]")
    endFind = InStr(endParmFind, strwhere, ")")
    
    For i = 0 To ParmCount - 1
        strVal = Mid(strwhere, startFind, endFind - startFind + 1)
        startParmFind = InStr(startFind, strwhere, "[")
        endParmFind = InStr(startFind + 1, strwhere, "]")
        If startParmFind > 0 Then
            strParm = Mid(strwhere, startParmFind + 1, endParmFind - startParmFind - 1)
            strParm = "[" & strParm & "]"
            For x = 0 To UBound(arnames) - 1
                If strParm <> arnames(x) Then
                    startFind = InStr(1, strwhere, "(")
                    endFind = InStr(startFind + 1, strwhere, ")")
                Else
                    startFind = InStr(endFind + 1, strwhere, "(")
                    startParmFind = InStr(startFind + 1, strwhere, "[")
                    endParmFind = InStr(startParmFind, strwhere, "]")
                    endFind = InStr(endParmFind, strwhere, ")")
                End If
            Next
        End If
    Next
    Text2.Text = strwhere

CompErr:
    Select Case Err.Number
        Case Is > 0
            Err.Description = "There was a problem building the SQL Statement, please wait a few moments and try again."
            Err.Number = SQLErr
            Err.Raise SQLErr
            Exit Sub
        Case Is = 0
            Exit Sub
    End Select

End Sub
