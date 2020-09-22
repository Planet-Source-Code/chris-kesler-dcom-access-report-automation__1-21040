VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for Rental Rates (Client)"
   ClientHeight    =   1476
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   4176
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1476
   ScaleWidth      =   4176
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGetRate 
      Caption         =   "&Get Rate"
      Height          =   264
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   852
   End
   Begin MSComCtl2.DTPicker dtpRentalDate 
      Height          =   264
      Left            =   1176
      TabIndex        =   2
      Top             =   600
      Width           =   1056
      _ExtentX        =   1863
      _ExtentY        =   466
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   36281
   End
   Begin VB.TextBox txtCategoryId 
      Height          =   288
      Left            =   1176
      TabIndex        =   0
      Text            =   "EC"
      Top             =   168
      Width           =   432
   End
   Begin VB.Label lblResult 
      BorderStyle     =   1  'Fixed Single
      Height          =   264
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   3888
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Rental Date:"
      Height          =   192
      Left            =   240
      TabIndex        =   3
      Top             =   660
      Width           =   888
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Category ID:"
      Height          =   192
      Left            =   216
      TabIndex        =   1
      Top             =   240
      Width           =   888
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRentalRates As CRentalRates

Private Sub cmdGetRate_Click()
'Search for the Vehicle category and rental date, using the
'CRentalRates object.
  Dim curDailyRate As Currency, curWeeklyRate As Currency

  If mRentalRates.GetRentalRates(txtCategoryId.Text, _
    dtpRentalDate.Value, curDailyRate, curWeeklyRate) Then
    lblResult.Caption = "Daily rate = " _
      & Format(curDailyRate, "Currency") _
      & ", Weekly rate = " _
      & Format(curWeeklyRate, "Currency")
  Else
    lblResult.Caption = "Rental rates not found"
  End If
End Sub

Private Sub Form_Load()
'Create an instance of CRentalRates. This
'will open the database and recordset.
  Set mRentalRates = New CRentalRates
End Sub
