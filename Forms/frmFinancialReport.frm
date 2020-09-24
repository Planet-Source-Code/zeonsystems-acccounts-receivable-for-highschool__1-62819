VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFinancialReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Report"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1290
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "&Close Window"
      Height          =   405
      Left            =   2940
      TabIndex        =   3
      Top             =   780
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Default         =   -1  'True
      Height          =   405
      Left            =   2940
      TabIndex        =   2
      Top             =   270
      Width           =   1635
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1230
      TabIndex        =   0
      Top             =   270
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Format          =   54460417
      CurrentDate     =   37993
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   1230
      TabIndex        =   1
      Top             =   780
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   393216
      Format          =   54460417
      CurrentDate     =   37993
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "End Date:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Start Date:"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   330
      Width           =   810
   End
End
Attribute VB_Name = "frmFinancialReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_finance As New ADODB.Recordset

Private Sub Command1_Click()
    Call Preview
End Sub

Private Sub Preview()
    rs_finance.Requery
    rs_finance.Filter = "Date>= '" & DTPicker1.Value & "' And Date<='" & DTPicker2.Value & "'"
    'If rs_finance.EOF Then Call Message: Exit Sub
    Set Dtr_Financial.DataSource = rs_finance
    Dtr_Financial.Show vbModal
End Sub

Private Sub DTPicker2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Preview
End Sub

Private Sub Form_Load()
    rs_finance.Open "Select * From qryFinance Order By YR, LastName Asc", CN, adOpenStatic, adLockReadOnly
  
    frmFinancialReport.Left = MDIForm1.ScaleWidth / 2 - frmFinancialReport.ScaleWidth / 2
    frmFinancialReport.Top = MDIForm1.ScaleHeight / 2 - frmFinancialReport.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set rs_finance = Nothing
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

