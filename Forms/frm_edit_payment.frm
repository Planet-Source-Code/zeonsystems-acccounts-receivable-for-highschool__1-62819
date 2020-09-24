VERSION 5.00
Begin VB.Form frm_edit_payment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Payment"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1020
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   1020
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   600
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Left            =   2220
      TabIndex        =   5
      Top             =   630
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Payment of:"
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   300
      Width           =   840
   End
End
Attribute VB_Name = "frm_edit_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ORNUM As String

Private Sub Form_Load()
With rs_payment
    Text1.Text = .Fields("Description")
    Text2.Text = Format(.Fields("Amount"), "#,##0.00")
    ORNUM = .Fields("ORNo")
End With
End Sub

Private Sub Command1_Click()

With rs_payment
    .Fields("Description") = Text1.Text
    .Fields("Amount") = Text2.Text
    .Update
    .Requery
End With
MsgBox "Update successfull.", vbInformation
frmPayment.Fill_record
frmPayment.fill_value
Call show_report(ORNUM)

Unload Me

End Sub

Public Sub show_report(ByVal sOR As String)
Dim rs_or As New ADODB.Recordset
rs_or.Open "Select * From tblPayment", CN, adOpenKeyset, adLockOptimistic

rs_or.Filter = "ORNo='" & sOR & "'"
Set Dtr_OR.DataSource = rs_or
Dtr_OR.Show vbModal

Set rs_or = Nothing
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc("."), vbKeyBack
        Case Else: KeyAscii = 0
    End Select
End Sub
