VERSION 5.00
Begin VB.Form frmAddPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Transaction"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAddPayment.frx":0000
      Left            =   1200
      List            =   "frmAddPayment.frx":000A
      TabIndex        =   1
      Top             =   240
      Width           =   2955
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1080
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2475
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Payment of:"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      Height          =   195
      Left            =   810
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmAddPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public User As String
Dim New_Record As Boolean
Public oPos As Integer
Dim rs_or As New ADODB.Recordset

Private Sub Command1_Click()
If is_empty(Combo1) = True Then Exit Sub
If is_empty(Text1) = True Then Exit Sub
If not_curr(Text1) = True Then Exit Sub
If Text1 <= 0 Then MsgBox "Please enter a valid amount.", vbExclamation: Exit Sub
If Text1 > CDbl(frmPayment.Text5) Then MsgBox "Amount is greater than the amount due.", vbExclamation: _
    Text1.SetFocus: SendKeys "{Home}+{End}": Exit Sub
    
If frmAddPayment.Caption = "New Payment Transaction" Then
    With rs_payment
        .AddNew
        .Fields("AccountNo") = frmPayment.AcctNo
        .Fields("Description") = Combo1.Text
        .Fields("Amount") = Text1.Text
        .Fields("Date") = Date
        .Update
        .Requery
    End With
End If
frmPayment.fill_value
rs_payment.MoveLast
oPos = rs_payment!ORNO

Call show_report(oPos)
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Public Sub show_report(ByVal sOR As String)
frmPayment.fill_value
rs_or.Open "Select * From tblPayment", CN, adOpenKeyset, adLockOptimistic
rs_or.Filter = "ORNo='" & sOR & "'"

Set Dtr_OR.DataSource = rs_or

Dtr_OR.Show vbModal

Unload Me

End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    Text2 = frmLogin.Text1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_or = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc("."), vbKeyBack
        Case Else: KeyAscii = 0
    End Select
End Sub

