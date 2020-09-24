VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment Transaction"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "&Update Year Level"
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Calculator"
      Height          =   375
      Left            =   6360
      TabIndex        =   28
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Edit Payment"
      Height          =   375
      Left            =   240
      TabIndex        =   27
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   3330
      Width           =   1275
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0.00"
      Top             =   4230
      Width           =   1275
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   4800
      Width           =   1275
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   3810
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7140
      Picture         =   "frmPayment.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3780
      Width           =   315
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7110
      Picture         =   "frmPayment.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3330
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   900
      TabIndex        =   0
      Top             =   210
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Search"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3150
      TabIndex        =   1
      Top             =   150
      Width           =   1155
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&New Payment"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5070
      TabIndex        =   2
      Top             =   6510
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   6510
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   2535
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Width           =   7455
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1230
         TabIndex        =   16
         Top             =   1980
         Width           =   5865
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1230
         TabIndex        =   15
         Top             =   1590
         Width           =   5865
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1230
         TabIndex        =   14
         Top             =   1200
         Width           =   5865
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1230
         TabIndex        =   13
         Top             =   810
         Width           =   5865
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1230
         TabIndex        =   12
         Top             =   420
         Width           =   5865
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year Level:"
         Height          =   195
         Left            =   210
         TabIndex        =   9
         Top             =   1980
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No.:"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   1590
         Width           =   930
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   315
         Left            =   375
         TabIndex        =   7
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   555
         TabIndex        =   6
         Top             =   810
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID No.:"
         Height          =   195
         Left            =   495
         TabIndex        =   5
         Top             =   420
         Width           =   525
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         Height          =   2205
         Left            =   1110
         Top             =   210
         Width           =   6195
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   150
      TabIndex        =   10
      Top             =   3240
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   4194304
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "OR No."
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Desc. of payment"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   4650
      X2              =   7440
      Y1              =   4650
      Y2              =   4650
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   4650
      X2              =   7440
      Y1              =   4650
      Y2              =   4650
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fee:"
      Height          =   195
      Left            =   4680
      TabIndex        =   26
      Top             =   3390
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Discount:"
      Height          =   195
      Left            =   4680
      TabIndex        =   25
      Top             =   3840
      Width           =   1080
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Due:"
      Height          =   195
      Left            =   4680
      TabIndex        =   24
      Top             =   4290
      Width           =   945
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      Height          =   195
      Left            =   4710
      TabIndex        =   23
      Top             =   4830
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID NO.:"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   240
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   150
      X2              =   7560
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   150
      X2              =   7560
      Y1              =   6420
      Y2              =   6420
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs_id  As New ADODB.Recordset
Public no_rec As Boolean
Public AcctNo As String
Public rs_acct As New ADODB.Recordset
Dim rs_old_acct As New ADODB.Recordset

Private Sub Command6_Click()

If rs_payment.RecordCount = 0 Then MsgBox "No payment in the list. Please check it!", vbExclamation: Exit Sub

With frm_edit_payment
    .Caption = "Modify Record"
    .Show vbModal
End With

End Sub

Private Sub Command7_Click()
    frmCalculator.Show vbModal
End Sub

Private Sub Command8_Click()
    Unload frmStudent
    If rs_payment.RecordCount = 0 Then
        Command8.Enabled = False
        MsgBox "You can't update the student year level because he has no payment transaction yet."
    Else
        Command8.Enabled = True
        frmUpdateLevel.Show vbModal
    End If
End Sub

Private Sub Form_Load()
    rs_id.Open "Select * From qryStudentList", CN, adOpenKeyset, adLockOptimistic
    rs_payment.Open "Select * From tblPayment", CN, adOpenKeyset, adLockOptimistic
    rs_old_acct.Open "Select * From tblAccount", CN, adOpenKeyset, adLockOptimistic
    no_rec = False
    DisableControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_id = Nothing
    Set rs_acct = Nothing
    Set rs_payment = Nothing
    Set rs_old_acct = Nothing
End Sub

Private Sub Command2_Click()
On Error Resume Next

rs_id.Requery
rs_id.Find "ID='" & Text1.Text & "'"
If rs_id.EOF Then
    
    no_rec = False
    Call Message
    
    Text1.SetFocus
    Call highlight_focus(Text1)
    DisableControls
    Exit Sub
    If rs_payment.RecordCount = 0 Then Command8.Enabled = False
Else
    If rs_id.Fields("Enrolled") = False Then
        If MsgBox("Record is not currently enrolled. Do you want to marked the record as enrolled?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            rs_id.Fields("Enrolled") = True
            rs_id.Update
            'frmUpdateLevel.Show 1
        Else
            Exit Sub
        End If
    End If
    Call EnableControls
    Call get_Balance(AcctNo, Text5, Text4)
    rs_payment.Filter = "AccountNo='" & AcctNo & "'"
    fill_value

End If
End Sub

Private Sub Command3_Click()
If no_rec = False Then MsgBox "No record found", vbExclamation: Exit Sub
    frmAddFee.Show vbModal
End Sub

Private Sub Command4_Click()
Command8.Enabled = True
If no_rec = False Then MsgBox "No record found.", vbExclamation: Exit Sub
If Text5 = "0.00" Then
    Unload frmStudent
    If MsgBox("Account has been paid. This will update the year level of student.", vbOKOnly + vbInformation) = vbOK Then
        frmUpdateLevel.Show vbModal
    End If
Exit Sub
End If
    With frmAddPayment
        .Caption = "New Payment Transaction"
        .Show vbModal
    End With
End Sub

Private Sub Command5_Click()
If no_rec = False Then MsgBox "No record found.", vbExclamation: Exit Sub
    frmAddDiscount.Show vbModal
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub ListView1_Click()
On Error Resume Next
    If rs_payment.RecordCount < 1 Then Exit Sub
    rs_payment.AbsolutePosition = ListView1.SelectedItem.Index
End Sub

Private Sub Text1_Change()

If Text1 = "" Then
    Command2.Enabled = False
Else
    Command2.Enabled = True
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Command2_Click
Select Case KeyAscii
Case Asc("0") To Asc("9"), vbKeyBack
Case Else
    KeyAscii = 0
End Select
End Sub

Public Sub get_acct()
Dim rs As New ADODB.Recordset
Dim dyr As String
dyr = Mid(Year(Now), 3, 4)
AcctNo = rs_id!ID & SysSY.SY
If rs_acct.State = adStateOpen Then rs_acct.Close
rs_acct.Open "Select * From tblAccount Where AccountNo = '" & AcctNo & "' And SY = '" & SysSY.SY & "'", CN, adOpenKeyset, adLockOptimistic
rs.Open "Select * From tblFees", CN, adOpenKeyset, adLockOptimistic
If rs_acct.EOF Then
    With rs_acct
        .AddNew
        .Fields("AccountNo") = AcctNo
        .Fields("ID") = rs_id!ID
        .Fields("SY") = SysSY.SY
        .Fields("AmountDue") = Text4.Text
        .Fields("Status") = False
        .Update
        .Requery
    End With
    frmUpdateLevel.Show 1
    Call add_discount(AcctNo)

    Call add_tuition(AcctNo, Label15)
    Call add_fees(AcctNo, SysSY.SY, Label15)
Else
    Call old_account(Text1, SysSY.SY, AcctNo)
    Call Acct_Disc(AcctNo, Text3)

Set rs = Nothing
End If

End Sub

Public Sub fill_value()
Dim payableFee As Currency
    
    Call Acct_Fees(AcctNo, Text2)
    Call Acct_Disc(AcctNo, Text3)
    payableFee = Text2 - Text3
    Text4 = Format$(payableFee, "#,##0.00")
    Call get_Balance(AcctNo, Text5, Text4)
    Call Fill_record
End Sub

Public Sub Fill_record()

rs_payment.Requery

With ListView1
    .ListItems.Clear
    Do While Not rs_payment.EOF
        .ListItems.Add , , rs_payment!ORNO
        .ListItems(.ListItems.Count).SubItems(1) = IIf(IsNull(rs_payment!Date), "", rs_payment!Date)
        .ListItems(.ListItems.Count).SubItems(2) = IIf(IsNull(rs_payment!Description), "", rs_payment!Description)
        .ListItems(.ListItems.Count).SubItems(3) = Format(IIf(IsNull(rs_payment!Amount), "", rs_payment!Amount), "#,##0.00")
        rs_payment.MoveNext
    Loop
    If rs_payment.RecordCount > 0 Then rs_payment.AbsolutePosition = ListView1.SelectedItem.Index
End With

UpdateAccount
End Sub

Sub DisableControls()

    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Label11 = "--"
    Label12 = "--"
    Label13 = "--"
    Label14 = "--"
    Label15 = "--"
    Text2 = "0.00"
    Text3 = "0.00"
    Text4 = "0.00"
    Text5 = "0.00"

End Sub

Sub EnableControls()

no_rec = True
    With rs_id
        Label11 = .Fields("ID")
        Label12 = .Fields("Name")
        Label13 = .Fields("Address")
        Label14 = .Fields("ContactNo")
        Label15 = .Fields("YR")
    End With
    
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    Call get_acct
End Sub

'Update acount balance
Sub UpdateAccount()
    
    With rs_acct
        .Fields("AmountDue") = Text4.Text
        .Fields("Balance") = Text5.Text
        .Update
        .Requery
    End With

End Sub

Sub change_rec()
Call Command2_Click
End Sub
