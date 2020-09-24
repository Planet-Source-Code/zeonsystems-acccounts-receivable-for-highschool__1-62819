VERSION 5.00
Begin VB.Form frmDiscountOthers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Brothers/Sisters"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2280
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
   ScaleHeight     =   960
   ScaleWidth      =   2280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   510
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   510
      Width           =   1005
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1290
      MaxLength       =   2
      TabIndex        =   0
      Top             =   90
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount [%]:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmDiscountOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tDisc As Currency
Dim rsLevelList As New ADODB.Recordset
Dim rs_new As New ADODB.Recordset

Private Sub Form_Load()
    rsLevelList.Open "Select * From tblLevel Where YR='" & frmPayment.rs_id("YR") & "'", CN, adOpenStatic, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsLevelList = Nothing
End Sub

Private Sub Command1_Click()

If is_empty(Text2) = True Then Exit Sub
If not_curr(Text2) = True Then Exit Sub

If frmAddDiscount.rs_new.RecordCount = 1 Or Not frmAddDiscount.rs_new.EOF Then
    MsgBox "Students can't avail two discounts. Delete the discount before updating.", vbInformation: Exit Sub
End If

tDisc = rsLevelList!Amount * Text2 / 100

With frmAddDiscount.rs_new
    .AddNew
    .Fields("AccountNo") = frmPayment.AcctNo
    .Fields("DiscountName") = frmAddDiscount.Combo1.Text
    .Fields("Amount") = tDisc
    .Update
    .Requery
End With

frmPayment.fill_value
frmAddDiscount.Fill_record

Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc("."), vbKeyBack
        Case Else: KeyAscii = 0
    End Select
End Sub
