VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddDiscount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discount"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
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
   ScaleHeight     =   2070
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1410
      TabIndex        =   1
      Top             =   1620
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmAddDiscount.frx":0000
      Left            =   450
      List            =   "frmAddDiscount.frx":0002
      TabIndex        =   0
      Text            =   "Brothers/Sisters"
      Top             =   150
      Width           =   3945
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3510
      TabIndex        =   3
      Top             =   1620
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   795
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   1402
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Discount Name"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddDiscount.frx":0004
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List:"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   300
   End
End
Attribute VB_Name = "frmAddDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim field_name As String
Dim field_amt As Currency
Public rs_new As New ADODB.Recordset

Private Sub Form_Load()
    Call fill_discount(Combo1)
    rs_new.Open "Select * From tblDiscount Where AccountNo = '" & frmPayment.AcctNo & "' AND DiscountName<>'" & "None" & "'", CN, adOpenKeyset, adLockOptimistic
    Fill_record
End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
    frmDiscountOthers.Show vbModal
Else
    find_field
End If
End Sub
'Add discount of student
Private Sub Command1_Click()
If is_empty(Combo1) = True Then Exit Sub
rs_new.Requery
rs_new.Find "DiscountName='" & Combo1.Text & "'"
      
    If rs_new.RecordCount = 1 Or Not rs_new.EOF Then
        MsgBox "Students can't avail two discounts. Delete the discount before updating.", vbInformation: Exit Sub
    End If
    
    With rs_new
        .AddNew
        .Fields("AccountNo") = frmPayment.AcctNo
        .Fields("DiscountName") = field_name
        .Fields("Amount") = field_amt
        .Update
        .Requery
    End With
    
    Fill_record
    frmPayment.fill_value
    Unload Me
End Sub

'Remove discount of student
Private Sub Command2_Click()

If rs_new.RecordCount = 0 Then MsgBox "No discount in the list. Please check it!", vbExclamation, AppTitle: Exit Sub

If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, "Delete record") = vbYes Then
    rs_new.AbsolutePosition = ListView1.SelectedItem.Index
    With rs_new
        .Delete
        .Requery
    End With
    
Fill_record
frmPayment.fill_value
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub find_field()
Dim rs As New ADODB.Recordset

rs.Open "Select * from tblDiscountType", CN, adOpenStatic, adLockReadOnly

rs.Requery

rs.Find "DiscountName='" & Combo1.Text & "'"
field_name = rs!DiscountName
field_amt = rs!Amount

Set rs = Nothing
End Sub

Public Sub Fill_record()
'rs_discount.Requery

With ListView1
    .ListItems.Clear
    Do While Not rs_new.EOF
        .ListItems.Add , , rs_new!DiscountName, 1, 1
        .ListItems(.ListItems.Count).SubItems(1) = Format(IIf(IsNull(rs_new!Amount), "", rs_new!Amount), "#,##0.00")
        rs_new.MoveNext
    Loop
    If rs_new.RecordCount > 0 Then rs_new.AbsolutePosition = ListView1.SelectedItem.Index
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_new = Nothing
End Sub

Private Sub ListView1_Click()
    If rs_new.RecordCount < 1 Then Exit Sub
         rs_new.AbsolutePosition = ListView1.SelectedItem.Index
End Sub
