VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiscount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Discounts"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
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
   ScaleHeight     =   4155
   ScaleWidth      =   4590
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3540
      TabIndex        =   3
      Top             =   3660
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   3660
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   1380
      TabIndex        =   1
      Top             =   3660
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New Record"
      Default         =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   3660
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2835
      Left            =   90
      TabIndex        =   4
      Top             =   720
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   5001
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
      Left            =   2610
      Top             =   120
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
            Picture         =   "frmDiscount.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "This form show the different type of discount of student"
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   390
      Width           =   4050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Discount Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    rs_discount.Open "Select * From tblDiscountType Order By DiscountName Asc", CN, adOpenKeyset, adLockOptimistic
    Fill_record
    frmDiscount.Left = MDIForm1.ScaleWidth / 2 - frmDiscount.ScaleWidth / 2
    frmDiscount.Top = MDIForm1.ScaleHeight / 2 - frmDiscount.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_discount = Nothing
End Sub

'New Record
Private Sub Command1_Click()
With frm_ae_discount
    .new_rec = True
    .Show vbModal
End With
End Sub

'Modify Record
Private Sub Command2_Click()

If rs_discount.RecordCount < 1 Then MsgBox "No discount in the list. Please check it!", vbExclamation: Exit Sub

With frm_ae_discount
    .new_rec = False
    .Caption = "Modify Record"
    .Show vbModal
End With

End Sub

'Delete Record
Private Sub Command3_Click()

If rs_discount.RecordCount = 0 Then MsgBox "No discount in the list. Please check it!", vbExclamation: Exit Sub
If chk_data(ListView1.SelectedItem.Text) = True Then MsgBox "Can't delete this record because there are students who are using this discount name.", vbExclamation: Exit Sub
If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, "Delete record") = vbYes Then

    With rs_discount
        .Delete
        .Requery
    End With
    
Fill_record
End If

End Sub

'Close Window
Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub Fill_record()

rs_discount.Requery
With ListView1
    .ListItems.Clear
        Do While Not rs_discount.EOF
            .ListItems.Add , , rs_discount!DiscountName, 1, 1
            .ListItems(.ListItems.Count).SubItems(1) = Format$(IIf(IsNull(rs_discount!Amount), "", rs_discount!Amount), "#,##0.00")
            rs_discount.MoveNext
        Loop
    If rs_discount.RecordCount > 0 Then rs_discount.AbsolutePosition = ListView1.SelectedItem.Index
End With
End Sub

Private Sub ListView1_Click()

    If rs_discount.RecordCount = 0 Then Exit Sub
    rs_discount.AbsolutePosition = ListView1.SelectedItem.Index
    
End Sub
