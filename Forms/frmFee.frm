VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fees"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
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
   ScaleHeight     =   5595
   ScaleWidth      =   4650
   Begin VB.CommandButton Command1 
      Caption         =   "&New Record"
      Default         =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   6
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   1380
      TabIndex        =   5
      Top             =   5100
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      Top             =   5100
      Width           =   1005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3540
      TabIndex        =   3
      Top             =   5100
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4275
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   7541
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
         Text            =   "Fee Name"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
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
            Picture         =   "frmFee.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fees Record"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "This form show the different type of fees of student"
      Height          =   195
      Left            =   420
      TabIndex        =   0
      Top             =   390
      Width           =   3765
   End
End
Attribute VB_Name = "frmFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    rs_fees.Open "Select * From tblFeesType Order By FeeName Asc", CN, adOpenKeyset, adLockOptimistic
    Fill_record
    frmFee.Left = MDIForm1.ScaleWidth / 2 - frmFee.ScaleWidth / 2
    frmFee.Top = MDIForm1.ScaleHeight / 2 - frmFee.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_fees = Nothing
End Sub

'New Record
Private Sub Command1_Click()
With frm_ae_fees
    .New_Rec = True
    .Show vbModal
End With
End Sub

'Modify Record
Private Sub Command2_Click()

If rs_fees.RecordCount < 1 Then MsgBox "No fees in the list. Please check it!", vbExclamation: Exit Sub

With frm_ae_fees
    .New_Rec = False
    .Caption = "Modify Record"
    .Show vbModal
End With
End Sub

'Delete Record
Private Sub Command3_Click()
If rs_fees.RecordCount = 1 Then MsgBox "You are not allowed to delete all the records. There must be at  least one record left.", vbExclamation: Exit Sub
If rs_fees.RecordCount < 1 Then MsgBox "No fees in the list. Please check it!", vbExclamation: Exit Sub
If rs_fees.Fields("Standard") = True Then MsgBox "Can't delete this fee. It is a standard fee.", vbExclamation: Exit Sub

If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, "Delete record") = vbYes Then

    With rs_fees
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
rs_fees.Requery

With ListView1
    .ListItems.Clear
    Do While Not rs_fees.EOF
        .ListItems.Add , , rs_fees.Fields("FeeName"), 1, 1
        .ListItems(.ListItems.Count).SubItems(1) = Format$(IIf(IsNull(rs_fees("Amount")), "", rs_fees("Amount")), "#,##0.00")
        rs_fees.MoveNext
    Loop
    If rs_fees.RecordCount > 0 Then rs_fees.AbsolutePosition = ListView1.SelectedItem.Index
End With
End Sub

Private Sub ListView1_Click()

    If rs_fees.RecordCount = 0 Then Exit Sub
    rs_fees.AbsolutePosition = ListView1.SelectedItem.Index
    Exit Sub

End Sub
