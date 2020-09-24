VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmYear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Year Level & Section"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
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
   MDIChild        =   -1  'True
   ScaleHeight     =   4215
   ScaleWidth      =   4155
   Begin VB.CommandButton cmdSections 
      Caption         =   "&Sections"
      Height          =   375
      Left            =   -1080
      TabIndex        =   7
      Top             =   4020
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2190
      TabIndex        =   5
      Top             =   3720
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   1260
      TabIndex        =   4
      Top             =   3720
      Width           =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New Record"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2835
      Left            =   90
      TabIndex        =   0
      Top             =   690
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   5001
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Imagelist1"
      SmallIcons      =   "Imagelist1"
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
         Text            =   "Year Level & Section"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   2117
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
            Picture         =   "frmYear.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "This form show the list of year levels."
      Height          =   255
      Left            =   420
      TabIndex        =   2
      Top             =   330
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Level Records"
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
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Width           =   1185
   End
End
Attribute VB_Name = "frmYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rec As String

Private Sub cmdSections_Click()
    If rs_level.RecordCount < 1 Then MsgBox "No level in the list.Please check it!", vbExclamation: Exit Sub
        FRMSECTIONS.Caption = "Sections in " & rs_level.AbsolutePosition & " " & "Year"
        FRMSECTIONS.Show vbModal
End Sub

Private Sub Form_Load()
    rs_level.Open "Select * From tblLevel Order By YR Asc", CN, adOpenKeyset, adLockOptimistic
    Fill_record
    frmYear.Left = MDIForm1.ScaleWidth / 2 - frmYear.ScaleWidth / 2
    frmYear.Top = MDIForm1.ScaleHeight / 2 - frmYear.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_level = Nothing
End Sub
'New Record
Private Sub Command1_Click()

    With frm_ae_level
        .new_rec = True
        .Show vbModal
    End With
End Sub
'Modify Record
Private Sub Command2_Click()

    If rs_level.RecordCount < 1 Then MsgBox "No level in the list.Please check it!", vbExclamation: Exit Sub
    
        With frm_ae_level
            .new_rec = False
            .Caption = "Modify Record"
            .Text1.Locked = True
            .Show vbModal
        End With
            
End Sub

'Delete Record
Private Sub Command3_Click()

If rs_level.RecordCount < 1 Then MsgBox "No year level in the list.Please check it!", vbExclamation: Exit Sub
If check_data(ListView1.SelectedItem.Text) = True Then MsgBox "Can't delete this record because there are students who are enrolled in this year level and section.", vbExclamation: Exit Sub
If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, "Delete record") = vbYes Then

    With rs_level
        .Delete
        .Requery
    End With
    
Fill_record
End If

End Sub

'Close Form
Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub Fill_record()

rs_level.Requery
    With ListView1
        .ListItems.Clear
        
        Do While Not rs_level.EOF
            .ListItems.Add , , rs_level.Fields("YR"), 1, 1
            .ListItems(.ListItems.Count).SubItems(1) = Format$(IIf(IsNull(rs_level("Amount")), "", rs_level("Amount")), "#,##0.00")
            rs_level.MoveNext
        Loop
        If rs_level.RecordCount > 0 Then rs_level.AbsolutePosition = ListView1.SelectedItem.Index
    End With

End Sub

Private Sub ListView1_Click()

    If rs_level.RecordCount = 0 Then Exit Sub
    rs_level.AbsolutePosition = ListView1.SelectedItem.Index
    
End Sub
