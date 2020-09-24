VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Accounts"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
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
   ScaleHeight     =   4155
   ScaleWidth      =   4515
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3540
      TabIndex        =   3
      Top             =   3660
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   3660
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   1350
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
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2835
      Left            =   90
      TabIndex        =   4
      Top             =   720
      Width           =   4275
      _ExtentX        =   7541
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
         Text            =   "User Name"
         Object.Width           =   4516
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2822
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   3720
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
            Picture         =   "frmUser.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "This form show all the user of the system."
      Height          =   195
      Left            =   420
      TabIndex        =   6
      Top             =   390
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Records"
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
      Width           =   1125
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    rs_user.Open "Select * From tblUsers Order By UserName Desc", CN, adOpenKeyset, adLockOptimistic
    Fill_record
    frmUser.Left = MDIForm1.ScaleWidth / 2 - frmUser.ScaleWidth / 2
    frmUser.Top = MDIForm1.ScaleHeight / 2 - frmUser.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_user = Nothing
End Sub
'New Record
Private Sub Command1_Click()
With frm_ae_user
    .New_Rec = True
    .Show vbModal
End With
End Sub

'Modify Record
Private Sub Command2_Click()

If rs_user.RecordCount < 1 Then MsgBox "No user in the list. Please check it!", vbExclamation: Exit Sub

With frm_ae_user
    .New_Rec = False
    .Caption = "Modify Record"
    .Show vbModal
End With

End Sub

'Delete Record
Private Sub Command3_Click()

If rs_user.Fields("UserName") = "Administrator" Then MsgBox "You are not allowed to delete this administrator record.", vbExclamation: Exit Sub

If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, "Delete record") = vbYes Then

    With rs_user
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

rs_user.Requery
With ListView1
    .ListItems.Clear
    
    Do While Not rs_user.EOF
        .ListItems.Add , , rs_user.Fields("UserName"), 1, 1
        .ListItems(.ListItems.Count).SubItems(1) = IIf(IsNull(rs_user!Name), "", rs_user!Name)
        '.ListItems(.ListItems.Count).SubItems(2) = IIf(IsNull(""), "", " ")
        rs_user.MoveNext
    Loop
    If rs_user.RecordCount > 0 Then rs_user.AbsolutePosition = ListView1.SelectedItem.Index

End With
End Sub

Private Sub ListView1_Click()

    If rs_user.RecordCount = 0 Then Exit Sub
    rs_user.AbsolutePosition = ListView1.SelectedItem.Index
    
End Sub
