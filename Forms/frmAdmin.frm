VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Setup"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
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
   ScaleHeight     =   1275
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   810
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   810
      Width           =   1185
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1050
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   390
      Width           =   3015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   420
      Width           =   750
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_new_user As New ADODB.Recordset

Private Sub Command2_Click()
If is_empty(Text1) Then Exit Sub
With rs_new_user
    .AddNew
    .Fields("UserName") = "Administrator"
    .Fields("Admin") = True
    .Fields("Password") = Text1.Text
    .Update
    .Requery
End With
frmLogin.Show vbModal
Unload Me
End Sub

Private Sub Form_Load()
rs_new_user.Open "Select * From tblUsers", CN, adOpenKeyset, adLockOptimistic

End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_new_user = Nothing
End Sub
