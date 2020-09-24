VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4395
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
   ScaleHeight     =   1485
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3030
      TabIndex        =   3
      Top             =   1020
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1020
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1290
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   150
      Width           =   2865
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1290
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Old Password"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "New Password:"
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   630
      Width           =   1110
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_change_password As New ADODB.Recordset

Private Sub Command2_Click()

If is_empty(Text2) = True Then Exit Sub

If Text1 <> rs_change_password.Fields("Password") Then
    MsgBox "Incorrect Password", vbExclamation
    Call highlight_focus(Text1)
    Exit Sub

Else
    With rs_change_password
        .Fields("Password") = Text2.Text
        .Update
        .Requery
    End With

MsgBox "Password has been changed.", vbInformation
End If

Unload Me

End Sub

Private Sub Form_Load()
    rs_change_password.Open "Select * From tblUsers Where UserName ='" & SysUser.UN & "'", CN, adOpenKeyset, adLockOptimistic
    
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set rs_change_password = Nothing
End Sub
