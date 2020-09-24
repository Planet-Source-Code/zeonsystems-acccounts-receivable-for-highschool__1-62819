VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
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
   ScaleHeight     =   2010
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2850
      TabIndex        =   7
      Top             =   1500
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   1620
      TabIndex        =   5
      Top             =   1500
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2850
      TabIndex        =   4
      Top             =   1500
      Width           =   1125
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
      Left            =   1140
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Admin"
      Top             =   1080
      Width           =   2865
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
      Left            =   1140
      TabIndex        =   2
      Text            =   "Admin"
      Top             =   630
      Width           =   2865
   End
   Begin VB.Label Label3 
      Caption         =   "Please enter your username and password to login."
      Height          =   405
      Left            =   1020
      TabIndex        =   6
      Top             =   120
      Width           =   3225
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "frmLogin.frx":0000
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   1110
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   150
      TabIndex        =   0
      Top             =   660
      Width           =   840
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs_login As New ADODB.Recordset
Dim rs_sy As New ADODB.Recordset
Dim Tries As Integer

Private Sub Form_Load()
    rs_login.Open "Select * From tblUsers", CN, adOpenStatic, adLockReadOnly
    rs_sy.Open "Select * From tblSchoolYear", CN, adOpenStatic, adLockReadOnly
    If rs_login.RecordCount <= 0 Then
        MsgBox "Welcome to FVHS Accounts Receivable. Please enter the administrator password to use the system.", vbOKOnly + vbInformation, "Welcome to "
        Unload Me
        frmAdmin.Show vbModal
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_login = Nothing
    Set rs_sy = Nothing
End Sub


Private Sub Command2_Click()

rs_login.Requery
rs_login.Find "UserName='" & Text1.Text & "'"
If Tries > 2 Then
    Command2.Enabled = False
    MsgBox "Contact the administrator for this instance...", vbExclamation
    Exit Sub
End If

If rs_login.EOF Then
    MsgBox "Password denied.", vbExclamation
    Call Denied
    Exit Sub
Else
    If Text2 <> rs_login!Password Then
        MsgBox "Password denied.", vbExclamation
        Call highlight_focus(Text2)
        Call Denied
        Exit Sub
    Else
        With SysUser
            .UN = Text1.Text
            .UP = Text2.Text
            .UA = rs_login!Admin
        End With
        frmAddPayment.Text2 = Text1
        SysSY.SY = rs_sy!SY
        MDIForm1.Caption = "FVHS Accounts Recivable" & "- " & SysSY.SY
        Me.Hide
    End If
End If

End Sub

Private Function Denied()
    Tries = Tries + 1
End Function

'Shut down system
Private Sub Command1_Click()
    End
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

