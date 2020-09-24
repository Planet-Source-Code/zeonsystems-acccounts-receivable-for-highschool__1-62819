VERSION 5.00
Begin VB.Form frm_ae_user 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Record"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
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
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   3525
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Admin:"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   1290
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1020
      MaxLength       =   15
      TabIndex        =   1
      Top             =   480
      Width           =   3555
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1020
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   870
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   2370
      TabIndex        =   4
      Top             =   1710
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   5
      Top             =   1710
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   4560
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   510
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   900
      Width           =   750
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   30
      X2              =   4530
      Y1              =   1620
      Y2              =   1620
   End
End
Attribute VB_Name = "frm_ae_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public New_Rec As Boolean

Private Sub Form_Load()
If New_Rec = False Then
    Text1.Text = rs_user!UserName
    Text3.Text = rs_user!Name
    If Text1 = "Administrator" Then Check1.Visible = False
    If rs_user!Admin = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End If
End Sub

'Save Record
Private Sub Command1_Click()

If is_empty(Text3) Then Exit Sub
If is_empty(Text1) Then Exit Sub
If is_empty(Text2) Then Exit Sub
If frm_ae_user.Caption = "Modify Record" Then
        
        With rs_user
            .Fields("Name") = Text3.Text
            .Fields("UserName") = Text1.Text
            .Fields("Password") = Text2.Text
            .Fields("Admin") = Check1.Value
            .Update
            .Requery
        End With
        
ElseIf frm_ae_user.Caption = "New Record" Then
    rs_user.Find "UserName='" & Text1.Text & "'"
    If Not rs_user.EOF Then
        MsgBox "The username you just typed is existing. Type another username.", vbInformation
        Call highlight_focus(Text1)
        Exit Sub
    End If
    With rs_user
        If New_Rec = True Then
            .AddNew
            .Fields("Name") = Text3.Text
            .Fields("UserName") = Text1.Text
            .Fields("Password") = Text2.Text
            .Fields("Admin") = Check1.Value
            .Update
            .Requery
        End If
    End With
End If
frmUser.Fill_record

If New_Rec = True Then
    MsgBox "New record has been successfully saved.", vbInformation
Else
    MsgBox "Changes to record has been successfully saved.", vbInformation
End If
Unload Me
End Sub

'Close Window
Private Sub Command2_Click()
    Unload Me
End Sub
