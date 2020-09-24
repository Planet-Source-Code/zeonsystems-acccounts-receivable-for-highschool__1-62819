VERSION 5.00
Begin VB.Form frm_ae_addfee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Record"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
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
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_ae_addfee.frx":0000
      Left            =   1260
      List            =   "frm_ae_addfee.frx":0007
      TabIndex        =   0
      Top             =   90
      Width           =   3165
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2790
      TabIndex        =   1
      Top             =   480
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   3
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Left            =   2100
      TabIndex        =   5
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fee Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   780
   End
End
Attribute VB_Name = "frm_ae_addfee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If is_empty(Combo1) = True Then Exit Sub
If is_empty(Text2) = True Then Exit Sub
With frmAddFee.rs_new
    .AddNew
    .Fields("Accountno") = frmPayment.AcctNo
    .Fields("FeeName") = Combo1.Text
    .Fields("Amount") = Text2.Text
    .Update
    .Requery
End With
frmAddFee.Fill_record
frmPayment.fill_value
MsgBox "New record has been successfully saved.", vbInformation: Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Text2_GotFocus()
    If Len(Combo1) > 15 Then
        MsgBox "You are only allowed to enter 15 characters.", vbInformation
        Combo1.SetFocus
        SendKeys "{Home} +{End}"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc("."), vbKeyBack
        Case Else: KeyAscii = 0
    End Select
End Sub
