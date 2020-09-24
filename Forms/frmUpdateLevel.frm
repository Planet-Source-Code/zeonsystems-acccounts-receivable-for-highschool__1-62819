VERSION 5.00
Begin VB.Form frmUpdateLevel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Level"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3345
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
   ScaleHeight     =   1110
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1380
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2190
      TabIndex        =   3
      Top             =   600
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   1000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1110
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   2115
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year Level:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmUpdateLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Changed As Boolean
Dim rsStud As New ADODB.Recordset

Private Sub Combo1_Click()
   ' Call fill_sec_yr(Combo1, Combo2)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call fill_yr(Combo1)
    rsStud.Open "Select * From tblStudent", CN, adOpenKeyset, adLockOptimistic
    rsStud.Filter = "ID='" & frmPayment.rs_id!ID & "'"
    Exit Sub
End Sub

Private Sub Command1_Click()
If Combo1.Text < frmPayment.Label15.Caption Then MsgBox "Cannot update to lower year level", vbExclamation: Exit Sub
If Combo1.Text - frmPayment.Label15.Caption >= 2 Then MsgBox "Cannot update to higher year level", vbExclamation: Exit Sub
rsStud.Fields("YR") = Combo1.Text
'rsStud.Fields("Section") = Combo2.Text
If Combo1.Text = "" Then MsgBox "Fill the year level.", vbInformation: Exit Sub
'If Combo2.Text = "" Then MsgBox "Fill in the section.", vbInformation: Exit Sub
rsStud.Update
rsStud.Requery
MsgBox "Year level has been updated.", vbInformation
With frmPayment.rs_acct
    .Fields("Status") = False
    .Update
    .Requery
End With
'frmPayment.get_acct
Unload Me
frmPayment.change_rec
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsStud = Nothing
End Sub

Private Sub Command2_Click()
    rsStud.Update
    rsStud.Requery
    With frmPayment.rs_acct
        .Fields("Status") = False
        .Update
        .Requery
    End With
    frmPayment.change_rec

    Unload Me
End Sub

