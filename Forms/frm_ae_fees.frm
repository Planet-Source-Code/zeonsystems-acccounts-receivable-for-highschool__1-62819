VERSION 5.00
Begin VB.Form frm_ae_fees 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Record"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
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
   ScaleHeight     =   1425
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Standard:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3570
      TabIndex        =   3
      Top             =   960
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3030
      TabIndex        =   1
      Top             =   540
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1110
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   3285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Left            =   2370
      TabIndex        =   5
      Top             =   570
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fee Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frm_ae_fees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public New_Rec As Boolean

Private Sub Form_Load()

If New_Rec = False Then
    Text1.Text = rs_fees!FeeName
    Text2.Text = Format$(rs_fees!Amount, "#,##0.00")
    If rs_fees!Standard = True Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
End If

End Sub

'Update fees
Private Sub Command1_Click()

If is_empty(Text1) Then Exit Sub
If is_empty(Text2) Then Exit Sub
If not_curr(Text2) Then Exit Sub

If frm_ae_fees.Caption = "Modify Record" Then
    
   With rs_fees
            .Fields("FeeName") = Text1.Text
            .Fields("Amount") = Text2.Text
            .Fields("Standard") = Check1.Value
            .Update
            .Requery
       End With
    
ElseIf frm_ae_fees.Caption = "New Record" Then
    
    rs_fees.Requery
    rs_fees.Find "FeeName='" & Text1.Text & "'"
    
    If Not rs_fees.EOF Then
        MsgBox "The fee name you just entered is already existing.", vbInformation
        Text1.SetFocus
        Call highlight_focus(Text1)
        Exit Sub
    End If
    With rs_fees
        If New_Rec = True Then .AddNew
            .Fields("FeeName") = Text1.Text
            .Fields("Amount") = Text2.Text
            .Fields("Standard") = Check1.Value
            .Update
            .Requery
    End With
    End If
frmFee.Fill_record

If New_Rec = True Then
    MsgBox "New record has been successfully saved.", vbInformation
Else
    MsgBox "Changes to record has been successfully saved.", vbInformation
End If
Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc("."), vbKeyBack
        Case Else: KeyAscii = 0
    End Select
End Sub
