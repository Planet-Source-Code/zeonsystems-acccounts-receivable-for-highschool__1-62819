VERSION 5.00
Begin VB.Form frm_ae_discount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Record"
   ClientHeight    =   1500
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
   ScaleHeight     =   1500
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   3
      Top             =   930
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   930
      Width           =   1065
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   510
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Discount Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Left            =   2100
      TabIndex        =   4
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "frm_ae_discount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public New_Rec As Boolean

Private Sub Form_Load()
If New_Rec = False Then
    Text1.Text = rs_discount!DiscountName
    Text2.Text = Format$(rs_discount!Amount, "#,##0.00")
End If
End Sub

'Update discount
Private Sub Command1_Click()

If is_empty(Text1) Then Exit Sub
If is_empty(Text2) Then Exit Sub
If not_curr(Text2) Then Exit Sub
If frm_ae_discount.Caption = "Modify Record" Then
    With rs_discount
        If New_Rec = False Then
            .Fields("DiscountName") = Text1.Text
            .Fields("Amount") = Text2.Text
            .Update
            .Requery
        End If
    End With
    
ElseIf frm_ae_discount.Caption = "New Record" Then
    
    rs_discount.Requery
    rs_discount.Find "DiscountName='" & Text1.Text & "'"
    
    If Not rs_discount.EOF Then
        MsgBox "The Discount Name you just entered is already existing.", vbInformation
        Text1.SetFocus
        Call highlight_focus(Text1)
        Exit Sub
    End If
    With rs_discount
        If New_Rec = True Then .AddNew
               .Fields("DiscountName") = Text1.Text
            .Fields("Amount") = Text2.Text
            .Update
            .Requery
        End With
End If
frmDiscount.Fill_record

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
