VERSION 5.00
Begin VB.Form frm_ae_sections 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Record"
   ClientHeight    =   1185
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
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3450
      TabIndex        =   3
      Top             =   660
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   2
      Top             =   660
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   270
      Width           =   1035
   End
End
Attribute VB_Name = "frm_ae_sections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public new_rec As Boolean

Private Sub Form_Load()
On Error Resume Next
If new_rec = False Then
    Text1.Text = rs_sections!Section
End If

End Sub

Private Sub Command1_Click()
If is_empty(Text1) Then Exit Sub
With rs_sections
    If new_rec = True Then .AddNew
    .Fields("YR") = rs_level.AbsolutePosition
    .Fields("Section") = Text1.Text
    .Update
    .Requery
End With
FRMSECTIONS.Fill_record
If new_rec = True Then
    MsgBox "New record has been successfully saved.", vbInformation, AppTitle
Else
    MsgBox "Changes to record has been successfully saved.", vbInformation, AppTitle
End If
Unload Me

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

