VERSION 5.00
Begin VB.Form frm_ae_level 
   Caption         =   "New Record"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   3630
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1830
      TabIndex        =   1
      Top             =   510
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   1410
      TabIndex        =   2
      Top             =   930
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   930
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   990
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Left            =   1110
      TabIndex        =   5
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Year Level:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   150
      Width           =   810
   End
End
Attribute VB_Name = "frm_ae_level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public New_Rec As Boolean

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
If New_Rec = False Then
    Text1.Text = rs_level!yr
    Text2.Text = Format$(rs_level!Amount, "#,##0.00")
End If
End Sub

'Update year level record
Private Sub Command1_Click()
If is_empty(Text1) Then Exit Sub
If is_empty(Text2) Then Exit Sub
If not_curr(Text2) Then Exit Sub

If frm_ae_level.Caption = "Modify Record" Then
    Text1.Locked = True
    With rs_level
        If New_Rec = False Then
            .Fields("YR") = Text1.Text
            .Fields("Amount") = Text2.Text
            .Update
            .Requery
        End If
    End With
    
ElseIf frm_ae_level.Caption = "New Record" Then
    
    rs_level.Requery
    rs_level.Find "YR='" & Text1.Text & "'"
    
    If Not rs_level.EOF Then
        MsgBox "The year level you just entered is already existing.", vbInformation
        Text1.SetFocus
        Call highlight_focus(Text1)
        Exit Sub
    End If
    With rs_level
        If New_Rec = True Then .AddNew
        .Fields("YR") = Text1.Text
        .Fields("Amount") = Text2.Text
        .Update
        .Requery
    End With
End If
frmYear.Fill_record
If New_Rec = True Then
    MsgBox "New record has been successfully saved.", vbInformation
Else
    MsgBox "Changes to record has been successfully saved.", vbInformation
End If

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
