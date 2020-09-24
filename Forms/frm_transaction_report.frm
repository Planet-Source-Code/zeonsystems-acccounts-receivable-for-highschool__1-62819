VERSION 5.00
Begin VB.Form frm_transaction_report 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Report"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Print all"
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   240
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Print by selection"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   600
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection"
      Enabled         =   0   'False
      Height          =   1365
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   4365
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frm_transaction_report.frx":0000
         Left            =   1050
         List            =   "frm_transaction_report.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1230
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter by:"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter text:"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   750
         Width           =   765
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2370
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   1
      Top             =   2370
      Width           =   945
   End
End
Attribute VB_Name = "frm_transaction_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If Option1.Value = True Then
    Set rs = Nothing
    rs.Open "Select * From qryTransaction Order By YR, LastName Asc", CN, adOpenStatic, adLockReadOnly
    If rs.EOF Then Call Message: Exit Sub
    Set Dtr_transaction.DataSource = rs
    Dtr_transaction.Show vbModal
ElseIf Option2.Value = True Then
    If is_empty(Combo1) Then Exit Sub
    If is_empty(Text1) Then Exit Sub
    Set rs = Nothing
    rs.Open "Select * From qryTransaction Where " & Combo1.Text & " Like '" & Text1.Text & "%' Order By YR, LastName Asc", CN, adOpenStatic, adLockReadOnly
    If rs.EOF Then Call Message: Exit Sub
    Set Dtr_transaction.DataSource = rs
    Dtr_transaction.Show vbModal
End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Option1_Click()
    Frame1.Enabled = False
End Sub

Private Sub Option2_Click()
    Frame1.Enabled = True
End Sub
Private Sub Text1_GotFocus()
    Call highlight_focus(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Command1_Click
End Sub

