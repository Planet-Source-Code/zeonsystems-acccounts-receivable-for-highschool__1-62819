VERSION 5.00
Begin VB.Form frm_stud_list 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Students List"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
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
   ScaleHeight     =   2835
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3540
      TabIndex        =   1
      Top             =   2340
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   2340
      Width           =   945
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selection"
      Enabled         =   0   'False
      Height          =   1365
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   4365
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1020
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frm_stud_list.frx":0000
         Left            =   1020
         List            =   "frm_stud_list.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter text:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   750
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter by:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Print by selection"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   570
      Width           =   2265
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Print all"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   210
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frm_stud_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_stud_list As New ADODB.Recordset

Private Sub Command1_Click()

If Option1.Value = True Then
    Set rs_stud_list = Nothing
    rs_stud_list.Open "Select * From qryStudent ORder By YR, LastName Asc", CN, adOpenStatic, adLockReadOnly
    If rs_stud_list.EOF Then Call Message: Exit Sub
    rs_stud_list.Filter = adFilterNone
    Text2 = rs_stud_list.RecordCount
    Set Dtr_Students.DataSource = rs_stud_list
    Dtr_Students.Show vbModal
ElseIf Option2.Value = True Then
    Set rs_stud_list = Nothing
    If is_empty(Combo1) = True Then Exit Sub
    If is_empty(Text1) = True Then Exit Sub
    rs_stud_list.Open "Select * From qrystudent Where " & Combo1.Text & " Like '" & Text1.Text & "%' ORder By YR, LastName Asc", CN, adOpenStatic, adLockReadOnly
    If rs_stud_list.EOF Then Call Message: Exit Sub
    Text2 = rs_stud_list.RecordCount
    Set Dtr_Students.DataSource = rs_stud_list
    Dtr_Students.Show vbModal
End If
Exit Sub
End Sub

Private Sub Option1_Click()
    Frame1.Enabled = False
End Sub

Private Sub Option2_Click()
    Frame1.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Text1_GotFocus()
    Call highlight_focus(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Command1_Click
End Sub

