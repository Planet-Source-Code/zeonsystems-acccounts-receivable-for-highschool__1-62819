VERSION 5.00
Begin VB.Form frmSearchStudents 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Text            =   " "
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSearchStudents.frx":0000
      Left            =   1320
      List            =   "frmSearchStudents.frx":0016
      TabIndex        =   1
      Text            =   "ID"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "View students by:"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1245
   End
End
Attribute VB_Name = "frmSearchStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs_student As New ADODB.Recordset
    
Private Sub Command1_Click()
    If is_empty(Combo1) Then Exit Sub
    rs_student.Requery
    Call StudentList
    End Sub

Private Sub Form_Load()
    rs_student.Open "Select * From tblStudent ", CN, adOpenKeyset, adLockOptimistic
    'Where AccountNo = '" & frmPayment.AcctNo & "'
    'Fill_record
    'Combo1.DataSource = rs_student
End Sub
