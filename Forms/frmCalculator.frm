VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clos&e"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Compute"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   " "
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   " "
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   " "
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Amount to be paid:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Number of months left:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Balance:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   630
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuCompute 
         Caption         =   "&Compute"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Clos&e"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Balance As Currency

Private Sub Command1_Click()
On Error GoTo Err
    If Val(Text2) > 12 Then MsgBox "You are not allowed to enter greater than 12.", vbInformation: Exit Sub
    Balance = frmPayment.Text5 / Text2.Text
    Text3.Text = "P " & Format$(Balance, "#,##0.00")
    Exit Sub
   
Err:
    Call highlight_focus(Text2)
    MsgBox "No divisor or zero divisor has been entered.", vbInformation
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Balance = frmPayment.Text5
    Text1.Text = "P " & Format$(Balance, "#,##0.00")
    Text2.MaxLength = 3
End Sub

Private Sub mnuClose_Click()
    Command2_Click
End Sub

Private Sub mnuCompute_Click()
    Command1_Click
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1_Click
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Else: KeyAscii = 0
    End Select
End Sub

