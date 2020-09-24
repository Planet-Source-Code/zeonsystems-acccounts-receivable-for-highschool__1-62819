VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmschoolyear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Current School Year"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3300
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
   ScaleHeight     =   1095
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   210
      Width           =   795
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   2790
      TabIndex        =   4
      Top             =   210
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2000
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text1"
      BuddyDispid     =   196612
      OrigLeft        =   3030
      OrigTop         =   210
      OrigRight       =   3285
      OrigBottom      =   495
      Max             =   9999
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   210
      Width           =   630
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1890
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1740
      X2              =   1845
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "School Year:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frmschoolyear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim currentSY As Single

Private Sub Form_Load()
    rs_school_yr.Open "Select * From tblSchoolYear", CN, adOpenKeyset, adLockOptimistic
    Text1.Text = Mid(rs_school_yr.Fields("SY"), 1, 4)
    Text2 = Right(rs_school_yr.Fields("SY"), 4)
End Sub


Private Sub Command1_Click()
Dim cfirst As Integer
Dim csecond As Integer
Dim nfirst As Integer
Dim nsecond As Integer

If is_empty(Text1) Then Exit Sub
currentSY = Val(Text1.Text)
If Mid(Text1, 1, 4) < Mid(rs_school_yr.Fields("SY"), 1, 4) Then MsgBox "Cannot change school year to previous year", vbExclamation: Exit Sub

If MsgBox("Modifying the current school year will affect the student payment." & vbCrLf _
    & "If school year has ended you can change the current school year to the next level." & vbCrLf & vbCrLf _
    & "Do you want to save the changes you made?", vbQuestion + vbYesNo) = vbYes Then
    
    rs_school_yr.Fields("SY") = Text1.Text & "-" & Text2.Text
    rs_school_yr.Update
    rs_school_yr.Requery
    SysSY.SY = Text1.Text & "-" & Text2.Text
    MsgBox "Update Successful.", vbInformation: Unload Me
    MDIForm1.Caption = "FVHS Accounts Receivable" & " - " & SysSY.SY
    Call UnEnrolled
End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_school_yr = Nothing
End Sub

Private Sub Text1_Change()
    Text2 = Text1.Text + 1
End Sub

Private Sub Text1_GotFocus()
    Call highlight_focus(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Command1_Click
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), vbKeyBack
        Case Asc("-"), vbKeyBack
        Case Else
            KeyAscii = 0
    End Select
End Sub


