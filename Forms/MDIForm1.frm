VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Accounts Receivable ver. 2.0 - School Year "
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   9345
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":5A0A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "Imagelist1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Student Records"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Year Level"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fees"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Discount"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "User Accounts"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Payment Transaction"
            ImageIndex      =   8
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Student List"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Financial Report"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "List of students w/ balance"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "List of paid students"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Discount Report"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Statement of Account"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Transaction Report"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   240
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10FFC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1106F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":110A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":110E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1111C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":111560
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1118FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":111C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11202E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1123C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":112762
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":112AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":112E96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPaymentForm 
         Caption         =   "&PaymentTransaction"
      End
      Begin VB.Menu Divider 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log &Out"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View Records"
      Begin VB.Menu mnuStudent 
         Caption         =   "&Student "
      End
      Begin VB.Menu mnuLevel 
         Caption         =   "Year &Levels"
      End
      Begin VB.Menu mnuFees 
         Caption         =   "&Fees"
      End
      Begin VB.Menu mnuDiscount 
         Caption         =   "&Discounts"
      End
      Begin VB.Menu mnuSchoolYear 
         Caption         =   "School &Year"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuStuddentList 
         Caption         =   "&Student List"
      End
      Begin VB.Menu mnuFinancialReport 
         Caption         =   "&Financial Report"
      End
      Begin VB.Menu mnuStudWithBalance 
         Caption         =   "List of Students w/ &Balance"
      End
      Begin VB.Menu mnuPaid 
         Caption         =   "List of &Paid Students"
      End
      Begin VB.Menu mnuDiscountReport 
         Caption         =   "&Discount Report"
      End
      Begin VB.Menu mnuStatementofAccount 
         Caption         =   "Statement of &Account"
      End
      Begin VB.Menu mnuTransactionReport 
         Caption         =   "&Transaction Report"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Admin"
      Begin VB.Menu mnuUsers 
         Caption         =   "&Users"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
    Call ConnectDB
    Dim rs_user As New ADODB.Recordset
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If MsgBox("Are you sure you want to exit?", vbExclamation + vbYesNo, "Exit system") = vbYes Then
    End:
Else
    Cancel = 1
End If
End Sub

Private Sub mnuAbout_Click()
    FRMABOUT.Show vbModal
End Sub

Private Sub mnuChangePassword_Click()
    frmChange.Show vbModal
End Sub



Private Sub mnuDiscountReport_Click()
    frm_discount.Show vbModal
End Sub

Private Sub mnuFinancialReport_Click()
    Call FinancialReport
End Sub

Private Sub mnuLogOut_Click()
    With frmLogin
        .Text1.Text = ""
        .Text2.Text = ""
        .Command3.Visible = True
        .Show vbModal
    End With
End Sub

Private Sub mnuPaid_Click()
    frm_paid_stud.Show vbModal
End Sub

Private Sub mnuSchoolYear_Click()
    If SysUser.UA = False Then MsgBox "Your account has been denied to access this feature. Please contact your administrator for this instances.", vbExclamation: Exit Sub
    frmschoolyear.Show vbModal
End Sub

Private Sub mnuStatementofAccount_Click()
    frm_statement_list.Show vbModal
End Sub

Private Sub mnuStuddentList_Click()
    frm_stud_list.Show vbModal
End Sub

Private Sub mnuStudent_Click()
    frmStudent.Show
    frmStudent.SetFocus
    frmStudent.WindowState = 0
End Sub

Private Sub mnuLevel_Click()
    frmYear.Show
    frmYear.SetFocus
    frmYear.WindowState = 0
End Sub

Private Sub mnuFees_Click()
    frmFee.Show
    frmFee.SetFocus
    frmFee.WindowState = 0
End Sub

Private Sub mnuDiscount_Click()
    frmDiscount.Show
    frmDiscount.SetFocus
    frmDiscount.WindowState = 0
End Sub

Private Sub mnuPaymentForm_Click()
        frmPayment.Show vbModal
End Sub

Private Sub mnuStudWithBalance_Click()
    frm_unpaid_stud.Show vbModal
End Sub

Private Sub mnuTransactionReport_Click()
    frm_transaction_report.Show 1
End Sub

Private Sub mnuUsers_Click()

If SysUser.UA = False Then MsgBox "Your account has been denied to access this feature. Please contact your administrator for this instances.", vbExclamation: Exit Sub
    With frmUser
        .Show
        .SetFocus
        .WindowState = 0
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index

Case 2  'Student Record Form
    mnuStudent_Click
Case 3  'Year Level Form
    mnuLevel_Click
Case 4  'Fees Form
    mnuFees_Click
Case 5  'Discount Form
    mnuDiscount_Click
Case 7  'User Accounts Form
    mnuUsers_Click
Case 9 'Payment Form
    mnuPaymentForm_Click

End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Text
'List of students
Case "Student List"
    frm_stud_list.Show vbModal
'Finacial report
Case "Financial Report"
    Call FinancialReport
'Student with balance
Case "List of students w/ balance"
   mnuStudWithBalance_Click
'List of paind students
Case "List of paid students"
    Call mnuPaid_Click
Case "Discount Report"
    mnuDiscountReport_Click
Case "Transaction Report"
    Call mnuTransactionReport_Click
Case "Statement of Account"
    mnuStatementofAccount_Click
End Select
End Sub
'list of First Year students
Private Sub StudentsList()
Dim rs_stud_list As New ADODB.Recordset
rs_stud_list.Open "Select * From qryStudent ORDER By LastName Asc", CN, adOpenStatic, adLockReadOnly
    Set Dtr_Students.DataSource = rs_stud_list
    Dtr_Students.Show vbModal
    Set rs_stud = Nothing
End Sub
'Financial report
Private Sub FinancialReport()
    frmFinancialReport.Show
    frmFinancialReport.SetFocus
    frmFinancialReport.WindowState = 0
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

