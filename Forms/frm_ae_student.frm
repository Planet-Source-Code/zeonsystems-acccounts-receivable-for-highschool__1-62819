VERSION 5.00
Begin VB.Form frm_ae_student 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Record"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
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
   ScaleHeight     =   2625
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo3 
      ForeColor       =   &H00400000&
      Height          =   315
      ItemData        =   "frm_ae_student.frx":0000
      Left            =   3900
      List            =   "frm_ae_student.frx":0002
      TabIndex        =   5
      Top             =   3060
      Width           =   1785
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_ae_student.frx":0004
      Left            =   1110
      List            =   "frm_ae_student.frx":000E
      TabIndex        =   6
      Top             =   1050
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5610
      MaxLength       =   13
      TabIndex        =   7
      Text            =   " NA"
      Top             =   1050
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5700
      TabIndex        =   10
      Top             =   2130
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   4230
      TabIndex        =   9
      Top             =   2130
      Width           =   1365
   End
   Begin VB.ComboBox Combo2 
      ForeColor       =   &H00400000&
      Height          =   315
      ItemData        =   "frm_ae_student.frx":0020
      Left            =   2550
      List            =   "frm_ae_student.frx":0022
      TabIndex        =   4
      Top             =   1080
      Width           =   1605
   End
   Begin VB.TextBox Text5 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1080
      MaxLength       =   60
      TabIndex        =   8
      Top             =   1530
      Width           =   5925
   End
   Begin VB.TextBox Text4 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6150
      MaxLength       =   2
      TabIndex        =   3
      Top             =   630
      Width           =   915
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   3900
      MaxLength       =   20
      TabIndex        =   2
      Top             =   630
      Width           =   1785
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
      Width           =   1785
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   0
      Top             =   210
      Width           =   1785
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section:"
      Height          =   195
      Left            =   3240
      TabIndex        =   19
      Top             =   3090
      Width           =   585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      Height          =   195
      Left            =   420
      TabIndex        =   18
      Top             =   1080
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   7170
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      Height          =   195
      Left            =   4290
      TabIndex        =   17
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yr."
      Height          =   195
      Left            =   2220
      TabIndex        =   16
      Top             =   1140
      Width           =   210
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   1530
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M.I:"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   5820
      TabIndex        =   14
      Top             =   660
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      Top             =   660
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   660
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID No.:"
      Height          =   195
      Left            =   210
      TabIndex        =   11
      Top             =   240
      Width           =   525
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   180
      X2              =   7170
      Y1              =   2010
      Y2              =   2010
   End
End
Attribute VB_Name = "frm_ae_student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public new_rec As Boolean
Dim get_year As New ADODB.Recordset

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
'Call fill_sec_yr(Combo2, Combo3)
End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call fill_yr(Combo2)
If new_rec = False Then
    Text1.Locked = True
    With rs_stud
        Text1.Text = .Fields("ID")
        Text2.Text = .Fields("LastName")
        Text3.Text = .Fields("FirstName")
        Text4.Text = .Fields("MI")
        Text5.Text = .Fields("Address")
        Text8.Text = .Fields("ContactNo")
        Combo1.Text = .Fields("Gender")
        Combo2.Text = .Fields("YR")
    End With
Else
    Combo2.ListIndex = 1
    End If
End Sub

'Save fields
Private Sub Command1_Click()

If Len(Text8) >= 10 Then
    Text8.Text = Format$(Text8, "0###-###-####")
ElseIf Len(Text8) = 9 Then
    Text8.Text = Format$(Text8, "0##-###-####")
End If
If new_rec = True Then
    If check_id(Text1) Then Exit Sub
End If
    
If is_empty(Text1) = True Then Exit Sub
If is_empty(Text2) = True Then Exit Sub
If is_empty(Text3) = True Then Exit Sub
If is_empty(Text4) = True Then Exit Sub
If is_empty(Text5) = True Then Exit Sub
If is_empty(Combo1) = True Then Exit Sub
If is_empty(Combo2) = True Then Exit Sub
'If is_empty(Combo3) = True Then Exit Sub
If is_empty(Text8) = True Then Exit Sub

If new_rec = False Then
    If Combo2.Text < rs_stud.Fields("YR") Then MsgBox "Cannot update to lower year level", vbExclamation: Exit Sub
    If Combo2.Text - rs_stud.Fields("YR") >= 2 Then MsgBox "Cannot update to higher year level", vbExclamation: Exit Sub
End If
Call ToUpper(Text1)
Call ToUpper(Text2)
Call ToUpper(Text3)
Call ToUpper(Text4)
Call ToUpper(Text5)
With rs_stud
    If new_rec = True Then .AddNew
    If new_rec = True Then .Fields("ID") = Text1.Text
        .Fields("LastName") = Text2.Text
        .Fields("FirstName") = Text3.Text
        .Fields("MI") = Text4.Text
        .Fields("Address") = Text5.Text
        .Fields("ContactNo") = Text8.Text
        .Fields("Gender") = Combo1.Text
        .Fields("YR") = Combo2.Text
        '.Fields("Section") = Combo3.Text
        .Update
    
End With
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

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case Asc("0") To Asc("9"), vbKeyBack
Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("A") To Asc("Z"), vbKeyBack
    Case Asc("a") To Asc("z"), vbKeyBack
    Case Asc("Ñ"), vbKeyBack
    Case Asc("ñ"), vbKeyBack
    Case Asc(" "), vbKeyBack
    Case Else: KeyAscii = 0
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("A") To Asc("Z"), vbKeyBack
    Case Asc("a") To Asc("z"), vbKeyBack
    Case Asc("Ñ"), vbKeyBack
    Case Asc("ñ"), vbKeyBack
    Case Asc(" "), vbKeyBack
    Case Else: KeyAscii = 0
End Select

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case Asc("A") To Asc("Z"), vbKeyBack
    Case Asc("a") To Asc("z"), vbKeyBack
    Case Asc("Ñ"), vbKeyBack
    Case Asc("ñ"), vbKeyBack
    Case Asc(" "), vbKeyBack
    Case Else: KeyAscii = 0
End Select

End Sub

Private Sub Text8_GotFocus()
Call highlight_focus(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)

Select Case KeyAscii

Case Asc("0") To Asc("9"), vbKeyBack
Case Else: KeyAscii = 0

End Select

End Sub

