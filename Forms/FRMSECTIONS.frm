VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMSECTIONS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sections"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
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
   ScaleHeight     =   4230
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&New Record"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   1290
      TabIndex        =   1
      Top             =   3720
      Width           =   885
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2250
      TabIndex        =   2
      Top             =   3720
      Width           =   885
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3210
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2835
      Left            =   90
      TabIndex        =   4
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5001
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "Imagelist1"
      SmallIcons      =   "Imagelist1"
      ForeColor       =   4194304
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sections"
         Object.Width           =   4762
      EndProperty
   End
   Begin MSComctlLib.ImageList Imagelist1 
      Left            =   1170
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMSECTIONS.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sections Records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "This form show the list of sections on years"
      Height          =   195
      Left            =   420
      TabIndex        =   5
      Top             =   330
      Width           =   3105
   End
End
Attribute VB_Name = "FRMSECTIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    With frm_ae_sections
        .new_rec = True
        .Show vbModal
    End With
End Sub

Private Sub Command2_Click()
    If rs_sections.RecordCount < 1 Then MsgBox "No level in the list.Please check it!", vbExclamation: Exit Sub
    
        With frm_ae_sections
            .new_rec = False
            .Caption = "Modify Record"
            .Show vbModal
        End With

End Sub

Private Sub Command3_Click()
If rs_sections.RecordCount < 1 Then MsgBox "No year level in the list.Please check it!", vbExclamation: Exit Sub
If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, "Delete record") = vbYes Then

    With rs_sections
        .Delete
        .Requery
    End With
    
Fill_record
End If

End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    rs_sections.Open "Select * From tblSections Where YR ='" & rs_level.AbsolutePosition & "'", CN, adOpenKeyset, adLockOptimistic
    Fill_record
End Sub

Public Sub Fill_record()

rs_sections.Requery
    With ListView1
        .ListItems.Clear
        
        Do While Not rs_sections.EOF
            .ListItems.Add , , rs_sections.Fields("Section"), 1, 1
            rs_sections.MoveNext
        Loop
        If rs_sections.RecordCount > 0 Then rs_sections.AbsolutePosition = ListView1.SelectedItem.Index
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_sections = Nothing
End Sub

Private Sub ListView1_Click()
    If rs_sections.RecordCount = 0 Then Exit Sub
    rs_sections.AbsolutePosition = ListView1.SelectedItem.Index

End Sub
