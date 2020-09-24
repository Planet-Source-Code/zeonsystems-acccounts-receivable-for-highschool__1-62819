VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStudent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Record"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   10110
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmStudent.frx":0000
      Height          =   4335
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7646
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LastName"
         Caption         =   "LastName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FirstName"
         Caption         =   "FirstName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "MI"
         Caption         =   "MI"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Gender"
         Caption         =   "Gender"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "YR"
         Caption         =   "YR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "ContactNo"
         Caption         =   "ContactNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Enrolled"
         Caption         =   "Enrolled"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   "0.000E+00"
            HaveTrueFalseNull=   1
            TrueValue       =   "Yes"
            FalseValue      =   "No"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   5
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Search"
      Height          =   345
      Left            =   4440
      TabIndex        =   4
      Top             =   270
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmStudent.frx":0015
      Left            =   1110
      List            =   "frmStudent.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2610
      TabIndex        =   6
      Top             =   270
      Width           =   1725
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Close"
      Height          =   375
      Left            =   8580
      TabIndex        =   3
      Top             =   5250
      Width           =   1295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   7170
      TabIndex        =   2
      Top             =   5250
      Width           =   1295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New Record"
      Default         =   -1  'True
      Height          =   375
      Left            =   5820
      TabIndex        =   1
      Top             =   5250
      Width           =   1295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search by:"
      Height          =   195
      Left            =   210
      TabIndex        =   5
      Top             =   330
      Width           =   780
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    rs_stud.Open "Select * From tblStudent Order By ID Desc", CN, adOpenKeyset, adLockPessimistic
    Set DataGrid1.DataSource = rs_stud
    Combo1.ListIndex = 0
    frmStudent.Left = MDIForm1.ScaleWidth / 2 - frmStudent.ScaleWidth / 2
    frmStudent.Top = MDIForm1.ScaleHeight / 2 - frmStudent.ScaleHeight / 2
    DataGrid1.AllowRowSizing = False
    DataGrid1.Columns(7).Width = 4500
End Sub

Private Sub Form_Unload(Cancel As Integer)

Set rs_stud = Nothing

End Sub

'Search text
Private Sub Command6_Click()
If is_empty(Combo1) = True Then Exit Sub
If is_empty(Text1) = True Then Exit Sub
rs_stud.Requery
If Combo1.ListIndex = 0 Then
    rs_stud.Find "ID='" & Text1.Text & "'"
    If rs_stud.EOF Then
        MsgBox "No record found", vbExclamation
        rs_stud.Requery
    End If
Else
    rs_stud.Find "LastName='" & Text1.Text & "'"
    If rs_stud.EOF Then
        MsgBox "No record found", vbExclamation
        rs_stud.Requery
    End If
End If
End Sub

'New Record
Private Sub Command1_Click()
With frm_ae_student
    .new_rec = True
    .Show vbModal
End With
End Sub

'Modify record
Private Sub Command2_Click()
With frm_ae_student
    .new_rec = False
    .Caption = "Modify record"
    .Show vbModal
End With
End Sub

'Delete Record
Private Sub Command4_Click()
If rs_stud.RecordCount = 1 Then MsgBox "You are not allowed to delete all the records. There must be at  least one record left.", vbExclamation: Exit Sub

If MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo, "Delete record") = vbYes Then

    With rs_stud
        .Delete
       ' .Requery
    End With

End If
DataGrid1.Columns(7).Width = 4500
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub modify_Click()
    Call Command2_Click
End Sub

Private Sub new_Click()
    Call Command1_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Command6_Click
End Sub
