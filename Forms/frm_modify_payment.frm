VERSION 5.00
Begin VB.Form frm_modify_payment 
   Caption         =   "Modify Payment Transaction"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Text            =   " "
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   4
      Text            =   " "
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   " "
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ORNO.:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frm_modify_payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    With rs_payment
        .Fields("AccountNo") = frmPayment.AcctNo
        .Fields("Description") = Combo1.Text
        .Fields("Amount") = Text1.Text
        '.Fields("Date") = Date
        .Update
        .Requery
    End With
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
