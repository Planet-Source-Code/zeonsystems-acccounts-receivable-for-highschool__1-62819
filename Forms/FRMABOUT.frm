VERSION 5.00
Begin VB.Form FRMABOUT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
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
   ScaleHeight     =   3540
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2835
      Left            =   30
      ScaleHeight     =   2775
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   180
      Width           =   855
      Begin VB.Image Image1 
         Height          =   720
         Left            =   60
         Picture         =   "FRMABOUT.frx":0000
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3060
      TabIndex        =   10
      Top             =   3090
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Members"
      Height          =   1935
      Left            =   930
      TabIndex        =   3
      Top             =   1080
      Width           =   3285
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Andro Simanero"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Andrew Aniceto"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   1290
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Evance Paul Javier"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   1050
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Ian Callanga"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   810
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Michael Kim Cuti"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marvin Pablo"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   390
         Width           =   915
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2005, All Rights Reserved."
      Height          =   195
      Left            =   1020
      TabIndex        =   2
      Top             =   870
      Width           =   2865
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Windows 32-bit Application"
      Height          =   195
      Left            =   990
      TabIndex        =   1
      Top             =   450
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts Receivable System"
      Height          =   195
      Left            =   990
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "FRMABOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

