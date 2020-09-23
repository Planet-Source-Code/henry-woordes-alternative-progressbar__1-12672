VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Alternative progressbar"
   ClientHeight    =   1140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   1140
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Decrease by 10%"
      Height          =   390
      Left            =   2175
      TabIndex        =   2
      Top             =   600
      Width           =   1590
   End
   Begin VB.PictureBox picProgressbar 
      Height          =   240
      Left            =   225
      ScaleHeight     =   180
      ScaleWidth      =   3480
      TabIndex        =   1
      Top             =   150
      Width           =   3540
      Begin VB.PictureBox picPgb 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   915
         TabIndex        =   3
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Increase by 10%"
      Height          =   390
      Left            =   225
      TabIndex        =   0
      Top             =   600
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pgbvalue As Integer


Private Sub Command1_Click()

    pgbvalue = pgbvalue + 10
    If pgbvalue > 100 Then pgbvalue = 100
    Progressbar picProgressbar, picPgb, pgbvalue

End Sub

Private Sub Command2_Click()
    
    pgbvalue = pgbvalue - 10
    If pgbvalue < 0 Then pgbvalue = 0
    Progressbar picProgressbar, picPgb, pgbvalue

End Sub

Private Sub Form_Load()

    Progressbar picProgressbar, picPgb, 0

End Sub
