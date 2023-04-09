VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H000040C0&
   Caption         =   "SJB Quiz Bee"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdquit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "PLAY"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(AMPID CAMPUS)"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   7440
      Picture         =   "frm1.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   1365
      Left            =   720
      Picture         =   "frm1.frx":17BB42
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUIZ BEE 2018"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   65.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   9375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Saint John Bosco I.A.S"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "Frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdplay_Click()
Dim frm1 As New frm2
frm2.Show
Me.Hide
End Sub

Private Sub cmdquit_Click()
End
End Sub
