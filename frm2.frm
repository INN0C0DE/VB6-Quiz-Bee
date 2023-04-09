VERSION 5.00
Begin VB.Form frm2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Choose a Subject"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form2"
   ScaleHeight     =   6735
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FILIPINO"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   20.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CHOOSE A SUBJECT DO YOU WANT TO ANSWER:"
      BeginProperty Font 
         Name            =   "Lucida Fax"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9255
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5160
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SCIENCE"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MATH"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3240
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENGLISH"
         BeginProperty Font 
            Name            =   "Lucida Fax"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmeng.Show
frm2.Hide
End Sub

Private Sub Command2_Click()
frmfil.Show
frm2.Hide
End Sub

Private Sub Command3_Click()
frmmath.Show
frm2.Hide
End Sub

Private Sub Command4_Click()
frmsci.Show
frm2.Hide
End Sub

Private Sub Command5_Click()
Dim frm2 As New frm1
frm1.Show
Me.Hide
End Sub

