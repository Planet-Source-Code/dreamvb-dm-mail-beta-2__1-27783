VERSION 5.00
Begin VB.Form frmmailers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Address"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "frmmailers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdview2 
      Caption         =   "<<&Hide"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7545
      TabIndex        =   12
      Top             =   1965
      Width           =   780
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Update"
      Height          =   375
      Left            =   6615
      TabIndex        =   11
      Top             =   1965
      Width           =   900
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5595
      TabIndex        =   10
      Top             =   1455
      Width           =   2550
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5595
      TabIndex        =   8
      Top             =   1080
      Width           =   2550
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   5595
      TabIndex        =   7
      Top             =   720
      Width           =   2550
   End
   Begin VB.CommandButton cmdview1 
      Caption         =   "&View >>"
      Height          =   375
      Left            =   2970
      TabIndex        =   4
      Top             =   570
      Width           =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2970
      TabIndex        =   3
      Top             =   135
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5625
      TabIndex        =   2
      Top             =   1965
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New"
      Height          =   375
      Left            =   4740
      TabIndex        =   1
      Top             =   1965
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   30
      TabIndex        =   0
      Top             =   135
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Email"
      Height          =   195
      Left            =   4755
      TabIndex        =   9
      Top             =   1515
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Home page"
      Height          =   195
      Left            =   4740
      TabIndex        =   6
      Top             =   1095
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Full name"
      Height          =   195
      Left            =   4725
      TabIndex        =   5
      Top             =   765
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4770
      Picture         =   "frmmailers.frx":08CA
      Top             =   165
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   4620
      X2              =   4620
      Y1              =   90
      Y2              =   2595
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   4635
      X2              =   4635
      Y1              =   120
      Y2              =   2625
   End
End
Attribute VB_Name = "frmmailers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label4_Click()

End Sub

Private Sub Command4_Click()
    
End Sub

Private Sub cmdview1_Click()
    frmmailers.Width = 8460
    cmdview1.Enabled = False
    cmdview2.Enabled = True
    
End Sub

Private Sub Command6_Click()

End Sub

Private Sub cmdview2_Click()
    frmmailers.Width = 4620
    cmdview1.Enabled = True
    cmdview2.Enabled = False
    
End Sub

