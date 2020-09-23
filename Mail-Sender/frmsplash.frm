VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   80
      Left            =   210
      Top             =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Writen and designed by Ben Jones"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   990
      TabIndex        =   1
      Top             =   2040
      Width           =   3585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Freeware Mail Sender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   915
      TabIndex        =   0
      Top             =   1785
      Width           =   3795
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   4275
      Picture         =   "frmsplash.frx":0000
      Top             =   2610
      Width           =   885
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   735
      Picture         =   "frmsplash.frx":1C62
      Top             =   45
      Width           =   4080
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "The program seems to be already running", vbInformation
        Unload frmsplash
        End
        Exit Sub
    End If
    
End Sub

Private Sub Timer1_Timer()
Static i As Integer
    i = i + 1
    If i = 10 Then
        Unload frmsplash
        frmMailSender.Show
        i = 0
    End If
        
End Sub
