VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About DM Easy Mail Sender Beta 1."
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Height          =   360
      Left            =   1500
      TabIndex        =   2
      Top             =   1470
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Freeware and easy to use E-mail sender for Windows 95,95x,WinMe,Win2k"
      Height          =   465
      Left            =   855
      TabIndex        =   1
      Top             =   630
      Width           =   3150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DM Easy Mail Sender Beta 2"
      Height          =   195
      Left            =   1215
      TabIndex        =   0
      Top             =   240
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmabout.frx":0442
      Top             =   195
      Width           =   480
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload frmabout
    Set frmabout = Nothing
    
End Sub

