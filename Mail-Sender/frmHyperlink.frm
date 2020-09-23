VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Config DM Easy Mailer."
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   ClipControls    =   0   'False
   Icon            =   "frmHyperlink.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtServPort 
      Height          =   285
      Left            =   225
      TabIndex        =   2
      Top             =   1620
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   350
      Left            =   2295
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2805
      Width           =   990
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   350
      Left            =   1245
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2805
      Width           =   990
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   350
      Left            =   180
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2805
      Width           =   990
   End
   Begin VB.TextBox txtMailName 
      Height          =   285
      Left            =   225
      TabIndex        =   3
      Top             =   2295
      Width           =   3090
   End
   Begin VB.TextBox txtServName 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   930
      Width           =   4350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current SMTP servers port number eg 25"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1350
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter the name of the email from eg tom@yourserver.com"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   4050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter the current name of the SMTP server you wish to use."
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   690
      Width           =   4230
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmHyperlink.frx":0442
      Top             =   75
      Width           =   480
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    If FirstTimeLoad Then
        Ans = MsgBox("You have not set anything up are you sure you want to quit now", _
        vbYesNo Or vbInformation)
        If Ans = vbNo Then
            Exit Sub
        Else
            Unload frmConfig
            Unload frmMailSender
            Set frmConfig = Nothing
            Set frmMailSender = Nothing
            End
        End If
    Else
        Unload frmConfig
        Set frmConfig = Nothing
        frmMailSender.Show
    End If
    
End Sub

Private Sub cmdOk_Click()
    If Len(Trim(txtServName.Text)) <= 0 Then
        MsgBox "You must enter a vaild server name", vbCritical, "Error no name entered"
        Exit Sub
    ElseIf IsNumeric(txtServPort.Text) = False Then
        MsgBox "Invaild port number", vbCritical, "Invaild Port Number"
        Exit Sub
    ElseIf Len(Trim(txtMailName.Text)) <= 0 Then
        MsgBox "You must at least enter something for the E-Mail display name", vbInformation, "No E-Mail name entered"
        Exit Sub
    ElseIf isVaildEmail(txtMailName) = False Then
        MsgBox "Invaild E-Mail Address Entered", vbCritical, "Invaild E-Mail Addesss"
        Exit Sub
    Else
        WritePrivateProfileString "DM-EASYMAIL", "Servername", txtServName.Text, AddBackSlash(App.Path) & "config.ini"
        WritePrivateProfileString "DM-EASYMAIL", "Serverport", txtServPort.Text, AddBackSlash(App.Path) & "config.ini"
        WritePrivateProfileString "DM-EASYMAIL", "Mailtext", txtMailName.Text, AddBackSlash(App.Path) & "config.ini"
        MsgBox "Your new setting will take affect the next time you load the program", vbInformation, "Finished"
        Unload frmConfig
        
    End If
    
End Sub

Private Sub cmdReset_Click()
    txtServName = ""
    txtServPort = ""
    txtMailName = ""
    
End Sub

Private Sub Form_Load()
    txtMailName.Text = TMail.MailFrom
    txtServName.Text = TMail.MailServer
    txtServPort.Text = TMail.MailServerPort
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConfig = Nothing
    
End Sub

Private Sub txtMailName_GotFocus()
    txtMailName.BackColor = 14073525
    
End Sub

Private Sub txtMailName_LostFocus()
    txtMailName.BackColor = vbWhite
    
End Sub

Private Sub txtServName_GotFocus()
    txtServName.BackColor = 14073525
 
End Sub

Private Sub txtServName_LostFocus()
    txtServName.BackColor = vbWhite
    
End Sub

Private Sub txtServPort_GotFocus()
     txtServPort.BackColor = 14073525
     
End Sub

Private Sub txtServPort_LostFocus()
    txtServPort.BackColor = vbWhite
    
End Sub
