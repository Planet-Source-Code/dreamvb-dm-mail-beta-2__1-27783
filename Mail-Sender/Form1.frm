VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Begin VB.Form frmMailSender 
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7140
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox resMaxBut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1890
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   21
      Top             =   7155
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox resMinBut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1890
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   20
      Top             =   6915
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox resCloseBut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   180
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   19
      Top             =   7140
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox resTBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   135
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   15
      Top             =   6900
      Visible         =   0   'False
      Width           =   1650
   End
   Begin Project1.Bevel Bevel2 
      Height          =   3645
      Left            =   0
      TabIndex        =   4
      Top             =   2205
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   6429
   End
   Begin Project1.Line3D Line3D2 
      Height          =   105
      Left            =   0
      TabIndex        =   13
      Top             =   870
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   185
   End
   Begin Project1.Line3D Line3D1 
      Height          =   105
      Left            =   0
      TabIndex        =   12
      Top             =   345
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   185
   End
   Begin VB.TextBox txtMail 
      Height          =   285
      Index           =   2
      Left            =   885
      TabIndex        =   3
      ToolTipText     =   "Type in your subject here"
      Top             =   1785
      Width           =   6120
   End
   Begin VB.PictureBox skntitle 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   476
      TabIndex        =   11
      Top             =   0
      Width           =   7140
      Begin VB.PictureBox TMinBut 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6225
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   45
         Width           =   255
      End
      Begin VB.PictureBox TMaxBut 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6525
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   45
         Width           =   255
      End
      Begin VB.PictureBox TCloseBut 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6825
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   16
         Top             =   45
         Width           =   255
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Easy Mail Sender Beta 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Top             =   75
         Width           =   2445
      End
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   885
      TabIndex        =   1
      ToolTipText     =   "Type in the address you whish to send to here"
      Top             =   1110
      Width           =   6120
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1215
      Top             =   5970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   1125
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   503
      ButtonWidth     =   1323
      ButtonHeight    =   503
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList3"
      HotImageList    =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "To    "
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      ToolTipText     =   "Mail send status"
      Top             =   5865
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   12091
            Picture         =   "Form1.frx":0966
            Text            =   "Status : Ide"
            TextSave        =   "Status : Ide"
         EndProperty
      EndProperty
   End
   Begin DHTMLEDLibCtl.DHTMLEdit DHTML 
      Height          =   3000
      Left            =   75
      TabIndex        =   7
      ToolTipText     =   "Main Message goes here"
      Top             =   2775
      Width           =   6960
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2498
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3064
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":364A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4216
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":47FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":53C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   45
      TabIndex        =   6
      Top             =   2340
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fonts"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ordered List"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Unordered List"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Text Outdent"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Text Indent"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Align Center"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Horizonal Line"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Line Break"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Breakeing Space"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Hyperlink"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Image"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   615
      Top             =   5955
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":59AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6052
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":63A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":64B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6808
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":71FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7550
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":78A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7D06
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":816A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":84BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":880E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Project1.Bevel Bevel1 
      Height          =   1170
      Left            =   30
      TabIndex        =   5
      Top             =   990
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   2064
   End
   Begin VB.TextBox txtMail 
      Height          =   285
      Index           =   1
      Left            =   885
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "This is were the mail come from"
      Top             =   1455
      Width           =   6120
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   714
      ButtonWidth     =   767
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Message"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnublnk"
                  Text            =   "Blank Message"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnunise"
                  Text            =   "Nise Day"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnunew"
                  Text            =   "Flower"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Send Message"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Message"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Message"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Config Options"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock smtp 
      Left            =   6390
      Top             =   6225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar4 
      Height          =   285
      Left            =   90
      TabIndex        =   10
      Top             =   1455
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   503
      ButtonWidth     =   1323
      ButtonHeight    =   503
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList3"
      DisabledImageList=   "ImageList3"
      HotImageList    =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "From"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subject"
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   1800
      Width           =   540
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   330
      Picture         =   "Form1.frx":8B60
      Top             =   6645
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   75
      Picture         =   "Form1.frx":8EEA
      Top             =   6645
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmMailSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String
Dim MailBody As String
Dim mIndex As Integer
Dim HTML_BODY As String
Dim IsMax As Boolean
Sub IsMaxSized()
    Select Case IsMax
        Case True
            IsMax = False
            frmMailSender.WindowState = vbNormal
        Case False
            IsMax = True
            frmMailSender.WindowState = vbMaximized
    End Select
    
End Sub
Sub ApplySkin()
Dim SknHeight As Integer, SknButWidth As Integer, SknButHeight As Integer
    SknHeight = resTBar.Height / 2
    TCloseBut.Left = skntitle.ScaleWidth - 20
    TMaxBut.Left = TCloseBut.Left - 20
    TMinBut.Left = TMaxBut.Left - 20
    StretchBlt skntitle.hdc, 0, 0, skntitle.Width, SknHeight, resTBar.hdc, 0, 0, resTBar.Width, SknHeight, vbSrcCopy
    StretchBlt TCloseBut.hdc, 0, 0, 17, 17, resCloseBut.hdc, 0, 0, 17, 17, vbSrcCopy
    StretchBlt TMaxBut.hdc, 0, 0, 17, 17, resMaxBut.hdc, 0, 0, 17, 17, vbSrcCopy
    StretchBlt TMinBut.hdc, 0, 0, 17, 17, resMinBut.hdc, 0, 0, 17, 17, vbSrcCopy
    skntitle.Refresh
    
End Sub
Sub SaveMessage()
            FileOut = FreeFile
            txtMail(0).Text = Trim(txtMail(0).Text)
            txtMail(1).Text = Trim(txtMail(1).Text)
            txtMail(2).Text = Trim(txtMail(2).Text)
            
            If Len(txtMail(0).Text) <= 0 Then
                MsgBox "The mail could not be saved No Mail From has been completed.", vbCritical, "Error"
                Exit Sub
            ElseIf Len(txtMail(1).Text) <= 0 Then
                MsgBox "The mail message could not be saved No Recipient has been completed.", vbCritical, "Error"
                Exit Sub
            ElseIf Len(txtMail(2).Text) <= 0 Then
                MsgBox "The mail message could not be saved No Subject has been completed.", vbCritical, "Error"
                Exit Sub
            ElseIf isVaildEmail(txtMail(0).Text) = False Then
                MsgBox "You have not entered a vaild email address", vbCritical, "Invaild Email Address"
                Exit Sub
            Else
                lzFileName = SaveFile(frmMailSender.hwnd)
                FileExt = UCase(Right(lzFileName, 3))
                If Not FileExt = "DME" Then lzFileName = lzFileName & ".dme"
                If Len(FileExt) <= 0 Then Exit Sub
                
                Open lzFileName For Output As #FileOut
                    Print #FileOut, "[DM Easy Email Sender Beta 1 Do not edit below this line]"
                    Print #FileOut, "[MailFrom]" & txtMail(0).Text
                    Print #FileOut, "[MAILTO]" & txtMail(1).Text
                    Print #FileOut, "[SUBJECT]" & txtMail(2).Text
                    Print #FileOut, "[MAILDATA]" & DHTML.DocumentHTML & Chr(25)
                Close #FileOut
            End If

End Sub
Function InsertHtmlCode(strCode As String) As String
Dim DOC As Object
Dim sel As Object
 On Error Resume Next
    Set DOC = DHTML.DOM
    Set sel = DOC.selection
    Set tr = sel.createRange
    tr.pasteHTML (strCode)
    
End Function

Function Ave() As String
Dim A As String
    ' This must be left in all the time as to let other people know about the program
    A = A & "<hr>"
    A = A & "<p>Message sent with DM Easy E-Mail Sender get your free copy today<br>" & vbCrLf
    A = A & "<br>" & vbCrLf
    A = A & "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
    A = A & "  <a href=€http://www.developers-answers.net/programs/€>DM Easy E-Mail home page</a></p>" & vbCrLf
    A = A & "</body>" & vbCrLf
    A = A & "</html>" & vbCrLf
    Ave = Replace(A, Chr(128), Chr(34))
    
End Function

Sub WaitFor(ResponseCode As String)
    ' This code in this function was not writen by me just found on the net
    ' But just like to say thank's to who ever did write it.
    
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64
            Exit Sub
        End If
    Wend

Response = "" ' Sent response code to blank **IMPORTANT**
End Sub
Sub SendEmail()
Dim MailBody As String
On Error Resume Next
    smtp.Close
    smtp.LocalPort = 0
    smtp.Protocol = sckTCPProtocol
    smtp.RemoteHost = TMail.MailServer
    smtp.RemotePort = TMail.MailServerPort
    smtp.Connect
    StatusBar1.Panels(1).Picture = Image1(0).Picture
    WaitFor ("220")
    StatusBar1.Panels(1).Text = "Status: Connecting to " & smtp.RemoteHost
    Replay "HELO " & smtp.LocalHostName
    WaitFor ("250")
    StatusBar1.Panels(1).Text = "Status: Sending mail message"
    '
    Replay "MAIL FROM: " & TMail.MailFrom
    WaitFor ("250")
    Replay "RCPT TO: " & TMail.MailTo
    WaitFor ("250")
    Replay "DATA"
    WaitFor ("354")
    Replay TMail.MailBody
    WaitFor ("250")
    StatusBar1.Panels(1).Text = "Status: Mail message sent"
    Replay "QUIT "
    StatusBar1.Panels(1).Text = "Status: Closing connection."
    WaitFor ("221")
    smtp.Close
    StatusBar1.Panels(1).Picture = Image1(1).Picture
    StatusBar1.Panels(1).Text = "Status : Ide"
    
End Sub

Sub Replay(StrBuff As String)
    If smtp.State = sckConnected Then
        smtp.SendData StrBuff & vbCrLf
    End If
    
End Sub

Private Sub DHTML_onclick()
    mIndex = 3
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = True
    Toolbar1.Buttons(8).Image = ImageList1.ListImages(7).Index
    Toolbar1.Buttons(9).Image = ImageList1.ListImages(5).Index
    Toolbar1.Buttons(10).Image = ImageList1.ListImages(9).Index
    
End Sub

Private Sub Form_Load()
Dim SknPath As String
Dim ProgLoad As Boolean


    SknPath = AddBackSlash(App.Path) & "skin\default\"
    
    FlatBorder txtMail(0).hwnd
    FlatBorder txtMail(1).hwnd
    FlatBorder txtMail(2).hwnd
    

    Bevel1.BevelSytle VbRaised
    Bevel2.BevelSytle VbRaised
    
    If FindFile(AddBackSlash(App.Path) & "config.ini") = False Then
        FirstTimeLoad = True
        frmMailSender.Hide
        MsgBox "Your new mail sender needs to be setup before you can use it.", vbInformation, "Setup Mail Sender"
        frmConfig.Show
        Exit Sub
    Else
        frmMailSender.Show
        txtMail(1).Text = ReadConfig("DM-EASYMAIL", "Mailtext")
        TMail.MailServer = ReadConfig("DM-EASYMAIL", "Servername")
        TMail.MailServerPort = Val(ReadConfig("DM-EASYMAIL", "Serverport"))
        TMail.MailFrom = ReadConfig("DM-EASYMAIL", "Mailtext")
    End If
    DHTML.SetFocus

    If FindFile(SknPath & "top.bmp") = False Then
        ProgLoad = False
    ElseIf FindFile(SknPath & "close.bmp") = False Then
        ProgLoad = False
    ElseIf FindFile(SknPath & "max.bmp") = False Then
        ProgLoad = False
    ElseIf FindFile(SknPath & "min.bmp") = False Then
        ProgLoad = False
    Else
        ProgLoad = True
    End If
    
    If ProgLoad = False Then
        MsgBox "The program could not be started due to missing files" _
        & "Please make sure that all the files exist and have not been deleted by mistake" _
        & vbCrLf & vbCrLf & "Please report any errors to dreamvb@yahoo.com don't forget to include the error number Thanks.", vbCritical, "Error 1054 Unable to start program."
        Unload frmsplash
        Unload frmMailSender
        End
    Else
        resTBar.Picture = LoadPicture(SknPath & "top.bmp")
        resCloseBut.Picture = LoadPicture(SknPath & "close.bmp")
        resMinBut.Picture = LoadPicture(SknPath & "min.bmp")
        resMaxBut.Picture = LoadPicture(SknPath & "max.bmp")
        ApplySkin
    End If


    
End Sub

Private Sub Form_Paint()
    ApplySkin
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Line3D1.Width = frmMailSender.ScaleWidth
    Line3D2.Width = frmMailSender.ScaleWidth
    Bevel1.Width = frmMailSender.ScaleWidth - 4
    Bevel2.Width = Bevel1.Width + 2
    txtMail(0).Width = Bevel1.Width - 62
    txtMail(1).Width = Bevel1.Width - 62
    txtMail(2).Width = Bevel1.Width - 62
    DHTML.Width = frmMailSender.ScaleWidth - 12
    DHTML.Height = frmMailSender.ScaleHeight - StatusBar1.Height - DHTML.Top - 10
    Bevel2.Height = frmMailSender.ScaleHeight - StatusBar1.Height - DHTML.Top + 40
    
    If frmMailSender.Height <= 2640 And frmMailSender.Width <= 2640 Then
        frmMailSender.Width = 2640
        frmMailSender.Height = 2640
    End If
    
    If frmMailSender.WindowState = vbNormal Then
        frmMailSender.Caption = ""
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Response = ""
    MailBody = ""
    mIndex = 0
    Set frmMailSender = Nothing
    Set frmabout = Nothing
    Set frmConfig = Nothing
    Unload frmMailSender
    Unload frmabout
    Unload frmConfig
    smtp.Close
    
End Sub

Private Sub lblCaption_DblClick()
    IsMaxSized
    
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MoveForm frmMailSender
    End If
    
    
End Sub

Private Sub skntitle_DblClick()
    IsMaxSized
    
End Sub

Private Sub skntitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        MoveForm frmMailSender
    End If
    
    
End Sub

Private Sub smtp_DataArrival(ByVal bytesTotal As Long)
    smtp.GetData Response

End Sub

Private Sub smtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    smtp.Close
    StatusBar1.Panels(1).Text = "Status: There was an error sending the message"
    
End Sub

Private Sub TCloseBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StretchBlt TCloseBut.hdc, 0, 0, 17, 17, resCloseBut.hdc, 34, 0, 17, 17, vbSrcCopy
    TCloseBut.Refresh
    
End Sub

Private Sub TCloseBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StretchBlt TCloseBut.hdc, 0, 0, 17, 17, resCloseBut.hdc, 0, 0, 17, 17, vbSrcCopy
    TCloseBut.Refresh
    Unload frmMailSender
    
End Sub

Private Sub TMaxBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StretchBlt TMaxBut.hdc, 0, 0, 17, 17, resMaxBut.hdc, 34, 0, 17, 17, vbSrcCopy
    TMaxBut.Refresh
    
End Sub

Private Sub TMaxBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ison As Boolean
    StretchBlt TMaxBut.hdc, 0, 0, 17, 17, resMaxBut.hdc, 0, 0, 17, 17, vbSrcCopy
    TMaxBut.Refresh
    IsMaxSized
    

End Sub

Private Sub TMinBut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StretchBlt TMinBut.hdc, 0, 0, 17, 17, resMinBut.hdc, 34, 0, 17, 17, vbSrcCopy
    TMinBut.Refresh
    
End Sub

Private Sub TMinBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StretchBlt TMinBut.hdc, 0, 0, 17, 17, resMinBut.hdc, 0, 0, 17, 17, vbSrcCopy
    TMinBut.Refresh
    frmMailSender.WindowState = vbMinimized
    frmMailSender.Caption = lblCaption.Caption
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim FileExt As String, lzFileName As String, MailData As String
Dim FileIn As Long, FileOut As Long
Dim ipart As Integer, lpart As Integer

    FileIn = FreeFile
    
    Select Case Button.Index
        Case 3
            On Error Resume Next
            TMail.MailTo = Trim(txtMail(0).Text)
            TMail.Subject = Trim(txtMail(2).Text)
            TMail.StrDate = Format(Now, "ddd, dd mmm yyyy hh:mm:ss  +0100")
            TMail.MailMess = DHTML.DocumentHTML & Ave
            
    
            If Len(TMail.MailTo) <= 0 Then
                MsgBox "The mail could not be send please include a recipient name.", vbCritical, "Error"
                Exit Sub
            End If
            If isVaildEmail(TMail.MailTo) = False Then
                MsgBox "You have not enter a invalid email address the mail will not be sent.", vbCritical, "Inviald E-Mail Address"
                txtMail(0).SetFocus
                txtMail(0).SelStart = 0
                txtMail(0).SelLength = Len(txtMail(0))
                Exit Sub
            End If
            If Len(TMail.Subject) <= 0 Then
                TMail.Subject = "No Subject...."
            End If
            
    ' Main mail message body
            TMail.MailBody = "Date: " & TMail.StrDate & vbCrLf _
            & "From: " & Mid(TMail.MailFrom, 1, InStr(TMail.MailFrom, "@") - 1) & " " & "<" & TMail.MailFrom & ">" & vbCrLf _
            & "X-Mailer: Dm Mail Sender V1.1" & vbCrLf _
            & "X-Accept-Language: en" & vbCrLf _
            & "MIME-Version: 1.0" & vbCrLf _
            & "To: " & TMail.MailTo & vbCrLf _
            & "Subject: " & TMail.Subject & vbCrLf _
            & "Content-Type: text/html;" & vbCrLf _
            & vbTab & "charset=" & Chr(34) & "iso-8859-1" & Chr(34) & vbCrLf _
            & "Content-Transfer-Encoding: 7bit" & vbCrLf _
            & vbCrLf & TMail.MailMess & vbCrLf & "."
            SendEmail
            
        Case 5
            lzFileName = OpenFile(frmMailSender.hwnd)
            If Len(lzFileName) <= 0 Then Exit Sub
            FileExt = UCase(Right(lzFileName, 3))
            If Not FileExt = "DME" Then
                MsgBox "This is not a valid DM Easy Mail document.", vbCritical, "Error"
                Exit Sub
            Else
                Open lzFileName For Binary As #FileIn
                    MailData = Space(LOF(FileIn))
                    Get #FileIn, , MailData
                Close #FileIn
                ipart = InStr(MailData, "[MailFrom]")
                lpart = InStr(ipart, MailData, vbCrLf)
                txtMail(0).Text = Mid(MailData, ipart + 10, lpart - ipart - 10)
            
                ipart = InStr(MailData, "[MAILTO]")
                lpart = InStr(ipart, MailData, vbCrLf)
                txtMail(1).Text = Mid(MailData, ipart + 8, lpart - ipart - 8)
                
                ipart = InStr(MailData, "[SUBJECT]")
                lpart = InStr(ipart, MailData, vbCrLf)
                txtMail(2).Text = Mid(MailData, ipart + 9, lpart - ipart - 9)
                    
                ipart = InStr(MailData, "[MAILDATA]")
                lpart = InStr(ipart, MailData, Chr(25))
                DHTML.DocumentHTML = Mid(MailData, ipart + 10, lpart - ipart - 10)
                
                MailData = ""
                lzFileName = ""
                FileExt = ""
                ipart = 0
                lpart = 0
            End If
            
        Case 6
            SaveMessage
            
        Case 8
            If mIndex = 3 Then
                DHTML.ExecCommand DECMD_CUT
            Else
                Clipboard.SetText txtMail(mIndex).SelText
                txtMail(mIndex).SelText = ""
            End If
            
        Case 9
            If mIndex = 3 Then
                DHTML.ExecCommand DECMD_COPY
            Else
                Clipboard.SetText txtMail(mIndex).SelText
            End If
        Case 10
            If mIndex = 3 Then
                DHTML.ExecCommand DECMD_PASTE
            Else
                txtMail(mIndex).SelText = Clipboard.GetText
            End If
            
        Case 12
            frmConfig.Show vbModal
        Case 14
            frmabout.Show vbModal
        Case 15
            Unload frmMailSender
            
        End Select
        
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim Ans
    Select Case ButtonMenu.Key
        Case "mnublnk"
            txtMail(0) = Trim(txtMail(0))
            If Len(txtMail(0).Text) > 0 And Len(txtMail(2).Text) > 0 Then
                Ans = MsgBox("Do you want to save your changes first", vbYesNo Or vbInformation)
                If Ans = vbYes Then
                    SaveMessage
                    DHTML.NewDocument
                    Exit Sub
                Else
                    DHTML.NewDocument
                End If
            End If
        Case "mnunise"
            txtMail(0) = Trim(txtMail(0))
            txtMail(1) = Trim(txtMail(1))
            If Len(txtMail(0).Text) > 0 Or Len(txtMail(2).Text) > 0 Then
                Ans = MsgBox("Do you want to save your changes first", vbYesNo Or vbInformation)
                If Ans = vbYes Then
                    SaveMessage
                    DHTML.LoadDocument AddBackSlash(App.Path) & "headers\header1.htm"
                    Exit Sub
                Else
                    DHTML.LoadDocument AddBackSlash(App.Path) & "headers\header1.htm"
                    Exit Sub
                End If
            Else
                    DHTML.LoadDocument AddBackSlash(App.Path) & "headers\header1.htm"
            End If
        Case "mnunew"
            txtMail(0) = Trim(txtMail(0))
            txtMail(1) = Trim(txtMail(1))
            If Len(txtMail(0).Text) > 0 Or Len(txtMail(2).Text) > 0 Then
                Ans = MsgBox("Do you want to save your changes first", vbYesNo Or vbInformation)
                If Ans = vbYes Then
                    SaveMessage
                    DHTML.LoadDocument AddBackSlash(App.Path) & "headers\header1.htm"
                    Exit Sub
                Else
                    DHTML.LoadDocument AddBackSlash(App.Path) & "headers\header1.htm"
                    Exit Sub
                End If
            Else
                    DHTML.LoadDocument AddBackSlash(App.Path) & "headers\header2.htm"
            End If
            
    End Select
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Index
        Case 2
            DHTML.ExecCommand DECMD_BOLD
        Case 3
            DHTML.ExecCommand DECMD_ITALIC
        Case 4
            DHTML.ExecCommand DECMD_UNDERLINE
        Case 5
            DHTML.ExecCommand DECMD_FONT
        Case 7
            DHTML.ExecCommand DECMD_ORDERLIST
        Case 8
            DHTML.ExecCommand DECMD_UNORDERLIST
        Case 9
            DHTML.ExecCommand DECMD_OUTDENT
        Case 10
            DHTML.ExecCommand DECMD_INDENT
        Case 12
            DHTML.ExecCommand DECMD_JUSTIFYLEFT
            Toolbar2.Buttons(13).Value = tbrUnpressed
            Toolbar2.Buttons(14).Value = tbrUnpressed
        Case 13
            DHTML.ExecCommand DECMD_JUSTIFYCENTER
            Toolbar2.Buttons(12).Value = tbrUnpressed
            Toolbar2.Buttons(14).Value = tbrUnpressed
        Case 14
            DHTML.ExecCommand DECMD_JUSTIFYRIGHT
            Toolbar2.Buttons(13).Value = tbrUnpressed
            Toolbar2.Buttons(12).Value = tbrUnpressed
        Case 16
            InsertHtmlCode "<hr>"
        Case 17
            InsertHtmlCode "<br>"
        Case 18
            InsertHtmlCode "&nbsp;"
        Case 19
            DHTML.ExecCommand DECMD_HYPERLINK
        Case 20
            DHTML.ExecCommand DECMD_IMAGE
        End Select
    
End Sub

Private Sub txtMail_Click(Index As Integer)
    mIndex = Index
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(9).Enabled = True
    Toolbar1.Buttons(10).Enabled = True

    Toolbar1.Buttons(8).Image = ImageList1.ListImages(7).Index
    Toolbar1.Buttons(9).Image = ImageList1.ListImages(5).Index
    Toolbar1.Buttons(10).Image = ImageList1.ListImages(9).Index
    
End Sub

Private Sub txtMail_GotFocus(Index As Integer)
    txtMail(Index).BackColor = 14073525
    
End Sub

Private Sub txtMail_LostFocus(Index As Integer)
    txtMail(Index).BackColor = vbWhite
    
End Sub

