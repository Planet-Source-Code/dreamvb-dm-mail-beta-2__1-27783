VERSION 5.00
Begin VB.Form frmmenu 
   ClientHeight    =   1230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   1695
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnupaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuselall 
         Caption         =   "Select& All"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
