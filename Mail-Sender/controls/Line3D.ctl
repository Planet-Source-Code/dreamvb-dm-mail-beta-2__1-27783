VERSION 5.00
Begin VB.UserControl Line3D 
   ClientHeight    =   105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   105
   ScaleWidth      =   4800
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   -150
      X2              =   1260
      Y1              =   45
      Y2              =   45
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -150
      X2              =   1260
      Y1              =   30
      Y2              =   30
   End
End
Attribute VB_Name = "Line3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
Dim Msg
 On Error Resume Next
    UserControl.Height = 105
    Line1(0).X2 = UserControl.Width
    Line1(1).X2 = UserControl.Width
 If Err Then Err.Clear
 
End Sub

Public Property Get TopColour() As OLE_COLOR
    TopColour = Line1(0).BorderColor
End Property

Public Property Let TopColour(ByVal New_TopColour As OLE_COLOR)
    Line1(0).BorderColor() = New_TopColour
    PropertyChanged "TopColour"
End Property

Public Property Get BottomColour() As OLE_COLOR
    BottomColour = Line1(1).BorderColor
End Property

Public Property Let BottomColour(ByVal New_BottomColour As OLE_COLOR)
    Line1(1).BorderColor() = New_BottomColour
    PropertyChanged "BottomColour"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Line1(0).BorderColor = PropBag.ReadProperty("TopColour", &H808080)
    Line1(1).BorderColor = PropBag.ReadProperty("BottomColour", &HFFFFFF)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TopColour", Line1(0).BorderColor, &H808080)
    Call PropBag.WriteProperty("BottomColour", Line1(1).BorderColor, &HFFFFFF)
End Sub


