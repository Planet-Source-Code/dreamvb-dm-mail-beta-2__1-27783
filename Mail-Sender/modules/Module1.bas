Attribute VB_Name = "Module1"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Type MailConfig
    MailServerPort As Integer
    MailServer As String
    MailFrom As String
    MailTo As String
    Subject As String
    MailBody As String
    MailMess As String
    StrDate As String
End Type

Public TMail As MailConfig
Public FirstTimeLoad As Boolean
Function AddBackSlash(lzPath As String) As String
    If Right(lzPath, 1) <> "\" Then AddBackSlash = lzPath & "\" Else AddBackSlash = lzPath
    
End Function
Public Function FindFile(lzFileName As String) As Boolean
    If Dir(lzFileName) <> "" Then FindFile = True Else FindFile = False
    
End Function
Public Sub FlatBorder(ByVal hwnd As Long)
Dim TFlat As Long
  TFlat = GetWindowLong(hwnd, GWL_EXSTYLE)
  TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
  SetWindowLong hwnd, GWL_EXSTYLE, TFlat
  SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
  
End Sub

Public Function OpenFile(THwnd As Long) As String
Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = THwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.DM Mail)" + Chr$(0) + "*.DME"
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = App.Path & "\"
    ofn.lpstrTitle = "Open Email"
    ofn.flags = 0
    A = GetOpenFileName(ofn)
        If (A) Then
            OpenFile = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr(0)) - 1)
        End If
        
 End Function
Public Function SaveFile(THwnd As Long) As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = THwnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = "All Files(*.DM Mail)" + Chr$(0) + "*.DME"
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Save Project"
        ofn.flags = 0
        
        A = GetSaveFileName(ofn)
        If (A) Then
                SaveFile = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, Chr(0)) - 1)
        End If
 End Function
Public Function ReadConfig(AppName As String, StrKey As String) As String
Dim StrBuff As String
Dim Xpos As Integer
    StrBuff = String(255, Chr(0))
    GetPrivateProfileString AppName, StrKey, "ERROR", StrBuff, 255, AddBackSlash(App.Path) & "config.ini"
    ReadConfig = Left(StrBuff, InStr(StrBuff, Chr(0)) - 1)
    
End Function

Function isVaildEmail(EmailName As String) As Boolean
Dim ipart As Integer, lpart As Integer, Length As Integer
Dim isVaild As Boolean
Dim sEmail As String
    If Len(Trim(EmailName) <= 0) Then isVaildEmail = False
    sEmail = Trim(EmailName)
    ipart = InStr(sEmail, "@")
    lpart = InStr(ipart + 1, sEmail, ".")

        Length = Len(Trim(Mid(sEmail, lpart + 1, 3)))
        If ipart <= 0 Or lpart <= 0 Then
            isVaild = False
        ElseIf Length < 3 Then
            isVaild = False
        ElseIf ipart = 1 Then
            isVaild = False
        ElseIf lpart = Len(sEmail) Then
            isVaild = False
        Else
            isVaild = True
        End If
        isVaildEmail = isVaild
End Function

Function MoveForm(mHwnd As Form)
    ReleaseCapture
    SendMessage mHwnd.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 1

End Function
