Attribute VB_Name = "Controls"
'Controls.bas By CyberAscii...

'Declarations
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal DWreserved As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


'Constants
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETCURSEL = &H188
Public Const LB_SETCURSEL = &H186
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const RGN_XOR = 3
Public Const SND_ASYNC = &H1
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const VK_SPACE = &H20
Public Const WM_ENABLE = &HA
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_KEYUP = &H101
Public Const WM_KEYDOWN = &H100
Public Const WM_COMMAND = &H111


'Types
Private Type POINTAPI
x As Long
Y As Long
End Type

Function cGetCaption(ByVal TheWnd As Long) As String

    Dim WndLength As Long, cCap As String
    WndLength = GetWindowTextLength(TheWnd)
    cCap$ = String(WndLength&, 0&)
    Call GetWindowText(TheWnd&, cCap$, WndLength& + 1)
    cGetCaption$ = cCap$

End Function

Function cGetText(ByVal TheWnd As Long) As String

    Dim WndLength As Long, cTxt As String
    WndLength = SendMessage(TheWnd, WM_GETTEXTLENGTH, 0&, 0&)
    cTxt$ = String(WndLength&, 0&)
    Call SendMessageByString(TheWnd&, WM_GETTEXT, WndLength& + 1, cTxt$)
    cGetText$ = cTxt$

End Function

Public Sub cCloseWindow(ByVal TheWnd As Long)

Call PostMessage(TheWnd, WM_CLOSE, 0&, 0&)

End Sub
Sub cSetCaption(Window As Long, Caption As String)

Dim SetCaption As Long
SetCaption& = SendMessageByString(Window&, WM_SETTEXT, 0, Caption)

End Sub
Public Sub cSetText(Window As Long, ByVal Text As String)

Call SendMessageByString(Window&, WM_SETTEXT, 0, Text$)

End Sub
Sub cClickButton(Button As Long)

Dim ClickIt As Long
ClickIt& = SendMessage(Button&, WM_KEYDOWN, VK_SPACE, vbNullString)
ClickIt& = SendMessage(Button&, WM_KEYUP, VK_SPACE, vbNullString)

End Sub
Function cGetWindow(TheWindow As String) As Long

cGetWindow& = FindWindow(TheWindow, vbNullString)

End Function
Public Sub cHideWindow(ByVal TheWnd As Long)

    DoEvents:
    Call ShowWindow(TheWnd, SW_HIDE)

End Sub

Public Sub cShowWindow(ByVal TheWnd As Long)

    DoEvents:
    Call ShowWindow(TheWnd, SW_SHOW)

End Sub
Public Sub cOnTop(ByVal TheWnd As Long, ByVal IsOnTop As Boolean)

    If bOnTop = True Then
    Call SetWindowPos(TheWnd&, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
    Call SetWindowPos(TheWnd&, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
    End If

End Sub
