Attribute VB_Name = "Api"
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const WM_CLOSE = &H10
Public Const LB_GETTEXT = &H189
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

