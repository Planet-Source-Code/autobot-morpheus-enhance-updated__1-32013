VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " MorpheusEnhance"
   ClientHeight    =   720
   ClientLeft      =   6165
   ClientTop       =   5025
   ClientWidth     =   2160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   2160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTray 
      Caption         =   "Send To Tray"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Timer tmrKill 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   1680
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' Project    :       MorpheusEnhance
' Description:       Feel free to use this code in anyway but it is copywrited
'                    and if you use my code you must put my name somewhere in your program credits
' Created by :       ÅutoBot
' Machine    :       OPTIMUS-PRIME
' Date-Time  :       2/22/2002-1:40:26 AM
'
' Parameters :       N/A
'--------------------------------------------------------------------------------

Option Explicit

Private Sub cmdHide_Click()

    On Error Resume Next

    'This enables the timer that will hide any popup ads
    tmrKill.Enabled = True
    
    Dim kazaa As Long, shellembedding As Long, shelldocobjectview As Long, msctlsstatusbar As Long
    Dim internetexplorerserver As Long
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    internetexplorerserver = FindWindowEx(shelldocobjectview, 0&, "internet explorer_server", vbNullString)
    Call ShowWindow(internetexplorerserver, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    Call ShowWindow(shelldocobjectview, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    Call ShowWindow(shellembedding, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, 0&, "msctls_statusbar32", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, msctlsstatusbar, "msctls_statusbar32", vbNullString)
    Call ShowWindow(msctlsstatusbar, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    Call ShowWindow(kazaa, SW_NORMAL)
    Call ShowWindow(kazaa, SW_MAXIMIZE)

End Sub

Private Sub cmdShow_Click()
    
    On Error Resume Next

    'This stops the timer that hides popups
    tmrKill.Enabled = False
    
    'This shows any popups that have been formerly hidden
    Dim ieframe As Long
    ieframe = FindWindow("ieframe", vbNullString)
    Call ShowWindow(ieframe, SW_SHOW)

    Dim kazaa As Long, shellembedding As Long, shelldocobjectview As Long, msctlsstatusbar As Long
    Dim internetexplorerserver As Long
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    internetexplorerserver = FindWindowEx(shelldocobjectview, 0&, "internet explorer_server", vbNullString)
    Call ShowWindow(internetexplorerserver, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    Call ShowWindow(shelldocobjectview, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    Call ShowWindow(shellembedding, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, 0&, "msctls_statusbar32", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, msctlsstatusbar, "msctls_statusbar32", vbNullString)
    Call ShowWindow(msctlsstatusbar, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    Call ShowWindow(kazaa, SW_NORMAL)
    Call ShowWindow(kazaa, SW_MAXIMIZE)

End Sub

Private Sub cmdTray_Click()

    Hook Me.hwnd
    AddIconToTray Me.hwnd, Me.Icon, Me.Icon.Handle, App.Title
    'Ok Me.hwnd is what you hooked above, Me.Icon uses the main forms icon
    'in the systray, Me.Icon.Handle calls from the SysTray.bas, and App.Title
    'is what the tool tip tray will be, App.Title will make it the project name
    Me.Hide
    'Me.Hide hides the form when you click Tray

End Sub

Private Sub mnuRestore_Click()
    
    Unhook
    Me.Show
    'Important when restoring is to Unhook and RemoveIconFromTray
    RemoveIconFromTray

End Sub

Private Sub mnuHide_click()

    On Error Resume Next

    'This enables the timer that will hide any popup ads
    tmrKill.Enabled = True
    
    Dim kazaa As Long, shellembedding As Long, shelldocobjectview As Long, msctlsstatusbar As Long
    Dim internetexplorerserver As Long
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    internetexplorerserver = FindWindowEx(shelldocobjectview, 0&, "internet explorer_server", vbNullString)
    Call ShowWindow(internetexplorerserver, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    Call ShowWindow(shelldocobjectview, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    Call ShowWindow(shellembedding, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, 0&, "msctls_statusbar32", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, msctlsstatusbar, "msctls_statusbar32", vbNullString)
    Call ShowWindow(msctlsstatusbar, SW_HIDE)
    kazaa = FindWindow("kazaa", vbNullString)
    Call ShowWindow(kazaa, SW_NORMAL)
    Call ShowWindow(kazaa, SW_MAXIMIZE)

End Sub

Private Sub mnuShow_click()

    On Error Resume Next

    'This stops the timer that hides popups
    tmrKill.Enabled = False
    
    'This shows any popups that have been formerly hidden
    Dim ieframe As Long
    ieframe = FindWindow("ieframe", vbNullString)
    Call ShowWindow(ieframe, SW_SHOW)

    Dim kazaa As Long, shellembedding As Long, shelldocobjectview As Long, msctlsstatusbar As Long
    Dim internetexplorerserver As Long
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    internetexplorerserver = FindWindowEx(shelldocobjectview, 0&, "internet explorer_server", vbNullString)
    Call ShowWindow(internetexplorerserver, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    shelldocobjectview = FindWindowEx(shellembedding, 0&, "shell docobject view", vbNullString)
    Call ShowWindow(shelldocobjectview, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    shellembedding = FindWindowEx(kazaa, 0&, "shell embedding", vbNullString)
    Call ShowWindow(shellembedding, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, 0&, "msctls_statusbar32", vbNullString)
    msctlsstatusbar = FindWindowEx(kazaa, msctlsstatusbar, "msctls_statusbar32", vbNullString)
    Call ShowWindow(msctlsstatusbar, SW_SHOW)
    kazaa = FindWindow("kazaa", vbNullString)
    Call ShowWindow(kazaa, SW_NORMAL)
    Call ShowWindow(kazaa, SW_MAXIMIZE)

End Sub

Private Sub mnuExit_click()

    'This makes sure the icon doesn't stay in the tray when you exit
    RemoveIconFromTray
    Unload frmMain

End Sub

Public Sub SysTrayMouseEventHandler()

    'Read into the SysTray.bas you will see what this is for
    SetForegroundWindow Me.hwnd
    PopupMenu mnuMenu, vbPopupMenuRightButton

End Sub

'Puts my form on top of all others
Private Sub Form_Load()

    PutOnTop frmMain

End Sub

'This function hides any popup ads, the timer is set to update every 1000ms or 1 second
Private Sub tmrKill_Timer()
    
    On Error Resume Next

    Dim ieframe As Long
    ieframe = FindWindow("ieframe", vbNullString)
    Dim WindowCaption As String, TL As Long
    TL = SendMessageLong(ieframe, WM_GETTEXTLENGTH, 0&, 0&)
    WindowCaption = String(TL + 1, " ")
    Call SendMessageByString(ieframe, WM_GETTEXT, TL + 1, WindowCaption)
    WindowCaption = Left(WindowCaption, TL)
    Debug.Print WindowCaption

    If WindowCaption = "http://ads.musiccity.com/ - Microsoft Internet Explorer" Then
        
        Call ShowWindow(ieframe, SW_HIDE)

    End If

End Sub

