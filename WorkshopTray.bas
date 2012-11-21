Attribute VB_Name = "WorkshopTray"
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
                         (ByVal dwMessage As Long, _
                          lpData As NOTIFYICONDATA) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                         (ByVal hwnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                         (ByVal lpPrevWndFunc As Long, _
                          ByVal hwnd As Long, _
                          ByVal Msg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

Private Const NIF_ICON = &H2
Private Const NIF_INFO = &H10
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1
Private Const NIM_SETFOCUS = &H4

Private Const GWL_WNDPROC = (-4)

Private Const WM_MOUSEMOVE = &H200

Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDOWN = &H201

Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDOWN = &H204

Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDOWN = &H207

Private Const PK_TRAYICON = &H401

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
        'The following data members are only valid in Windows 2000!
        '(uncomment the following lines to use them)
        'dwState As Long
        'dwStateMask As Long
        'szInfo As String * 256
        'uTimeoutOrVersion As Long
        'szInfoTitle As String * 64
        'dwInfoFlags As Long
End Type

Private nfIconData As NOTIFYICONDATA
Private appHandle As Long
Private WndProc As Long
Private isHooked As Boolean

Public Sub CreateSystrayIcon(wHandle As Long, lpIcon As Long, lpIconHandle As Long, Tip As String)
    With nfIconData
        .hwnd = wHandle
        .uID = lpIcon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = lpIconHandle
        .szTip = Tip & vbNullChar
        .cbSize = Len(nfIconData)
    End With
    Shell_NotifyIcon NIM_ADD, nfIconData
End Sub

Public Sub DestroySystrayIcon()
    Shell_NotifyIcon NIM_DELETE, nfIconData
End Sub

'Public Sub Hook(wHandle As Long)
'    If isHooked = False Then
'        appHandle = wHandle
'        WndProc = SetWindowLong(wHandle, GWL_WNDPROC, AddressOf WindowProc)
'        isHooked = True
'    End If
'End Sub
'
'Public Sub UnHook()
'    If isHooked = True Then
'        Call SetWindowLong(appHandle, GWL_WNDPROC, WndProc)
'        isHooked = False
'    End If
'End Sub
'
'Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    If isHooked = True Then
'        If uMsg = PK_TRAYICON Then
'        Select Case lParam
'            Case WM_RBUTTONUP
'                'load a popup menu
'                WS.PopupMenu mnuPopup
'                WindowProc = True
'                Exit Function
'            Case WM_LBUTTONUP
'                WS.WindowState = vbNormal
'                WS.Show
'                Call UnHook
'                Call DestroySystrayIcon
'                WindowProc = True
'                Exit Function
'        End Select
'            WindowProc = True
'            Exit Function
'        End If
'        WindowProc = CallWindowProc(WndProc, hwnd, uMsg, wParam, lParam)
'    End If
'End Function

