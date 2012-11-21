Attribute VB_Name = "Version"
Option Explicit

Public sVerMaj As Long
Public sVerMin As Long
Public sVerRev As Long

Public sReleaseDate As String
Public sRunningVersion As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
                                                    
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

Public Sub StayOnTop(Form As Form)

    Dim SetWin As Long
    SetWin = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    
End Sub

Public Sub SetVersionInfo()
   
    sVerMaj = App.Major
    sVerMin = App.Minor
    sVerRev = App.Revision
    
    sRunningVersion = sVerMaj & "." & sVerMin & "." & sVerRev
    
    sReleaseDate = "08.30.2006"
    
End Sub
