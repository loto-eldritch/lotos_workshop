VERSION 5.00
Begin VB.Form SwitchForm 
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SwitchForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame SwitchFrame 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   4650
      Begin VB.Timer timeOut 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   3780
         Top             =   540
      End
      Begin VB.Label Status 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   240
         Left            =   90
         TabIndex        =   4
         Top             =   1125
         Width           =   4470
      End
      Begin VB.Label SwitchStatus 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   915
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   4470
      End
   End
   Begin VB.Label ProgressBar 
      BackColor       =   &H00FF8080&
      Height          =   105
      Left            =   25
      TabIndex        =   3
      Tag             =   "4575"
      Top             =   1350
      Width           =   15
   End
   Begin VB.Label lbl_ProgressBG 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   1305
      Width           =   4650
   End
End
Attribute VB_Name = "SwitchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const OLDREGKEY As String = "software\xpg\character constructor"
Private Const NEWREGKEY As String = "Software\Experimental Playground\Loto's Character Workshop"

Private Const PATCH_INFO_LOC As String = "http://www.experimental-playground.com/projects/workshop/patch"
Private Const WORKSHOP_PATCH_LOC As String = "http://www.experimental-playground.com/projects/workshop/patch.bin"

Private WORKSHOP_APP_LOC As String
Private PATCH_BIN_LOC As String

Private Const HKLM = HKEY_LOCAL_MACHINE

Private TIMED_OUT As Boolean

Private Sub TransferRegistrySettings()

    Dim sStr As String
    Dim lNum As Long
    
    'copy patcher version
    sStr = GetRegValue(HKLM, OLDREGKEY, "PATCHER_VERSION")
    If Val(sStr) <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, STR_VAL, "PATCHER_VERSION", sStr)
    'copy bonus option
    lNum = GetRegValue(HKLM, OLDREGKEY, "BONUS_OPT")
    If lNum <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, NUM_VAL, "BONUS_OPT", Str(lNum))
    'copy realm
    lNum = GetRegValue(HKLM, OLDREGKEY, "REALM")
    If lNum <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, NUM_VAL, "REALM", Str(lNum))
    'copy crafter skill
    lNum = GetRegValue(HKLM, OLDREGKEY, "CRAFTER_LEVEL")
    If lNum <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, NUM_VAL, "CRAFTER_LEVEL", Str(lNum))
    'copy cbx
    lNum = GetRegValue(HKLM, OLDREGKEY, "CBX")
    If lNum <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, NUM_VAL, "CBX", Str(lNum))
    'copy cby
    lNum = GetRegValue(HKLM, OLDREGKEY, "CBY")
    If lNum <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, NUM_VAL, "CBY", Str(lNum))
    'copy scx
    lNum = GetRegValue(HKLM, OLDREGKEY, "SCX")
    If lNum <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, NUM_VAL, "SCX", Str(lNum))
    'copy scy
    lNum = GetRegValue(HKLM, OLDREGKEY, "SCY")
    If lNum <> REG_KEY_NOT_EXISTS Then Call SetNewValue(HKLM, NEWREGKEY, NUM_VAL, "SCY", Str(lNum))
     
End Sub

Private Sub Form_Load()
    
    If App.PrevInstance Then End
    
    Dim hFile As Long
    Dim lCnt As Long
    
    Dim sBuf As String * 16384  '16 KB chunk
    Dim sBuffer As String
    
    Dim lRead As Long
    Dim lWrite As Long
    
    Dim lRet As Long
    Dim lErr As Long
    
    Me.Show
    StayOnTop Me
    
    DoEvents

    On Error GoTo Err:
    
    WORKSHOP_APP_LOC = App.Path & "\Workshop.exe"
    PATCH_BIN_LOC = App.Path & "\patch.bin"
    
    SwitchStatus.Caption = "Loto's Character Workshop" & vbCrLf & vbCrLf & _
                            "     ¤ Updating to Current Version"
    
    DoEvents
        
    'kill the old registry key entry
    If KeyExists(HKLM, OLDREGKEY) Then
        Call TransferRegistrySettings
        Call DeleteRegKey(HKLM, "software\xpg\character constructor")
        Call DeleteRegKey(HKLM, "software\xpg")
    End If
    
    'Sleep 250
    '-------NEW CODE---------------------------------------
    TIMED_OUT = False
    timeOut.Enabled = True
    Do
        DoEvents
        lRet = CreateFile(WORKSHOP_APP_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    Loop While (lRet = -1) And Not TIMED_OUT
    timeOut.Enabled = False
    
    TIMED_OUT = False
    timeOut.Enabled = True
    Do
        DoEvents
        hFile = CreateFile(PATCH_BIN_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    Loop While (hFile = -1) And Not TIMED_OUT
    timeOut.Enabled = False
    
    '------------------------------------------------------
    
    'check to see if the application file is there before killing to prevent an error
    'lRet = CreateFile(WORKSHOP_APP_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    
    'check to see if the patch.bin file is present before opening it
    'hFile = CreateFile(PATCH_BIN_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    
    If lRet <> -1 Then
        If hFile <> -1 Then
            lRet = CloseHandle(lRet)
            Kill WORKSHOP_APP_LOC
        Else
            lErr = 1
            GoTo Err:
        End If
    Else
        'download latest version
        Call Update(ProgressBar, Status)
        lErr = 0
        GoTo Err:
    End If
    
    'if the patch.bin is there, open it and read it into a buffer
    Do
        DoEvents
        If lCnt Mod 25 = 0 Then SwitchStatus.Caption = SwitchStatus.Caption & "."
        sBuf = vbNullString
        
        lRet = ReadFile(hFile, ByVal sBuf, Len(sBuf), lRead, ByVal CLng(0))
        
        If lRead = 0 Then Exit Do
        
        sBuffer = sBuffer & sBuf
        lCnt = lCnt + 1
    Loop
    SwitchStatus.Caption = SwitchStatus.Caption & "OK!"
    
    'close the patch.bin file handle
    lRet = CloseHandle(hFile)
    
    sBuffer = Trim(sBuffer)
    
    'check to see if the patch file is there before killing
    lRet = CreateFile(PATCH_BIN_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    If lRet <> -1 Then
        lRet = CloseHandle(lRet)
        Kill PATCH_BIN_LOC
    End If
    
    'create the new workshop.exe file
    hFile = CreateFile(WORKSHOP_APP_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
    
    If hFile = -1 Then
        lErr = 2
        GoTo Err:
    Else
        lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
        lRet = CloseHandle(hFile)
    End If
    
    lErr = 0
    
Err:
    Select Case lErr
        Case 0
            SwitchStatus.Caption = SwitchStatus.Caption & vbCrLf & "     ¤ Update Success!"
            DoEvents
        Case 1
            SwitchStatus.Caption = SwitchStatus.Caption & vbCrLf & "     ¤ Patch is Either Missing or Corrupted!"
            DoEvents
        Case 2
            SwitchStatus.Caption = SwitchStatus.Caption & vbCrLf & "     ¤ Unable to Access Character Workshop to Issue Patch!"
            DoEvents
    End Select
    
    SwitchStatus.Caption = SwitchStatus.Caption & vbCrLf & "     ¤ Restarting The Workshop..."
    DoEvents
    
    lRet = ShellExecute(SwitchForm.hwnd, "open", WORKSHOP_APP_LOC, vbNullString, vbNullString, SW_SHOW)
    
    End
    
End Sub

Public Sub Update(ByRef ProgressBar As Label, ByRef Status As Label)
    
    Dim hFile As Long
    Dim lRet As Long
   
    Dim lPatchSize As Long
    
    Dim lWrite As Long
    
    Dim sBuffer As String
    Dim sMessage As String
        
    Dim sPatchInfo As String
    
    Status.Caption = "Checking Connection..."
    ProgressBar.Width = ProgressBar.Width + 100
    
    If Connection_Online Then
    
        Status.Caption = "Checking for Updates..."
        ProgressBar.Width = ProgressBar.Width + 100
        
        'download patch information file
        sPatchInfo = Trim(Download_File(PATCH_INFO_LOC))
        
        'get current patch size
        lPatchSize = Val(Mid(sPatchInfo, _
                        InStr(sPatchInfo, "<size>") + 6, _
                        InStr(sPatchInfo, "</size>") - InStr(sPatchInfo, "<size>") - 6))
        
        'set new download message
        sMessage = "Downloading Latest Version"
        
        'download new patch
        sBuffer = Trim(Download_File(WORKSHOP_PATCH_LOC, Status, sMessage, ProgressBar, lPatchSize))
        
        'create new patch file
        hFile = CreateFile(WORKSHOP_APP_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
        
        If hFile = -1 Then Exit Sub
        
        'write new patch file to disk
        lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
        lRet = CloseHandle(hFile)
                    
        'clear buffer
        sBuffer = vbNullString
        
        'set new status caption
        Status.Caption = vbNullString
        ProgressBar.Width = Val(ProgressBar.Tag)
        'Call Sleep(1000)
    Else
        Status.Caption = "No Internet Connection Found"
    End If
    
End Sub

Private Sub timeOut_Timer()

    TIMED_OUT = True
    
End Sub
