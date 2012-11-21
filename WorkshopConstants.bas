Attribute VB_Name = "WorkshopConstants"
Option Explicit
'*******************************************************************************************************
'system api functions

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SetWindowPos Lib "user32" _
                        (ByVal hwnd As Long, _
                         ByVal hWndInsertAfter As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal cx As Long, _
                         ByVal cy As Long, _
                         ByVal wFlags As Long) As Long
                                                    
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                        (lpVersionInformation As OSVERSIONINFO) As Long
'*******************************************************************************************************
'system defined types
Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'*******************************************************************************************************
'system defined constants

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1

'*******************************************************************************************************
'user defined variables

Public MAX_HISTORY_COUNT As Long

Public RUNNING_VERSION As String
Public ABOUT_BOX_MESSAGE As String

Public PATCH_BIN_LOC As String
Public UPDATE_APP_LOC As String
Public WORKSHOP_APP_LOC As String

Public TEMPLATE_RECENT_PATH As String
Public TEMPLATE_RECENT_FILENAME As String

Public ERROR_LOG As String

Public WS_MAX_HEIGHT As Long
Public WS_MIN_HEIGHT As Long

Public WINDOWS_THEME As Long

'global SC form variable, True if the app is still Initializing, false otherwise
Public SPLASH_LOADING As Boolean
Public APPLICATION_PRE_START As Boolean
Public TEMPLATE_LOADING As Boolean

'global ws variable, if this is true then when a user clicks the X to close
'the workshop will go to the systray, also if the user minimizes the workshop it will do the same
Public EXIT_TO_SYSTRAY As Boolean
Public MINIMIZE_TO_SYSTRAY As Boolean
Public WORKSHOP_TRAY As Boolean
Public WORKSHOP_ENABLED As Boolean

'*******************************************************************************************************
'Auto Update flags
Public AUTO_UPDATE As Boolean
Public CHECK_UPDATE As Boolean
Public UPDATE_AVAILABLE As Boolean
Public UPDATE_CHECK_COUNTER As Long
Public Const UPDATE_CHECK_INTERVAL As Long = 10
'*******************************************************************************************************
'user defined constants
Public Const RELEASE_DATE As String = "9.29.2011"

Public Const ABOUT_BOX_TITLE As String = "About Loto's Character Workshop"
Public Const WORKSHOP_TITLE As String = "Loto's Character Workshop"
Public Const WORKSHOP_UPDATE_MESSAGE As String = " - Update Available!"

Public Const TEMPLATE_EXTENSION As String = ".lwt"

Public Const REGKEY As String = "Software\Experimental Playground\Loto's Character Workshop"

Public Const WEBSITE As String = "http://www.experimental-playground.com"
Public Const WEBSITE_FORUM As String = "http://www.experimental-playground.com"

Public Const PATCH_INFO_LOC As String = WEBSITE & "/projects/workshop/patch"
Public Const PATCH_NOTES_LOC As String = WEBSITE & "/projects/workshop/notes.txt"
Public Const UPDATE_PATCH_LOC As String = WEBSITE & "/projects/workshop/update.bin"
Public Const WORKSHOP_PATCH_LOC As String = WEBSITE & "/projects/workshop/patch.bin"
Public Const WORKSHOP_USAGE As String = WEBSITE & "/projects/workshop/workshop_patch.cfm"
Public Const WORKSHOP_AGENT As String = "WORKSHOP"

Public Const LISTINDEX_NOT_FOUND As Long = -99

Public Const WS_NEW_TOON As Long = 1
Public Const WS_CUR_TOON As Long = 0

Public Const WS_NORM_HMAX As Long = 9600
Public Const WS_LUNA_HMAX As Long = 9600

Public Const WS_MAX_WIDTH As Long = 16100

Public Const WS_FRAME_LEFT As Long = 4260
Public Const WS_FRAME_TOP As Long = 4530

Public Const WS_CBFRAME_LEFT As Long = 90
Public Const WS_CBFRAME_TOP As Long = 7920

Public Const WS_DOLL_TOP As Long = 0
Public Const WS_DOLL_LEFT As Long = 0

Public Const WS_DOLL_WIDTH As Long = 2880
Public Const WS_DOLL_HEIGHT As Long = 3600

Public Const WS_DOLL_ICON_UP As Long = 0
Public Const WS_DOLL_ICON_DN As Long = 1

Public Const WS_DOLL_CHEST As Long = 0
Public Const WS_DOLL_ARMS As Long = 1
Public Const WS_DOLL_GEM  As Long = 2
Public Const WS_DOLL_LRING   As Long = 3
Public Const WS_DOLL_LWRIST  As Long = 4
Public Const WS_DOLL_LEGS  As Long = 5
Public Const WS_DOLL_RHAND  As Long = 6
Public Const WS_DOLL_LHAND  As Long = 7
Public Const WS_DOLL_2HAND  As Long = 8
Public Const WS_DOLL_RANGED  As Long = 9
Public Const WS_DOLL_FEET  As Long = 10
Public Const WS_DOLL_RWRIST  As Long = 11
Public Const WS_DOLL_RRING  As Long = 12
Public Const WS_DOLL_WAIST  As Long = 13
Public Const WS_DOLL_HANDS  As Long = 14
Public Const WS_DOLL_HEAD  As Long = 15
Public Const WS_DOLL_CLOAK  As Long = 16
Public Const WS_DOLL_NECK  As Long = 17
Public Const WS_DOLL_RIGHTSPARE  As Long = 18
Public Const WS_DOLL_LEFTSPARE  As Long = 19
Public Const WS_DOLL_2HANDSPARE  As Long = 20
Public Const WS_DOLL_RANGEDSPARE  As Long = 21
Public Const WS_DOLL_MYTHICAL As Long = 22

'realm ability info label indicies
Public Const WS_RAI_TITLE As Long = 0
Public Const WS_RAI_NEXT_RL As Long = 1
Public Const WS_RAI_NEXT_RR As Long = 2
Public Const WS_RAI_RSP_TOTAL As Long = 3
Public Const WS_RAI_RSP_REMAINING As Long = 4

'spellcraft attribute label indicies
Public Const WS_ATTR_STR As Long = 0
Public Const WS_ATTR_CON As Long = 1
Public Const WS_ATTR_DEX As Long = 2
Public Const WS_ATTR_QUI As Long = 3
Public Const WS_ATTR_INT As Long = 4
Public Const WS_ATTR_EMP As Long = 5
Public Const WS_ATTR_PIE As Long = 6
Public Const WS_ATTR_CHA As Long = 7
Public Const WS_ATTR_HIT As Long = 8
Public Const WS_ATTR_POW As Long = 9

'spellcraft resist label indicies
Public Const WS_ATTR_CRUSH As Long = 10
Public Const WS_ATTR_SLASH As Long = 11
Public Const WS_ATTR_THRUST As Long = 12
Public Const WS_ATTR_HEAT As Long = 13
Public Const WS_ATTR_COLD As Long = 14
Public Const WS_ATTR_MATTER As Long = 15
Public Const WS_ATTR_BODY As Long = 16
Public Const WS_ATTR_SPIRIT As Long = 17
Public Const WS_ATTR_ENERGY As Long = 18
Public Const WS_ATTR_ESSENCE As Long = 19

Public Const SP_VS_PROGRESS_CHANGE As Long = 10
Public Const SP_SM_PROGRESS_CHANGE As Long = 25
Public Const SP_MD_PROGRESS_CHANGE As Long = 75
Public Const SP_LG_PROGRESS_CHANGE As Long = 100

Public Const WS_TIMER_LOAD As Long = 0
Public Const WS_TIMER_UNLOAD As Long = 1

Public Const RA_MAX_REALM_RANK = 13

Private Const OC_MESSAGE_MAX As Long = 8
Private Const OC_MESSAGE_MIN As Long = 1

Public Function InitCap(sBuffer As String) As String

    Dim lCtr As Long
    Dim sTemp As String
   
    If Len(sBuffer) <> 0 Then
        
        sBuffer = UCase$(Left$(sBuffer, 1)) & Right$(sBuffer, Len(sBuffer) - 1)
        
        sBuffer = Trim$(sBuffer)
        lCtr = 1
        While lCtr <> Len(sBuffer)
            If Mid$(sBuffer, lCtr, 1) = Chr$(32) Then
                sTemp = Left$(sBuffer, lCtr) & UCase$(Mid$(sBuffer, (lCtr + 1), 1)) & Right$(sBuffer, (Len(sBuffer) - lCtr) - 1)
                sBuffer = sTemp
            End If
            
            lCtr = lCtr + 1
        Wend
          
    End If

    InitCap = sBuffer
    
End Function

Public Function StripTag(sBuffer As String, sStartTag As String, sEndTag As String, Optional Count As Long) As String
'strips tags including the data they contain from an XML file

    Dim sTemp1 As String
    Dim sTemp2 As String
    
    Dim sTemp As String
    
    Dim lCnt As Long
    
    sTemp = LCase$(sBuffer)
    
    If InStr(sTemp, sStartTag) <> 0 And InStr(sTemp, sEndTag) <> 0 Then
    
        If Count Then
            For lCnt = 1 To Count
                If InStr(sTemp, sStartTag) <> 0 And InStr(sTemp, sEndTag) <> 0 Then
                    sTemp1 = Mid$(sTemp, 1, InStr(sTemp, sStartTag) - 1)
                    sTemp2 = Mid$(sTemp, InStr(sTemp, sEndTag) + Len(sEndTag))
                    sTemp = Trim$(sTemp1) & Trim$(sTemp2)
                Else
                    Exit For
                End If
            Next lCnt
        Else
            While InStr(sTemp, sStartTag) <> 0 And InStr(sTemp, sEndTag) <> 0
                sTemp1 = Mid$(sTemp, 1, InStr(sTemp, sStartTag) - 1)
                sTemp2 = Mid$(sTemp, InStr(sTemp, sEndTag) + Len(sEndTag))
                sTemp = Trim$(sTemp1) & Trim$(sTemp2)
            Wend
        End If
        
    End If

    StripTag = sTemp
    
End Function

Public Sub WriteError(sMessage As String)

    
    Dim hFile As Long
    Dim lSizeL As Long
    Dim lSizeH As Long
    Dim lResult As Long
    Dim lWrite As Long
    
    hFile = CreateFile(ERROR_LOG, _
                        GENERIC_WRITE Or GENERIC_READ, _
                        FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                        ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    If hFile = -1 Then
        
        hFile = CreateFile(ERROR_LOG, _
                           GENERIC_WRITE Or GENERIC_READ, _
                           FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                           ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
    
        If hFile = -1 Then Exit Sub
        
    End If
    
    lSizeL = GetFileSize(hFile, lSizeH)
    
    lResult = SetFilePointer(hFile, lSizeL, lSizeH, FILE_BEGIN)
        
    lResult = WriteFile(hFile, ByVal sMessage, Len(sMessage), lWrite, ByVal CLng(0))
    lResult = CloseHandle(hFile)
    
End Sub

Public Function GetListindexByString(sSearch As String, cmbDropbox As ComboBox)

    Dim lCtr As Long
    Dim bFound As Boolean
    
    bFound = False
    
    For lCtr = 0 To cmbDropbox.ListCount - 1
    
        If Trim$(LCase$(sSearch)) = Trim$(LCase$(cmbDropbox.list(lCtr))) Then
            bFound = True
            Exit For
        End If
        
    Next lCtr

    If bFound = False Then lCtr = LISTINDEX_NOT_FOUND
    
    GetListindexByString = lCtr
    
End Function

Public Function Trunc(fNumber As Single) As Long

    Dim sBuf As String
    
    sBuf = STR(fNumber)
    
    If InStr(sBuf, ".") <> 0 Then
        sBuf = Replace(sBuf, ".", "<")
    End If
    
    Trunc = Val(sBuf)
    
End Function

Public Function GetOCMessage() As String
    
    Dim r As Long
    
    Randomize Timer
    
    Do While r = 0 Or r > OC_MESSAGE_MAX
    
        r = CLng(OC_MESSAGE_MAX * Rnd()) + 1
        
    Loop
    
    GetOCMessage = Trim$(STR$(r + 100))
    
End Function

Public Sub StayOnTop(fForm As Form)
    
    Dim lResult As Long
     
    lResult = SetWindowPos(fForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    
End Sub

Public Sub SetVersionInfo()

    RUNNING_VERSION = App.Major & "." & App.Minor & "." & App.Revision

End Sub

Public Sub SetApplicationPath()

    UPDATE_APP_LOC = App.Path & "\Update.exe"
    
    WORKSHOP_APP_LOC = App.Path & "\Workshop.exe"
    
    PATCH_BIN_LOC = App.Path & "\patch.bin"
    
End Sub

Public Sub SetAboutBoxInfo()

    ABOUT_BOX_MESSAGE = "Loto's Character Workshop" & vbCrLf & _
            "Version: " & RUNNING_VERSION & vbCrLf & _
            "Release Date: " & RELEASE_DATE
End Sub

Public Sub JumpToWeb(URL As String)

    Dim lResult As Long
    
    lResult = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)
    
End Sub

Public Function RemoveString(sText As String, sQuery As String) As String

'Example: NewString$ = RemoveString(Text$, ":")
Dim Temp1 As String
Dim Temp2 As String

Dim Text As String

    Text = sText
    
    Do While InStr(Text, sQuery) <> 0
        Temp1 = Left(Text, InStr(Text, sQuery) - 1)
        Temp2 = Mid(Text, InStr(Text, sQuery) + Len(sQuery))
        Text = Temp1 & Temp2
    Loop

    RemoveString = Text

End Function

Public Sub LoadRecentHistory()

    Dim i As Long
    Dim sBuffer As String
    
    For i = 1 To MAX_HISTORY_COUNT
    
        sBuffer = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "RECENT_FILE_" & i)
        
        If (sBuffer <> vbNullString) And (Val(sBuffer) <> REG_KEY_NOT_EXISTS) Then
            Load WS.mnuRecentFile(i)
            WS.mnuRecentFile(i).Caption = Mid(sBuffer, InStrRev(sBuffer, "\") + 1)
        Else
            Exit For
        End If
    Next i
    
    If WS.mnuRecentFile.Count > 1 Then WS.mnuRecentFile(0).Visible = False
                
End Sub

Public Sub ClearRecentHistory()

    Dim i As Long
    
    For i = 1 To MAX_HISTORY_COUNT
    
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, STR_VAL, "RECENT_FILE_" & i, vbNullString)
        
    Next i
    
    Call EmptyMenu
    
End Sub

Public Sub PopulateRecentMenu(sPath As String, sFilename As String)

    Dim i As Long
    Dim lCnt As Long
    
    Dim sBuffer As String
    Dim sPathValues(50) As String   'hard cap max
    
    If Len(sPath) <> 0 And Len(sFilename) <> 0 Then
    
        'get the nullchar out!
        sPathValues(0) = RemoveString((sPath & sFilename), Chr$(0))
        
        'populate the new array of paths
        For lCnt = 1 To MAX_HISTORY_COUNT
            
            sBuffer = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "RECENT_FILE_" & lCnt)
            
            If (sBuffer <> sPathValues(0)) And (sBuffer <> vbNullString) And (Val(sBuffer) <> REG_KEY_NOT_EXISTS) Then
                
                For i = 1 To MAX_HISTORY_COUNT
                    If sPathValues(i) = vbNullString Then
                        sPathValues(i) = sBuffer
                        Exit For
                    End If
                Next i
                
            End If
            
        Next lCnt
        
        'clear the recent file menu
        Call EmptyMenu
        
        'repopulate the recent file menu
        For i = 1 To MAX_HISTORY_COUNT
        
            If sPathValues(i - 1) <> vbNullString Then
                Load WS.mnuRecentFile(i)
                WS.mnuRecentFile(i).Caption = Mid(sPathValues(i - 1), InStrRev(sPathValues(i - 1), "\") + 1)
            
                Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, STR_VAL, "RECENT_FILE_" & i, sPathValues(i - 1))
            Else
                Exit For
            End If
            
        Next i
        
        'set the seperator menu visibility
        If WS.mnuRecentFile.Count > 1 Then
            WS.mnuRecentFile(0).Visible = False
        Else
            WS.mnuRecentFile(0).Visible = True
        End If
        
    End If
    
End Sub

Private Sub EmptyMenu() 'Empty the menu completely but leave the divider (created in design-time)-
    
    Dim i As Integer
    
    WS.mnuRecentFile(0).Visible = True      'Make 'parent' menu item visible.
    
    For i = 1 To WS.mnuRecentFile.UBound    'Remove items that were added in runtime
        Unload WS.mnuRecentFile(i)          'But keep the divider that was created in design-time
    Next i
    
End Sub

Public Sub CheckDirectoryStructure(ByRef ProgressBar As Label, ByRef Status As Label)

'App.Path
'    \templates
'    \items
'        \spellcrafted
'        \chest pieces
'        \sleeves
'        \gems
'        \rings
'        \bracers
'        \leggings
'        \weapons
'        \boots
'        \belts
'        \gloves
'        \helms
'        \cloaks
'        \necklaces
    
    Dim lResult As Long
    
    Dim sMsg As String
    
    Dim lpSA As SECURITY_ATTRIBUTES
    
    lpSA.nLength = Len(lpSA)

    sMsg = "Checking Directory Structure..."
    DoEvents
    With Status
        .Caption = sMsg
        
        .Caption = sMsg & "Templates"
        lResult = CreateDirectory(App.Path & "\templates", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items"
        lResult = CreateDirectory(App.Path & "\items", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Chest Pieces"
        lResult = CreateDirectory(App.Path & "\items\chest pieces", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Sleeves"
        lResult = CreateDirectory(App.Path & "\items\sleeves", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Gems"
        lResult = CreateDirectory(App.Path & "\items\gems", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Rings"
        lResult = CreateDirectory(App.Path & "\items\rings", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Bracers"
        lResult = CreateDirectory(App.Path & "\items\bracers", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Belts"
        lResult = CreateDirectory(App.Path & "\items\belts", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Leggings"
        lResult = CreateDirectory(App.Path & "\items\leggings", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Weapons"
        lResult = CreateDirectory(App.Path & "\items\weapons", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Boots"
        lResult = CreateDirectory(App.Path & "\items\boots", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Gloves"
        lResult = CreateDirectory(App.Path & "\items\gloves", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Helms"
        lResult = CreateDirectory(App.Path & "\items\helms", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Cloaks"
        lResult = CreateDirectory(App.Path & "\items\cloaks", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Necklaces"
        lResult = CreateDirectory(App.Path & "\items\necklaces", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        
        .Caption = sMsg & "Items\Mythical"
        lResult = CreateDirectory(App.Path & "\items\mythical", lpSA)
        ProgressBar.Width = ProgressBar.Width + SP_SM_PROGRESS_CHANGE
        .Caption = sMsg & "Complete!"
        DoEvents
    End With    'Status
    
End Sub

Public Sub LoadRegSettings()

    Dim lResult As Long
    Dim sResult As Single
    
    Dim osV As OSVERSIONINFO
    
    'get windows xp theme information
    osV.dwOSVersionInfoSize = Len(osV)
    lResult = GetVersionEx(osV)
    
    WINDOWS_THEME = 0
    WS_MAX_HEIGHT = WS_NORM_HMAX
    
    If (osV.dwMajorVersion = 5) And (osV.dwMinorVersion = 1) Then
        lResult = GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\ThemeManager", "ThemeActive")
        If lResult = 1 Then
            WINDOWS_THEME = 1
            WS_MAX_HEIGHT = WS_LUNA_HMAX
        End If
    End If
    
    TEMPLATE_RECENT_PATH = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "TEMPLATE_RECENT_PATH")
    If Len(TEMPLATE_RECENT_PATH) = 0 Then
        TEMPLATE_RECENT_PATH = App.Path & "\templates\"
        Call SetNewValue(HKEY_LOCAL_MACHINE, _
                         REGKEY, _
                         STR_VAL, _
                         "TEMPLATE_RECENT_PATH", _
                         TEMPLATE_RECENT_PATH)
        
    End If
    
    MAX_HISTORY_COUNT = CLng(GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "MAX_HISTORY_COUNT"))
    If (MAX_HISTORY_COUNT = 0) Or (MAX_HISTORY_COUNT = REG_KEY_NOT_EXISTS) Then
        MAX_HISTORY_COUNT = 5
        Call SetNewValue(HKEY_LOCAL_MACHINE, _
                         REGKEY, _
                         NUM_VAL, _
                         "MAX_HISTORY_COUNT", _
                         STR(MAX_HISTORY_COUNT))
    End If
                         
    SC_SETTINGS.REALM = CLng(GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "REALM"))
    TOON.REALM = SC_SETTINGS.REALM
    
    If SC_SETTINGS.REALM = REG_KEY_NOT_EXISTS Then
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "REALM", STR(0))
        SC_SETTINGS.REALM = 0
    End If

    SC_SETTINGS.BONUS_OPTION = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BONUS_OPT")
    If SC_SETTINGS.BONUS_OPTION = REG_KEY_NOT_EXISTS Then
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BONUS_OPT", STR(0))
        SC_SETTINGS.BONUS_OPTION = 0
    End If
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "SYSTRAY_EXIT")
    Select Case lResult
        Case 0
            EXIT_TO_SYSTRAY = False
        Case Else
            EXIT_TO_SYSTRAY = True
    End Select
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "SYSTRAY_MINI")
    Select Case lResult
        Case 0
            MINIMIZE_TO_SYSTRAY = False
        Case Else
            MINIMIZE_TO_SYSTRAY = True
    End Select
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "AUTO_UPDATE")
    Select Case lResult
        Case 0
            AUTO_UPDATE = False
        Case Else
            AUTO_UPDATE = True
    End Select
    
    'buffs
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_STR")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_STR, SM_LOC_BUFFS) = lResult
        
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_CON")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_CON, SM_LOC_BUFFS) = lResult
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_DEX")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_DEX, SM_LOC_BUFFS) = lResult
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_QUI")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_QUI, SM_LOC_BUFFS) = lResult
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_INT")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_INT, SM_LOC_BUFFS) = lResult
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_EMP")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_EMP, SM_LOC_BUFFS) = lResult
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_PIE")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_PIE, SM_LOC_BUFFS) = lResult
    
    lResult = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_CHA")
    If lResult <= 0 Then lResult = 0
    TOON.STAT_MATRIX(SM_CHA, SM_LOC_BUFFS) = lResult
End Sub

