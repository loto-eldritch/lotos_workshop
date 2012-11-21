Attribute VB_Name = "UpdatePatch"
Option Explicit

Public Sub TrackPatch(sAgent As String)

    On Error Resume Next
    If Connection_Online Then
        Call ReadHTTP(sAgent, WORKSHOP_USAGE, 0)
    End If

End Sub

Public Sub CheckUpdate()

    Dim sPatchInfo As String
    Dim sCurrentVersion As String
    Dim sVerDate As String
    Dim lPatchSize As Long
    
    If Connection_Online Then
        'download patch information file
        sPatchInfo = Trim$(Download_File(PATCH_INFO_LOC))
        
        If (Len(sPatchInfo) > 0) And (InStr(sPatchInfo, "<charcon>") <> 0) And (InStr(sPatchInfo, "</charcon>") <> 0) Then
            
            'get current release date from info file
            sVerDate = Mid(sPatchInfo, _
                            InStr(sPatchInfo, "<release_date>") + 14, _
                            InStr(sPatchInfo, "</release_date>") - InStr(sPatchInfo, "<release_date>") - 14)
            
            'get current version from info file
            sCurrentVersion = Mid(sPatchInfo, _
                            InStr(sPatchInfo, "<version>") + 9, _
                            InStr(sPatchInfo, "</version>") - InStr(sPatchInfo, "<version>") - 9)
            
            'get current patch size
            lPatchSize = Val(Mid(sPatchInfo, _
                            InStr(sPatchInfo, "<size>") + 6, _
                            InStr(sPatchInfo, "</size>") - InStr(sPatchInfo, "<size>") - 6))
            
            'compare update info with current version
            If RUNNING_VERSION <> sCurrentVersion Then UPDATE_AVAILABLE = True
        Else
            UPDATE_AVAILABLE = False
        End If
    End If
    
End Sub

Public Sub AutoUpdate(ByRef ProgressBar As Label, ByRef Status As Label, hwnd As Long)
   
    Dim sPatchInfo As String
    
    Dim hFile As Long
    Dim lRet As Long
    Dim lPos As Long
    
    Dim sCurrentVersion As String
    
    Dim lPatchSize As Long
    Dim sPatchVersionCurrent As String
    Dim sPatchVersionRunning As String
    
    Dim lWrite As Long
    Dim sBuffer As String
    Dim sMessage As String
    
    Dim sVerDate As String
    
    Dim sAgent As String
        
    Status.Caption = "Checking Connection..."
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    If Connection_Online Then
               
        Status.Caption = "Checking for Updates..."
        ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
        
        'download patch information file
        sPatchInfo = Trim$(Download_File(PATCH_INFO_LOC))
        
        'if the site is online and the patch info file was successfully downloaded then we'll do some processing
        If (Len(sPatchInfo) > 0) And (InStr(sPatchInfo, "<charcon>") <> 0) And (InStr(sPatchInfo, "</charcon>") <> 0) Then
            'download switcher update if there is one before getting main update
            If InStr(sPatchInfo, "<patcher>") <> 0 Then
                
                lPos = InStr(sPatchInfo, "<patcher>")
                'get patcher version
                lPos = InStr(lPos, sPatchInfo, "<version>") + 9
                sPatchVersionCurrent = Trim$(Mid(sPatchInfo, lPos, InStr(lPos, sPatchInfo, "</version>") - lPos))
                sPatchVersionRunning = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "PATCHER_VERSION")
                
                If sPatchVersionRunning <> sPatchVersionCurrent Then
                    'get patcher size
                    lPos = InStr(lPos, sPatchInfo, "<size>") + 6
                    lPatchSize = Val(Mid(sPatchInfo, lPos, InStr(lPos, sPatchInfo, "</size>") - lPos))
                    
                    'reset progressbar for new download
                    ProgressBar.Width = 0
                    ProgressBar.BackColor = &HC00000
                    
                    'set download message
                    sMessage = "Downloading Patcher..."
                    'download new patcher
                    sBuffer = Trim$(Download_File(UPDATE_PATCH_LOC, Status, sMessage, ProgressBar, lPatchSize))
                    Status.Caption = "Download Complete!"
                    
                    'delete older switcher
                    Status.Caption = "Updating Patcher..."
                    
                    'check to see if it exists first
                    hFile = CreateFile(UPDATE_APP_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
                    
                    'if it does, kill it, else skip to next step
                    If hFile <> -1 Then
                        lRet = CloseHandle(hFile)
                        Kill UPDATE_APP_LOC
                    End If
                    
                    'create new patcher file
                    hFile = CreateFile(UPDATE_APP_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
            
                    If hFile = -1 Then Exit Sub
                    
                    'write new patcher to disk
                    lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
                    lRet = CloseHandle(hFile)
                    
                    'reset sbuffer
                    sBuffer = vbNullString
                    
                    Status.Caption = "Updating Patcher...Success!"
                    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, STR_VAL, "PATCHER_VERSION", sPatchVersionCurrent)
                    
                End If
                
            End If
            
            'get current release date from info file
            sVerDate = Mid(sPatchInfo, _
                            InStr(sPatchInfo, "<release_date>") + 14, _
                            InStr(sPatchInfo, "</release_date>") - InStr(sPatchInfo, "<release_date>") - 14)
            
            'get current version from info file
            sCurrentVersion = Mid(sPatchInfo, _
                            InStr(sPatchInfo, "<version>") + 9, _
                            InStr(sPatchInfo, "</version>") - InStr(sPatchInfo, "<version>") - 9)
            
            'get current patch size
            lPatchSize = Val(Mid(sPatchInfo, _
                            InStr(sPatchInfo, "<size>") + 6, _
                            InStr(sPatchInfo, "</size>") - InStr(sPatchInfo, "<size>") - 6))
            
            'compare update info with current version
            If RUNNING_VERSION <> sCurrentVersion Then
            
                'set new download message
                sMessage = "Downloading Patch " & RUNNING_VERSION & "->" & sCurrentVersion & " " & sVerDate & ": "
                
                'download new patch
                sBuffer = Trim$(Download_File(WORKSHOP_PATCH_LOC, Status, sMessage, ProgressBar, lPatchSize))
                
                'check to see if there is already a patch.bin file if there is kill it
                hFile = CreateFile(PATCH_BIN_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
                If hFile <> -1 Then
                    lRet = CloseHandle(hFile)
                    Kill PATCH_BIN_LOC
                End If
                
                'create new patch file
                hFile = CreateFile(PATCH_BIN_LOC, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
            
                If hFile = -1 Then Exit Sub
            
                'write new patch file to disk
                lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
                lRet = CloseHandle(hFile)
                            
                'clear buffer
                sBuffer = vbNullString
                
                'set new status caption
                Status.Caption = "Download Complete...Restarting."
                ProgressBar.Width = Val(ProgressBar.Tag)
                'Call Sleep(1000)
                
                lRet = ShellExecute(hwnd, "open", UPDATE_APP_LOC, vbNullString, vbNullString, SW_SHOW)
                End
            Else
                Status.Caption = "No Update Required"
            End If
        Else
            Status.Caption = "Update Server Unavailable"
        End If
    Else
        Status.Caption = "No Internet Connection Found"
    End If
    
End Sub
