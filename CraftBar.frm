VERSION 5.00
Begin VB.Form CraftBar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crafting Quickbar Configuration"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameOptions 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   5655
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   5760
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Quickbar Restoration"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1065
         Index           =   3
         Left            =   90
         TabIndex        =   35
         Top             =   3660
         Width           =   5565
         Begin VB.Label option_Label 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   735
            Index           =   0
            Left            =   90
            TabIndex        =   36
            Top             =   240
            Width           =   5385
         End
      End
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Path to Character Files"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Index           =   2
         Left            =   90
         TabIndex        =   33
         Top             =   4740
         Width           =   5565
         Begin VB.Label lbl_CharacterPath 
            Caption         =   "path not set"
            ForeColor       =   &H00800000&
            Height          =   555
            Left            =   180
            TabIndex        =   34
            Top             =   210
            Width           =   5265
         End
      End
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Select Pieces to Include"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3495
         Index           =   1
         Left            =   3180
         TabIndex        =   30
         Top             =   120
         Width           =   2475
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Chest"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   0
            Left            =   300
            TabIndex        =   3
            Tag             =   "0"
            Top             =   450
            Width           =   795
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Arms"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   1
            Left            =   300
            TabIndex        =   5
            Tag             =   "1"
            Top             =   825
            Width           =   795
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Legs"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   2
            Left            =   300
            TabIndex        =   7
            Tag             =   "5"
            Top             =   1200
            Width           =   795
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Head"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   3
            Left            =   1170
            TabIndex        =   4
            Tag             =   "15"
            Top             =   450
            Width           =   795
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Hands"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   4
            Left            =   1170
            TabIndex        =   6
            Tag             =   "14"
            Top             =   825
            Width           =   795
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Feet"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   5
            Left            =   1170
            TabIndex        =   8
            Tag             =   "10"
            Top             =   1200
            Width           =   795
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "RH"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   6
            Left            =   300
            TabIndex        =   9
            Tag             =   "6"
            Top             =   1860
            Width           =   855
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "LH"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   7
            Left            =   300
            TabIndex        =   11
            Tag             =   "7"
            Top             =   2235
            Width           =   855
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "2H"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   8
            Left            =   300
            TabIndex        =   13
            Tag             =   "8"
            Top             =   2610
            Width           =   855
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Bow"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   9
            Left            =   300
            TabIndex        =   15
            Tag             =   "9"
            Top             =   2985
            Width           =   855
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "RH Spare"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   10
            Left            =   1170
            TabIndex        =   10
            Tag             =   "18"
            Top             =   1860
            Width           =   1095
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "LH Spare"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   11
            Left            =   1170
            TabIndex        =   12
            Tag             =   "19"
            Top             =   2235
            Width           =   1095
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "2H Spare"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   12
            Left            =   1170
            TabIndex        =   14
            Tag             =   "20"
            Top             =   2610
            Width           =   1095
         End
         Begin VB.CheckBox chk_QBarInclude 
            Caption         =   "Bow Spare"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   13
            Left            =   1170
            TabIndex        =   16
            Tag             =   "21"
            Top             =   2985
            Width           =   1095
         End
         Begin VB.Label lbl_Location 
            Caption         =   "Weapons"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   31
            Top             =   1650
            Width           =   690
         End
         Begin VB.Label lbl_Location 
            Caption         =   "Armor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   270
            TabIndex        =   32
            Top             =   240
            Width           =   480
         End
         Begin VB.Shape shape_Options 
            BorderColor     =   &H8000000D&
            Height          =   1275
            Index           =   0
            Left            =   180
            Top             =   330
            Width           =   2115
         End
         Begin VB.Shape shape_Options 
            BorderColor     =   &H8000000D&
            Height          =   1635
            Index           =   1
            Left            =   180
            Top             =   1740
            Width           =   2115
         End
      End
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Configure QuickBars"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3495
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   120
         Width           =   3015
         Begin VB.ComboBox cmbCrafterSelect 
            ForeColor       =   &H00800000&
            Height          =   315
            ItemData        =   "CraftBar.frx":0000
            Left            =   150
            List            =   "CraftBar.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   1530
            Width           =   2715
         End
         Begin VB.CommandButton cmd_LoadGems 
            Caption         =   "&Load Gems"
            Enabled         =   0   'False
            Height          =   325
            Left            =   300
            TabIndex        =   1
            Top             =   2940
            Width           =   1185
         End
         Begin VB.CommandButton cmd_RestoreBars 
            Caption         =   "&Restore"
            Height          =   325
            Left            =   1545
            TabIndex        =   2
            Top             =   2940
            Width           =   1185
         End
         Begin VB.Label option_Label 
            Caption         =   "Select Crafter"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   5
            Left            =   180
            TabIndex        =   21
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Label option_Label 
            Caption         =   "QBar:"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   420
            TabIndex        =   29
            Top             =   570
            Width           =   525
         End
         Begin VB.Label lbl_QBar 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   28
            Top             =   570
            Width           =   525
         End
         Begin VB.Label option_Label 
            Caption         =   "Page:"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   420
            TabIndex        =   27
            Top             =   915
            Width           =   525
         End
         Begin VB.Label lbl_QPage 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   26
            Top             =   915
            Width           =   525
         End
         Begin VB.Image cmd_QBar 
            Height          =   135
            Index           =   0
            Left            =   1560
            Top             =   570
            Width           =   285
         End
         Begin VB.Image cmd_QBar 
            Height          =   135
            Index           =   1
            Left            =   1560
            Top             =   705
            Width           =   285
         End
         Begin VB.Image cmd_QPage 
            Height          =   135
            Index           =   0
            Left            =   1560
            Top             =   900
            Width           =   285
         End
         Begin VB.Image cmd_QPage 
            Height          =   135
            Index           =   1
            Left            =   1560
            Top             =   1035
            Width           =   285
         End
         Begin VB.Label option_Label 
            Alignment       =   2  'Center
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   960
            TabIndex        =   25
            Top             =   300
            Width           =   525
         End
         Begin VB.Label option_Label 
            Alignment       =   2  'Center
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   1920
            TabIndex        =   24
            Top             =   300
            Width           =   525
         End
         Begin VB.Label lbl_QBar 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   23
            Top             =   570
            Width           =   525
         End
         Begin VB.Label lbl_QPage 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   22
            Top             =   915
            Width           =   525
         End
         Begin VB.Shape shape_Options 
            BorderColor     =   &H8000000D&
            Height          =   1995
            Index           =   2
            Left            =   90
            Top             =   1350
            Width           =   2835
         End
         Begin VB.Label option_Label 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   705
            Index           =   6
            Left            =   210
            TabIndex        =   20
            Top             =   1950
            Width           =   2595
         End
      End
   End
   Begin VB.FileListBox file_Toons 
      Height          =   3210
      Left            =   5910
      Pattern         =   "*.ini"
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   30
      Width           =   1935
   End
End
Attribute VB_Name = "CraftBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private GEM_AUTO_LOAD_SELECTIONS As Long

Private Sub cmd_RestoreBars_Click()

    Dim hFile As Long
      
    Dim sCharacter As String
    Dim sServer As String
    
    Dim sPath As String
    Dim sBuf As String * 1024
    Dim sBuffer As String
    
    Dim sTemp As String
    
    Dim lRet As Long
    Dim lRead As Long
    Dim lWrite As Long
    
    sCharacter = Left$(cmbCrafterSelect.Text, InStr(cmbCrafterSelect.Text, ",") - 1)
    sServer = Right$(cmbCrafterSelect.Text, Len(cmbCrafterSelect.Text) - InStr(cmbCrafterSelect.Text, ","))
    
    sPath = lbl_CharacterPath.Caption & "\" & sCharacter & "-" & GetServerCode(sServer) & ".ws-ini"
    
    If sPath <> vbNullString Then
        
        hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, _
                           FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                           ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
            
        If hFile = -1 Then
            MsgBox "The selected character does not have a backup configuration to restore from!", vbCritical, "QuickBar Restoration"
        Else
        
            Do
                sBuf = vbNullString
                    
                lRet = ReadFile(hFile, ByVal sBuf, Len(sBuf), lRead, ByVal CLng(0))
                'exit if there was no more data to read
                If lRead = 0 Then Exit Do
                
                If sBuf <> vbNullString Then sBuffer = sBuffer & sBuf
            Loop
            
            'close the backup and delete it. we have it in a buffer now
            lRet = CloseHandle(hFile)
            lRet = DeleteFile(sPath)
                        
            sBuffer = Trim$(sBuffer)
            'now we have the character ini file in a buffer let's play with it
            
            'now we create the new character ini
            sPath = lbl_CharacterPath.Caption & "\" & sCharacter & "-" & GetServerCode(sServer) & ".ini"
                                    
            hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, _
                                FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                                ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
            If hFile = -1 Then
                MsgBox "Error - Couldn't Write Config!", vbCritical, "Quick Bar Error"
            Else
                lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
                If lWrite <> 0 Then
                    lRet = CloseHandle(hFile)
                    MsgBox "Your Quickbar has been restored successfully!", vbInformation, "Quick Bar Loaded!"
                Else
                    MsgBox "Your Quickbars could not be restored!", vbCritical, "QuickBar Error"
                End If
            End If
        End If
    End If

End Sub

Private Sub Form_Load()

    Dim sCharacterPath As String
    
    CraftBar.Width = 5840
    CraftBar.Height = 6000
    
    'temp code for quick tests (remove before compile)
    Call Init_ServerCodes
    
    Set cmd_QBar(0).Picture = LoadResPicture("SUP_UP", vbResBitmap)
    Set cmd_QBar(1).Picture = LoadResPicture("SDOWN_UP", vbResBitmap)
    
    Set cmd_QPage(0).Picture = LoadResPicture("SUP_UP", vbResBitmap)
    Set cmd_QPage(1).Picture = LoadResPicture("SDOWN_UP", vbResBitmap)
    
    sCharacterPath = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "CHARACTER_PATH")
        
    If Val(sCharacterPath) <> REG_KEY_NOT_EXISTS Then
        lbl_CharacterPath.Caption = sCharacterPath
        file_Toons.Path = sCharacterPath
        file_Toons.Pattern = "*.ini"
        Call LoadSpellcrafters
        cmd_LoadGems.Enabled = True
    Else
        file_Toons.Pattern = "*.ASRFIHASDF"
        lbl_CharacterPath.Caption = "Path to character files is not set. Click here to select."
        cmd_LoadGems.Enabled = False
    End If
    
    option_Label(6).Caption = "You must be logged out of the game, or at the character selection screen, before attempting to load gems or restore your quickbars!"
    option_Label(0).Caption = "After you have finished crafting the gems loaded to your quickbars, if you would like to return your previous quickbar configuration for normal game play, simply select the character from the box above and click Restore."
    
End Sub

Private Sub LoadSpellcrafters()

    Dim lCnt As Long
    Dim sServer As String
    Dim sCharacter As String
       
    Dim lRet As Long
    Dim hFile As Long
    Dim lRead As Long
    Dim sBuf As String * 1000
    Dim sBuffer As String
    
    Dim isCrafter As Boolean
    
    cmbCrafterSelect.Clear
    
    If file_Toons.ListCount > 0 Then
        'now that we have the path defined let's add crafters to the combo box cmbCrafterSelect
        For lCnt = 0 To file_Toons.ListCount - 1
            If InStr(file_Toons.list(lCnt), "-") Then sCharacter = Left$(file_Toons.list(lCnt), InStr(file_Toons.list(lCnt), "-") - 1)
            sServer = GetServerName(Val(Mid$(file_Toons.list(lCnt), InStr(file_Toons.list(lCnt), "-") + 1, InStr(file_Toons.list(lCnt), ".") - InStr(file_Toons.list(lCnt), "-") - 1)))
            
            'open file and search for the code for the spellcraft icon
            hFile = CreateFile(file_Toons.Path & "\" & file_Toons.list(lCnt), GENERIC_WRITE Or GENERIC_READ, _
                FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), _
                OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    
            If hFile = -1 Then Exit Sub 'If hFile is -1 the file is not there and there has been an error
            sBuffer = vbNullString
            Do
                sBuf = vbNullString
                lRet = ReadFile(hFile, ByVal sBuf, Len(sBuf), lRead, ByVal CLng(0))   'Read chars into sBuf
                sBuffer = sBuffer & RemoveString(sBuf, vbNullChar) 'strip the UTF-16 bullshit
            Loop Until lRead = 0
            lRet = CloseHandle(hFile)       'Close the file it's not needed anymore
        
            'trim start/end spacing off of buffer
            sBuffer = LCase$(Trim$(sBuffer))
            If InStr(sBuffer, "=44,13,,") Then
                isCrafter = True
            Else
                isCrafter = False
            End If
            
            If sCharacter <> vbNullString And sServer <> vbNullString And isCrafter Then
                cmbCrafterSelect.AddItem sCharacter & "," & sServer
            End If
            
        Next lCnt
        
        If cmbCrafterSelect.ListCount > 0 Then cmbCrafterSelect.ListIndex = 0
    End If
    
End Sub

Private Sub cmd_LoadGems_Click()

    Dim hFile As Long
    
    Dim bSplit As Boolean
    
    Dim sCharacter As String
    Dim sServer As String
    
    Dim sPath As String
    Dim sBuf As String * 1024
    Dim sBuffer As String
    Dim sBufferHI As String
    Dim sBufferLO As String
    
    Dim sQuickBar As String
    Dim sQuickBarHI As String
    Dim sQuickBarLO As String
    
    Dim sTemp As String
    
    Dim lCnt As Long
    Dim lPos As Long
    
    Dim lRet As Long
    Dim lRead As Long
    Dim lWrite As Long
    
    Dim lQBar As Long
    Dim lQPage As Long
    Dim lQSlot As Long
    
    If GEM_AUTO_LOAD_SELECTIONS > 0 Then
    
        If InStr(cmbCrafterSelect.Text, ",") <> 0 Then
            sCharacter = Left$(cmbCrafterSelect.Text, InStr(cmbCrafterSelect.Text, ",") - 1)
            sServer = Right$(cmbCrafterSelect.Text, Len(cmbCrafterSelect.Text) - InStr(cmbCrafterSelect.Text, ","))
                
            'set the path to the character config backup for the selected character and server combo
            sPath = lbl_CharacterPath.Caption & "\" & sCharacter & "-" & GetServerCode(sServer) & ".ws-ini"
            
            If sPath <> vbNullString Then
                'check for the existence of the config backup file before proceeding any further
                hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, _
                                    FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                                    ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
            Else
                MsgBox "File path could not be created, please define the path to your character configuration files.", vbCritical, "File Path Error"
                Exit Sub
            End If
            'close the file that we opened above
            lRet = CloseHandle(hFile)
            
            If hFile = -1 Then
                'if the backup file doesn't exist then let's go ahead and set up the bars!
                
                sPath = lbl_CharacterPath.Caption & "\" & sCharacter & "-" & GetServerCode(sServer) & ".ini"
                If sPath <> vbNullString Then
                
                    hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, _
                                       FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                                       ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
                    
                    If hFile = -1 Then Exit Sub
                            
                    Do
                        sBuf = vbNullString
                        
                        lRet = ReadFile(hFile, ByVal sBuf, Len(sBuf), lRead, ByVal CLng(0))
                        'exit if there was no more data to read
                        If lRead = 0 Then Exit Do
                        
                        If sBuf <> vbNullString Then sBuffer = sBuffer & sBuf
                    Loop
                    
                    lRet = CloseHandle(hFile)
                    
                    sBuffer = Trim$(sBuffer)
                    'now we have the character ini file in a buffer let's play with it
                    
                    lQBar = Val(lbl_QBar(0).Caption)
                    lQPage = (Val(lbl_QPage(0).Caption) - 1) * 10
                    lQSlot = 0
                    
                    If InStr(sBuffer, "[Quickbar") Then
                        Select Case lQBar
                            Case 1
                                sBufferHI = Left$(sBuffer, InStr(sBuffer, "[Quickbar]") - 1)
                                sQuickBar = Mid$(sBuffer, InStr(sBuffer, "[Quickbar]"), InStr(InStr(sBuffer, "[Quickbar]") + 10, sBuffer, "[") - InStr(sBuffer, "[Quickbar]")) & vbCrLf
                                sBufferLO = Mid$(sBuffer, InStr(sBuffer, "[Quickbar]") - 2 + Len(sQuickBar))
                            Case 2
                                sBufferHI = Left$(sBuffer, InStr(sBuffer, "[Quickbar2]") - 1)
                                sQuickBar = Mid$(sBuffer, InStr(sBuffer, "[Quickbar2]"), InStr(InStr(sBuffer, "[Quickbar2]") + 11, sBuffer, "[") - InStr(sBuffer, "[Quickbar2]")) & vbCrLf
                                sBufferLO = Mid$(sBuffer, InStr(sBuffer, "[Quickbar2]") - 2 + Len(sQuickBar))
                            Case 3
                                sBufferHI = Left$(sBuffer, InStr(sBuffer, "[Quickbar3]") - 1)
                                sQuickBar = Mid$(sBuffer, InStr(sBuffer, "[Quickbar3]"), InStr(InStr(sBuffer, "[Quickbar3]") + 11, sBuffer, "[") - InStr(sBuffer, "[Quickbar3]")) & vbCrLf
                                sBufferLO = Mid$(sBuffer, InStr(sBuffer, "[Quickbar3]") - 2 + Len(sQuickBar))
                        End Select
                        
                        For lCnt = 0 To chk_QBarInclude.UBound
                            If chk_QBarInclude(lCnt).Value = vbChecked Then
                                
                                '==================================================================================================================
                                If lQSlot > 9 Then
                                    lQSlot = 0
                                    lQPage = (lQPage + 10)
                                End If
                                
                                sTemp = "Hotkey_" & lQPage + lQSlot & "="
                                bSplit = False
                                
                                'check if the new hotkey is already defined in the quickbar
                                If InStr(sQuickBar, sTemp) Then
                                    'remove the hotkey from the buffer by splitting and rejoining
                                    sQuickBarHI = Left$(sQuickBar, InStr(sQuickBar, sTemp) - 1)
                                    lPos = InStr(sQuickBar, sTemp) + 1
                                    If InStr(lPos, sQuickBar, "Hotkey") <> 0 Then
                                        sQuickBarLO = Mid$(sQuickBar, InStr(lPos, sQuickBar, "Hotkey"))
                                    Else
                                        sQuickBarLO = vbNullString
                                    End If
                                    bSplit = True
                                End If
                                                        
                                'finish constructing the hotkey with the gemcode and such
                                If WS.lbl_GemNameSC1(Val(chk_QBarInclude(lCnt).Tag)).Caption <> vbNullString Then
                                    sTemp = sTemp & "45," & _
                                        GetGemCode(TOON.REALM, WS.lbl_GemNameSC1(Val(chk_QBarInclude(lCnt).Tag)).Caption) & _
                                        ",,-1" & vbCrLf
                                End If
                                                        
                                'if we had to split it to remove the existing hotkey, rejoin with the new key, else tack the new key onto the end
                                'this probably won't work right. may need to search for the slot-1 to get a position. or i could just
                                'blow away the whole bar and they'll just have to restore after crafting.
                                If bSplit Then
                                    sQuickBar = sQuickBarHI & sTemp & sQuickBarLO
                                Else
                                    sQuickBar = sQuickBar & vbCrLf & sTemp
                                End If
                                lQSlot = lQSlot + 1
                                
                                '==================================================================================================================
                                If lQSlot > 9 Then
                                    lQSlot = 0
                                    lQPage = (lQPage + 10)
                                End If
                                
                                sTemp = "Hotkey_" & lQPage + lQSlot & "="
                                bSplit = False
                                
                                'check if the new hotkey is already defined in the quickbar
                                If InStr(sQuickBar, sTemp) Then
                                    'remove the hotkey from the buffer by splitting and rejoining
                                    sQuickBarHI = Left$(sQuickBar, InStr(sQuickBar, sTemp) - 1)
                                    lPos = InStr(sQuickBar, sTemp) + 1
                                    If InStr(lPos, sQuickBar, "Hotkey") <> 0 Then
                                        sQuickBarLO = Mid$(sQuickBar, InStr(lPos, sQuickBar, "Hotkey"))
                                    Else
                                        sQuickBarLO = vbNullString
                                    End If
                                    bSplit = True
                                End If
                                
                                'finish constructing the hotkey with the gemcode and such
                                If WS.lbl_GemNameSC2(Val(chk_QBarInclude(lCnt).Tag)).Caption <> vbNullString Then
                                    sTemp = sTemp & "45," & _
                                        GetGemCode(TOON.REALM, WS.lbl_GemNameSC2(Val(chk_QBarInclude(lCnt).Tag)).Caption) & _
                                        ",,-1" & vbCrLf
                                End If
                                
                                'if we had to split it to remove the existing hotkey, rejoin with the new key, else tack the new key onto the end
                                'this probably won't work right. may need to search for the slot-1 to get a position. or i could just
                                'blow away the whole bar and they'll just have to restore after crafting.
                                If bSplit Then
                                    sQuickBar = sQuickBarHI & sTemp & sQuickBarLO
                                Else
                                    sQuickBar = sQuickBar & vbCrLf & sTemp
                                End If
                                lQSlot = lQSlot + 1
                                
                                '==================================================================================================================
                                If lQSlot > 9 Then
                                    lQSlot = 0
                                    lQPage = (lQPage + 10)
                                End If
                                
                                sTemp = "Hotkey_" & lQPage + lQSlot & "="
                                bSplit = False
                                
                                'check if the new hotkey is already defined in the quickbar
                                If InStr(sQuickBar, sTemp) Then
                                    'remove the hotkey from the buffer by splitting and rejoining
                                    sQuickBarHI = Left$(sQuickBar, InStr(sQuickBar, sTemp) - 1)
                                    lPos = InStr(sQuickBar, sTemp) + 1
                                    If InStr(lPos, sQuickBar, "Hotkey") <> 0 Then
                                        sQuickBarLO = Mid$(sQuickBar, InStr(lPos, sQuickBar, "Hotkey"))
                                    Else
                                        sQuickBarLO = vbNullString
                                    End If
                                    bSplit = True
                                End If
                                
                                If WS.lbl_GemNameSC3(Val(chk_QBarInclude(lCnt).Tag)).Caption <> vbNullString Then
                                    sTemp = sTemp & "45," & _
                                        GetGemCode(TOON.REALM, WS.lbl_GemNameSC3(Val(chk_QBarInclude(lCnt).Tag)).Caption) & _
                                        ",,-1" & vbCrLf
                                End If
                                
                                'if we had to split it to remove the existing hotkey, rejoin with the new key, else tack the new key onto the end
                                'this probably won't work right. may need to search for the slot-1 to get a position. or i could just
                                'blow away the whole bar and they'll just have to restore after crafting.
                                If bSplit Then
                                    sQuickBar = sQuickBarHI & sTemp & sQuickBarLO
                                Else
                                    sQuickBar = sQuickBar & vbCrLf & sTemp
                                End If
                                lQSlot = lQSlot + 1
                                
                                '==================================================================================================================
                                If lQSlot > 9 Then
                                    lQSlot = 0
                                    lQPage = (lQPage + 10)
                                End If
                                
                                sTemp = "Hotkey_" & lQPage + lQSlot & "="
                                bSplit = False
                                
                                'check if the new hotkey is already defined in the quickbar
                                If InStr(sQuickBar, sTemp) Then
                                    'remove the hotkey from the buffer by splitting and rejoining
                                    sQuickBarHI = Left$(sQuickBar, InStr(sQuickBar, sTemp) - 1)
                                    lPos = InStr(sQuickBar, sTemp) + 1
                                    If InStr(lPos, sQuickBar, "Hotkey") <> 0 Then
                                        sQuickBarLO = Mid$(sQuickBar, InStr(lPos, sQuickBar, "Hotkey"))
                                    Else
                                        sQuickBarLO = vbNullString
                                    End If
                                    bSplit = True
                                End If
                                                        
                                If WS.lbl_GemNameSC4(Val(chk_QBarInclude(lCnt).Tag)).Caption <> vbNullString Then
                                    sTemp = sTemp & "45," & _
                                        GetGemCode(TOON.REALM, WS.lbl_GemNameSC4(Val(chk_QBarInclude(lCnt).Tag)).Caption) & _
                                        ",,-1" & vbCrLf
                                End If
                               
                                'if we had to split it to remove the existing hotkey, rejoin with the new key, else tack the new key onto the end
                                'this probably won't work right. may need to search for the slot-1 to get a position. or i could just
                                'blow away the whole bar and they'll just have to restore after crafting.
                                If bSplit Then
                                    sQuickBar = sQuickBarHI & sTemp & sQuickBarLO
                                Else
                                    sQuickBar = sQuickBar & vbCrLf & sTemp
                                End If
                                lQSlot = lQSlot + 1
                                
                                '==================================================================================================================
                                'make a blank slot
                                If lQSlot > 9 Then
                                    lQSlot = 0
                                    lQPage = (lQPage + 10)
                                End If
                                sTemp = "Hotkey_" & lQPage + lQSlot & "="
                                bSplit = False
                                
                                'check if the new hotkey is already defined in the quickbar
                                If InStr(sQuickBar, sTemp) Then
                                    'remove the hotkey from the buffer by splitting and rejoining
                                    sQuickBarHI = Left$(sQuickBar, InStr(sQuickBar, sTemp) - 1)
                                    lPos = InStr(sQuickBar, sTemp) + 1
                                    If InStr(lPos, sQuickBar, "Hotkey") <> 0 Then
                                        sQuickBarLO = Mid$(sQuickBar, InStr(lPos, sQuickBar, "Hotkey"))
                                    Else
                                        sQuickBarLO = vbNullString
                                    End If
                                    bSplit = True
                                End If
                                
                                If bSplit Then sQuickBar = sQuickBarHI & sQuickBarLO
                                'set the slot number + 1 for the next round
                                lQSlot = lQSlot + 1
                                
                            End If
                        Next lCnt
                        
                        'we've completed construction of the new quickbar we need to put it all back together now
                        sTemp = sBufferHI & sQuickBar & sBufferLO
                        
                        'first let's check for the existence of a backup
                        sPath = lbl_CharacterPath.Caption & "\" & sCharacter & "-" & GetServerCode(sServer) & ".ws-ini"
                        
                        hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, _
                                       FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                                       ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
                    
                        If hFile = -1 Then
                            'it doesn't exist so we need to create it
                            lRet = CloseHandle(hFile)
                            hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, _
                                       FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                                       ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
                            
                            lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
                        End If
                        'close the file handle whether the backup was created or already existed.
                        lRet = CloseHandle(hFile)
                                       
                        'now we create the new character ini
                        sPath = lbl_CharacterPath.Caption & "\" & sCharacter & "-" & GetServerCode(sServer) & ".ini"
                        
                        hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, _
                                       FILE_SHARE_WRITE Or FILE_SHARE_READ, _
                                       ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
                        
                        If hFile = -1 Then
                            MsgBox "Error - Couldn't Write Config!", vbCritical, "Craft Bar Error"
                        Else
                            lRet = WriteFile(hFile, ByVal sTemp, Len(sTemp), lWrite, ByVal CLng(0))
                            lRet = CloseHandle(hFile)
                            MsgBox "Your crafting quickbar has been loaded successfully!", vbInformation, "Craft Bar Loaded!"
                        End If
                    End If
                End If
            Else
                MsgBox "You already have a set of crafting bars loaded. To avoid trashing your quickbars please RESTORE your original quickbars then load your next set!", vbCritical, "Craft Bar Error"
            End If
        End If
    Else
        MsgBox "You must select at least one item to spellcraft!", vbCritical, "Craft Bar Error!"
    End If
    
End Sub

Private Sub chk_QBarInclude_Click(Index As Integer)

    Dim lCnt As Integer
    
    GEM_AUTO_LOAD_SELECTIONS = 0
    
    For lCnt = 0 To chk_QBarInclude.UBound
        
        If chk_QBarInclude(lCnt).Value = vbChecked Then
            GEM_AUTO_LOAD_SELECTIONS = GEM_AUTO_LOAD_SELECTIONS + 5
        End If
        
    Next lCnt
    
    If GEM_AUTO_LOAD_SELECTIONS > 0 Then
        lbl_QPage(1).Caption = Val(lbl_QPage(0).Caption) + (GEM_AUTO_LOAD_SELECTIONS / 10) - 1
    Else
        lbl_QPage(1).Caption = lbl_QPage(0).Caption
    End If
    
End Sub

Private Sub cmd_QBar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Set cmd_QBar(Index).Picture = LoadResPicture("SUP_DOWN", vbResBitmap)
    Else
        Set cmd_QBar(Index).Picture = LoadResPicture("SDOWN_DOWN", vbResBitmap)
    End If
End Sub

Private Sub cmd_QBar_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Set cmd_QBar(Index).Picture = LoadResPicture("SUP_UP", vbResBitmap)
        If Val(lbl_QBar(0).Caption) <> 3 Then lbl_QBar(0).Caption = Val(lbl_QBar(0).Caption) + 1
    Else
        Set cmd_QBar(Index).Picture = LoadResPicture("SDOWN_UP", vbResBitmap)
        If Val(lbl_QBar(0).Caption) <> 1 Then lbl_QBar(0).Caption = Val(lbl_QBar(0).Caption) - 1
    End If
End Sub

Private Sub cmd_QPage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Set cmd_QPage(Index).Picture = LoadResPicture("SUP_DOWN", vbResBitmap)
    Else
        Set cmd_QPage(Index).Picture = LoadResPicture("SDOWN_DOWN", vbResBitmap)
    End If
End Sub

Private Sub cmd_QPage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 0 Then
        Set cmd_QPage(Index).Picture = LoadResPicture("SUP_UP", vbResBitmap)
        If Val(lbl_QPage(0).Caption) <> 10 Then lbl_QPage(0).Caption = Val(lbl_QPage(0).Caption) + 1
        Call chk_QBarInclude_Click(0)
    Else
        Set cmd_QPage(Index).Picture = LoadResPicture("SDOWN_UP", vbResBitmap)
        If Val(lbl_QPage(0).Caption) <> 1 Then lbl_QPage(0).Caption = Val(lbl_QPage(0).Caption) - 1
        Call chk_QBarInclude_Click(0)
    End If
End Sub

Private Sub lbl_QBar_Change(Index As Integer)

    lbl_QBar(1).Caption = lbl_QBar(0).Caption
    
End Sub

Private Sub lbl_QPage_Change(Index As Integer)

    If Index = 1 Then
        If Val(lbl_QPage(Index).Caption) > 10 Then
            lbl_QPage(Index).Caption = "ERR"
            lbl_QPage(Index).ForeColor = &HFF&
        Else
            lbl_QPage(Index).ForeColor = &H800000
        End If
    End If
    
End Sub

Private Sub lbl_CharacterPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbl_CharacterPath.ForeColor = &HFF8080
    
End Sub

Private Sub lbl_CharacterPath_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbl_CharacterPath.ForeColor = &H800000
    
    Dim lRet As Long
    Dim cmdFlags As Long
    
    Dim sPath As String
    Dim cmdFilter As String
    Dim cmdMessage As String
        
    cmdFlags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    cmdFilter = "Character Setting Files (*.ini)" & vbNullChar & "*.ini" & vbNullChar
    cmdMessage = "Select Path for Character Files..."
        
    sPath = CMD_OpenSave(lOpen, Me.hwnd, cmdFilter, 1, "", cmdMessage, cmdFlags)
    
    If (sPath <> vbNullString) And (InStr(sPath, ".ini")) Then
        'strip down the path
        sPath = Left$(sPath, InStrRev(sPath, "\") - 1)
        'assign path to the label for public viewing and to the file box behind the scenes
        lbl_CharacterPath.Caption = sPath
        file_Toons.Path = sPath
        file_Toons.Pattern = "*.ini"
        
        Call LoadSpellcrafters
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, STR_VAL, "CHARACTER_PATH", sPath)
        cmd_LoadGems.Enabled = True
    End If
      
End Sub
