VERSION 5.00
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Loto's Character Workshop"
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8670
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_Version 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   7470
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label ProgressBar 
      BackColor       =   &H00FF8080&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Tag             =   "8610"
      Top             =   2025
      Width           =   8610
   End
   Begin VB.Label Status 
      BackColor       =   &H00000000&
      Caption         =   "Initializing..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   50
      TabIndex        =   1
      Top             =   2280
      Width           =   8670
   End
   Begin VB.Image Image_Splash 
      Height          =   1980
      Left            =   0
      Picture         =   "Splash.frx":0000
      Top             =   0
      Width           =   8670
   End
   Begin VB.Label lbl_ProgressBG 
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1980
      Width           =   8715
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   
    ERROR_LOG = App.Path & "\error_log.txt"
    
    Dim lProgress As Long
    
    Dim i As Long
    
    SPLASH_LOADING = True
    
    ProgressBar.Width = 0
    
    Me.Show
    
    '----------------------------------------------------------------------------------------------
    Status.Caption = "Checking Version Information..."
    
    Call SetVersionInfo
    Call SetAboutBoxInfo
    Call SetApplicationPath
    
    lbl_Version.Caption = RUNNING_VERSION
    
    '----------------------------------------------------------------------------------------------
    'track usage
    Call TrackPatch(WORKSHOP_AGENT & "_START")
    
    '----------------------------------------------------------------------------------------------
    'load registry settings
    Status.Caption = "Loading Registry Settings..."
    DoEvents
    Call LoadRegSettings
    
    '----------------------------------------------------------------------------------------------
    'check for program updates
    'If AUTO_UPDATE Then
    '    Call AutoUpdate(ProgressBar, Status, Me.hwnd)
    'Else
    '    Call CheckUpdate
    'End If

    '----------------------------------------------------------------------------------------------
    'reset the progress bar for the rest of the program Initialization
    ProgressBar.BackColor = &HFF8080
    ProgressBar.Width = (SP_LG_PROGRESS_CHANGE * 2)
    
    '----------------------------------------------------------------------------------------------
    'check directory structure
    Call CheckDirectoryStructure(ProgressBar, Status)
  
    '----------------------------------------------------------------------------------------------
    'Init charcon arrays
    Call InitCharacterArrays(ProgressBar, Status)
    
    'Init realmability arrays
    Call InitRealmAbilityArrays(ProgressBar, Status)
    
    'Init spellcraft arrays
    Call InitSpellcraftArrays(ProgressBar, Status)
    
    'Init alchemy arrays
    Call InitAlchemyArrays(ProgressBar, Status)
    
    '----------------------------------------------------------------------------------------------
    Status.Caption = "Loading Interface..."
    DoEvents
    Call WS.InitWorkshop(ProgressBar, Status)
    
    If ProgressBar.Width <> Val(ProgressBar.Tag) Then ProgressBar.Width = Val(ProgressBar.Tag)
    
    DoEvents
    
    SPLASH_LOADING = False
    'set the title bar to display update available if there is one available
    If UPDATE_AVAILABLE Then WS.Caption = WORKSHOP_TITLE & WORKSHOP_UPDATE_MESSAGE
    'set the interval to 0 we will check every 15 minutes for an update behind the scenes
    UPDATE_CHECK_COUNTER = 0
    'display the workshop
    WS.Show
    'unload the splash screen
    Unload Me
        
End Sub

