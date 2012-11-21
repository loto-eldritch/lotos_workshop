VERSION 5.00
Begin VB.Form Options 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Workshop Options"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   10650
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
      Height          =   4005
      Left            =   0
      TabIndex        =   11
      Tag             =   "General Settings"
      Top             =   -90
      Width           =   5640
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Auto-Update"
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
         Height          =   765
         Index           =   0
         Left            =   30
         TabIndex        =   29
         Top             =   180
         Width           =   2295
         Begin VB.CheckBox chkAutoUpdateOnStart 
            Caption         =   "Auto Update on Startup"
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   330
            Width           =   2115
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   325
         Left            =   4665
         TabIndex        =   28
         Top             =   3510
         Width           =   855
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   325
         Left            =   3720
         TabIndex        =   27
         Top             =   3510
         Width           =   855
      End
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Buff Values"
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
         Height          =   2520
         Index           =   1
         Left            =   2370
         TabIndex        =   12
         Top             =   180
         Width           =   3225
         Begin VB.CommandButton cmd_MaxBuffs 
            Caption         =   "Reset Buffs"
            Height          =   325
            Index           =   1
            Left            =   1695
            TabIndex        =   26
            Top             =   1950
            Width           =   1305
         End
         Begin VB.CommandButton cmd_MaxBuffs 
            Caption         =   "Max Buffs"
            Height          =   325
            Index           =   0
            Left            =   300
            TabIndex        =   24
            Top             =   1950
            Width           =   1305
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   7
            Left            =   2340
            MaxLength       =   3
            TabIndex        =   10
            Top             =   1200
            Width           =   690
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   6
            Left            =   2340
            MaxLength       =   3
            TabIndex        =   6
            Top             =   900
            Width           =   690
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   5
            Left            =   2340
            MaxLength       =   3
            TabIndex        =   9
            Top             =   600
            Width           =   690
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   4
            Left            =   2340
            MaxLength       =   3
            TabIndex        =   5
            Top             =   300
            Width           =   690
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   3
            Left            =   690
            MaxLength       =   3
            TabIndex        =   8
            Top             =   1200
            Width           =   690
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   2
            Left            =   690
            MaxLength       =   3
            TabIndex        =   4
            Top             =   900
            Width           =   690
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   1
            Left            =   690
            MaxLength       =   3
            TabIndex        =   7
            Top             =   600
            Width           =   690
         End
         Begin VB.TextBox txt_BuffValue 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   0
            Left            =   690
            MaxLength       =   3
            TabIndex        =   3
            Top             =   300
            Width           =   690
         End
         Begin VB.Label option_Label 
            Caption         =   "Max = Buff value with 25% ToA Bonus"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   25
            Top             =   1530
            Width           =   2895
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Cha"
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
            Height          =   195
            Index           =   7
            Left            =   1830
            TabIndex        =   20
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Pie"
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
            Height          =   195
            Index           =   6
            Left            =   1830
            TabIndex        =   19
            Top             =   900
            Width           =   375
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Emp"
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
            Height          =   195
            Index           =   5
            Left            =   1830
            TabIndex        =   18
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Int"
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
            Height          =   195
            Index           =   4
            Left            =   1830
            TabIndex        =   17
            Top             =   300
            Width           =   375
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Qui"
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
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   16
            Top             =   1200
            Width           =   375
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Dex"
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
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   15
            Top             =   900
            Width           =   375
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Con"
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
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   14
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lbl_AttributeBuff 
            Caption         =   "Str"
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
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   13
            Top             =   300
            Width           =   375
         End
      End
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Recent History (50 Max)"
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
         Height          =   1095
         Index           =   3
         Left            =   30
         TabIndex        =   22
         Top             =   2730
         Width           =   2295
         Begin VB.CommandButton cmd_ClearHistory 
            Caption         =   "Clear Recent History"
            Height          =   325
            Left            =   120
            TabIndex        =   0
            Top             =   630
            Width           =   1830
         End
         Begin VB.Image cmd_HistCountDec 
            Height          =   165
            Left            =   600
            Top             =   300
            Width           =   195
         End
         Begin VB.Image cmd_HistCountInc 
            Height          =   165
            Left            =   1275
            Top             =   300
            Width           =   195
         End
         Begin VB.Label lbl_History 
            Alignment       =   2  'Center
            Caption         =   "50"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   832
            TabIndex        =   23
            Top             =   300
            Width           =   420
         End
      End
      Begin VB.Frame frameOptionsSub 
         Caption         =   "Exit/Minimize Behavior"
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
         Height          =   1715
         Index           =   2
         Left            =   30
         TabIndex        =   21
         Top             =   990
         Width           =   2295
         Begin VB.CheckBox chkSystrayExit 
            Caption         =   "Exit to Systray (System Clock Bar)"
            ForeColor       =   &H00800000&
            Height          =   390
            Left            =   120
            TabIndex        =   1
            Top             =   330
            Width           =   1860
         End
         Begin VB.CheckBox chkSystrayMinimize 
            Caption         =   "Minimize to Systray (System Clock Bar)"
            ForeColor       =   &H00800000&
            Height          =   390
            Left            =   120
            TabIndex        =   2
            Top             =   930
            Width           =   1860
         End
      End
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PANEL_INDEX As Long

Private Sub Check1_Click()

End Sub

Private Sub Form_Load()
 
    Options.Width = 5720
    Options.Height = 4290
    
    Options.Caption = "Workshop Options - General Settings"
    
    PANEL_INDEX = 0
    
    Set cmd_HistCountDec.Picture = LoadResPicture("LARGE_LEFT_UP", vbResBitmap)
    Set cmd_HistCountInc.Picture = LoadResPicture("LARGE_RIGHT_UP", vbResBitmap)
   
    lbl_History.Caption = MAX_HISTORY_COUNT
    
    If AUTO_UPDATE Then chkAutoUpdateOnStart.Value = vbChecked
    If EXIT_TO_SYSTRAY Then chkSystrayExit.Value = vbChecked
    If MINIMIZE_TO_SYSTRAY Then chkSystrayMinimize.Value = vbChecked
    
    txt_BuffValue(0).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_STR")
    If (txt_BuffValue(0).Text) < 0 Then txt_BuffValue(0).Text = "0"
    
    txt_BuffValue(1).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_CON")
    If (txt_BuffValue(1).Text) < 0 Then txt_BuffValue(1).Text = "0"
    
    txt_BuffValue(2).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_DEX")
    If (txt_BuffValue(2).Text) < 0 Then txt_BuffValue(2).Text = "0"
    
    txt_BuffValue(3).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_QUI")
    If (txt_BuffValue(3).Text) < 0 Then txt_BuffValue(3).Text = "0"
    
    txt_BuffValue(4).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_INT")
    If (txt_BuffValue(4).Text) < 0 Then txt_BuffValue(4).Text = "0"
    
    txt_BuffValue(5).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_EMP")
    If (txt_BuffValue(5).Text) < 0 Then txt_BuffValue(5).Text = "0"
    
    txt_BuffValue(6).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_PIE")
    If (txt_BuffValue(6).Text) < 0 Then txt_BuffValue(6).Text = "0"
    
    txt_BuffValue(7).Text = GetRegValue(HKEY_LOCAL_MACHINE, REGKEY, "BUFF_CHA")
    If (txt_BuffValue(7).Text) < 0 Then txt_BuffValue(7).Text = "0"
    
    
End Sub

Private Sub cmd_MaxBuffs_Click(Index As Integer)

    If Index = 0 Then
        txt_BuffValue(0).Text = "155"
        txt_BuffValue(1).Text = "155"
        txt_BuffValue(2).Text = "155"
        txt_BuffValue(3).Text = "93"
        txt_BuffValue(4).Text = "81"
        txt_BuffValue(5).Text = "0"
        txt_BuffValue(6).Text = "81"
        txt_BuffValue(7).Text = "0"
    ElseIf Index = 1 Then
        txt_BuffValue(0).Text = "0"
        txt_BuffValue(1).Text = "0"
        txt_BuffValue(2).Text = "0"
        txt_BuffValue(3).Text = "0"
        txt_BuffValue(4).Text = "0"
        txt_BuffValue(5).Text = "0"
        txt_BuffValue(6).Text = "0"
        txt_BuffValue(7).Text = "0"
    End If

'Formula is delve *1.25(if nurt is 1pt higher than spell level) *1.25 (at 25% buff effectiveness) with a cap of 93 for specs, 62 for bases (on players).
End Sub

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()

    'do reg saving stuff here
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "MAX_HISTORY_COUNT", lbl_History.Caption)
    MAX_HISTORY_COUNT = Val(lbl_History.Caption)
    
    If chkAutoUpdateOnStart.Value = vbChecked Then
        AUTO_UPDATE = True
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "AUTO_UPDATE", "-1")
    Else
        AUTO_UPDATE = False
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "AUTO_UPDATE", "0")
    End If
        
    If chkSystrayMinimize.Value = vbChecked Then
        MINIMIZE_TO_SYSTRAY = True
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "SYSTRAY_MINI", "-1")
    Else
        MINIMIZE_TO_SYSTRAY = False
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "SYSTRAY_MINI", "0")
    End If
    
    If chkSystrayExit.Value = vbChecked Then
        EXIT_TO_SYSTRAY = True
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "SYSTRAY_EXIT", "-1")
    Else
        EXIT_TO_SYSTRAY = False
        Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "SYSTRAY_EXIT", "0")
    End If
    
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_STR", STR(Val(txt_BuffValue(0).Text)))
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_CON", STR(Val(txt_BuffValue(1).Text)))
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_DEX", STR(Val(txt_BuffValue(2).Text)))
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_QUI", STR(Val(txt_BuffValue(3).Text)))
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_INT", STR(Val(txt_BuffValue(4).Text)))
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_EMP", STR(Val(txt_BuffValue(5).Text)))
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_PIE", STR(Val(txt_BuffValue(6).Text)))
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "BUFF_CHA", STR(Val(txt_BuffValue(7).Text)))
    
    If TOON.REALM = REALM_HIBERNIA And TOON.CLASS = TCH_VAMPIIR Then
        TOON.STAT_MATRIX(SM_STR, SM_LOC_BUFFS) = 0
        TOON.STAT_MATRIX(SM_CON, SM_LOC_BUFFS) = 0
        TOON.STAT_MATRIX(SM_DEX, SM_LOC_BUFFS) = 0
        TOON.STAT_MATRIX(SM_QUI, SM_LOC_BUFFS) = 0
        TOON.STAT_MATRIX(SM_INT, SM_LOC_BUFFS) = 0
        TOON.STAT_MATRIX(SM_EMP, SM_LOC_BUFFS) = 0
        TOON.STAT_MATRIX(SM_PIE, SM_LOC_BUFFS) = 0
        TOON.STAT_MATRIX(SM_CHA, SM_LOC_BUFFS) = 0
    Else
        TOON.STAT_MATRIX(SM_STR, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_STR).Text)
        TOON.STAT_MATRIX(SM_CON, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_CON).Text)
        TOON.STAT_MATRIX(SM_DEX, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_DEX).Text)
        TOON.STAT_MATRIX(SM_QUI, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_QUI).Text)
        TOON.STAT_MATRIX(SM_INT, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_INT).Text)
        TOON.STAT_MATRIX(SM_EMP, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_EMP).Text)
        TOON.STAT_MATRIX(SM_PIE, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_PIE).Text)
        TOON.STAT_MATRIX(SM_CHA, SM_LOC_BUFFS) = Val(txt_BuffValue(WS_ATTR_CHA).Text)
    End If
    
    Call WS.RefreshAttributeLabels
    
    Unload Me
    
End Sub

Private Sub cmd_ClearHistory_Click()

    Call ClearRecentHistory
    
End Sub

Private Sub cmd_HistCountDec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_HistCountDec.Picture = LoadResPicture("LARGE_LEFT_DOWN", vbResBitmap)
    
End Sub

Private Sub cmd_HistCountDec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_HistCountDec.Picture = LoadResPicture("LARGE_LEFT_UP", vbResBitmap)
    
    If Val(lbl_History.Caption) <> 5 Then lbl_History.Caption = Val(lbl_History.Caption) - 1
    
End Sub

Private Sub cmd_HistCountInc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_HistCountInc.Picture = LoadResPicture("LARGE_RIGHT_DOWN", vbResBitmap)
    
End Sub

Private Sub cmd_HistCountInc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_HistCountInc.Picture = LoadResPicture("LARGE_RIGHT_UP", vbResBitmap)
    
    If Val(lbl_History.Caption) <> 50 Then lbl_History.Caption = Val(lbl_History.Caption) + 1
    
End Sub

Private Sub txt_BuffValue_Change(Index As Integer)

    txt_BuffValue(Index).Text = Val(txt_BuffValue(Index).Text)
    
End Sub

Private Sub txt_BuffValue_GotFocus(Index As Integer)

    txt_BuffValue(Index).SelLength = Len(txt_BuffValue(Index).Text)
    
End Sub

Private Sub txt_BuffValue_KeyPress(Index As Integer, KeyAscii As Integer)

    If (KeyAscii < 47 Or KeyAscii > 58) Then
        If (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyDelete) Then
            KeyAscii = 0
        End If
    Else
        If Len(txt_BuffValue(Index).Text) = 1 And Val(txt_BuffValue(Index).Text) = 0 Then txt_BuffValue(Index).Text = vbNullString
        If Len(txt_BuffValue(Index).Text) = 3 Then txt_BuffValue(Index).Text = vbNullString
    End If
    
    
    txt_BuffValue(Index).SelStart = Len(txt_BuffValue(Index).Text)
    
End Sub
