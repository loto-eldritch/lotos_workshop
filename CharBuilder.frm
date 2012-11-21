VERSION 5.00
Begin VB.Form CB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loto's Character Workshop - Character Builder"
   ClientHeight    =   6930
   ClientLeft      =   6885
   ClientTop       =   1080
   ClientWidth     =   11730
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CharBuilder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11730
   Visible         =   0   'False
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuShowSCAttrib 
         Caption         =   "Show Spellcrafted Attributes"
      End
   End
   Begin VB.Menu mnuSpellcraft 
      Caption         =   "&Spellcraft"
   End
End
Attribute VB_Name = "CB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Refresh_Attribute_Labels()
    
    Dim Ctr As Integer
    
    For Ctr = 0 To 7
        lbl_StatValue(Ctr).Caption = TOON.STAT_MATRIX((Ctr + 1), 1) + TOON.STAT_MATRIX((Ctr + 1), 2) + TOON.STAT_MATRIX((Ctr + 1), 3)
    Next Ctr
        
    lbl_ResistValue(0).Tag = TOON.STAT_MATRIX(12, 1)
    lbl_ResistValue(0) = Val(lbl_ResistValue(0).Tag) + Sum_Column(12, 1)
    
    lbl_ResistValue(1).Tag = TOON.STAT_MATRIX(13, 1)
    lbl_ResistValue(1) = Val(lbl_ResistValue(1).Tag) + Sum_Column(13, 1)
    
    lbl_ResistValue(2).Tag = TOON.STAT_MATRIX(14, 1)
    lbl_ResistValue(2) = Val(lbl_ResistValue(2).Tag) + Sum_Column(14, 1)
    
    lbl_ResistValue(3).Tag = TOON.STAT_MATRIX(18, 1)
    lbl_ResistValue(3) = Val(lbl_ResistValue(3).Tag) + Sum_Column(18, 1)
    
    lbl_ResistValue(4).Tag = TOON.STAT_MATRIX(16, 1)
    lbl_ResistValue(4) = Val(lbl_ResistValue(4).Tag) + Sum_Column(16, 1)
    
    lbl_ResistValue(5).Tag = TOON.STAT_MATRIX(19, 1)
    lbl_ResistValue(5) = Val(lbl_ResistValue(5).Tag) + Sum_Column(19, 1)
    
    lbl_ResistValue(6).Tag = TOON.STAT_MATRIX(15, 1)
    lbl_ResistValue(6) = Val(lbl_ResistValue(6).Tag) + Sum_Column(15, 1)
    
    lbl_ResistValue(7).Tag = TOON.STAT_MATRIX(20, 1)
    lbl_ResistValue(7) = Val(lbl_ResistValue(7).Tag) + Sum_Column(20, 1)
    
    lbl_ResistValue(8).Tag = TOON.STAT_MATRIX(17, 1)
    lbl_ResistValue(8) = Val(lbl_ResistValue(8).Tag) + Sum_Column(17, 1)
    
    For Ctr = 0 To 8
        If Val(lbl_ResistValue(Ctr).Tag) <> 0 Then
            lbl_ResistValue(Ctr).FontBold = True
            lbl_ResistName(Ctr).FontBold = True
            lbl_ResistValue(Ctr).ToolTipText = "Racial Resist +" & lbl_ResistValue(Ctr).Tag
            lbl_ResistName(Ctr).ToolTipText = "Racial Resist +" & lbl_ResistValue(Ctr).Tag
        Else
            lbl_ResistValue(Ctr).FontBold = False
            lbl_ResistName(Ctr).FontBold = False
            lbl_ResistValue(Ctr).ToolTipText = vbNullString
            lbl_ResistName(Ctr).ToolTipText = vbNullString
        End If
    Next Ctr

End Sub

Private Sub Reset_Attribute_Labels()
    
    Dim Ctr As Integer
    
    lbl_StatName(0).Caption = "Str"
    lbl_StatName(1).Caption = "Con"
    lbl_StatName(2).Caption = "Dex"
    lbl_StatName(3).Caption = "Qui"
    lbl_StatName(4).Caption = "Int"
    lbl_StatName(5).Caption = "Emp"
    lbl_StatName(6).Caption = "Pie"
    lbl_StatName(7).Caption = "Cha"
    
    For Ctr = 0 To 7
        lbl_StatName(Ctr).FontBold = False
        lbl_StatValue(Ctr).FontBold = False
        lbl_StatValue(Ctr).Tag = 0
        lbl_StatValue(Ctr).ForeColor = &H8000000D
        lbl_StatValue(Ctr).Caption = vbNullString
        lbl_PST(Ctr).Caption = vbNullString
    Next Ctr
    
    For Ctr = 0 To 8
        lbl_ResistName(Ctr).FontBold = False
        lbl_ResistName(Ctr).Tag = 0
        lbl_ResistValue(Ctr).FontBold = False
        lbl_ResistValue(Ctr) = 0
    Next Ctr
    
    txt_Level.Text = "5"
    
End Sub

Private Sub cmb_Class_Click()

    Call Reset_Attribute_Labels         'zero them out
    
    Call Assign_Race_Attributes(TOON.REALM, TOON.RACE, TOON)
    Call Assign_pstStats(TOON.REALM, cmb_Class.Text, TOON)
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
    
    Call Refresh_Attribute_Labels       'refresh them based on toon info
    
    Select Case TOON.pStat
        Case 1  'str
            lbl_StatName(0).Caption = lbl_StatName(0).Caption
            lbl_PST(0).Caption = "P"
            lbl_StatName(0).FontBold = True
            lbl_StatValue(0).FontBold = True
            
        Case 2  'con
            lbl_StatName(1).Caption = lbl_StatName(1).Caption
            lbl_PST(1).Caption = "P"
            lbl_StatName(1).FontBold = True
            lbl_StatValue(1).FontBold = True
            
        Case 3  'dex
            lbl_StatName(2).Caption = lbl_StatName(2).Caption
            lbl_PST(2).Caption = "P"
            lbl_StatName(2).FontBold = True
            lbl_StatValue(2).FontBold = True
            
        Case 4  'qui
            lbl_StatName(3).Caption = lbl_StatName(3).Caption
            lbl_PST(3).Caption = "P"
            lbl_StatName(3).FontBold = True
            lbl_StatValue(3).FontBold = True
            
        Case 5  'int
            lbl_StatName(4).Caption = lbl_StatName(4).Caption
            lbl_PST(4).Caption = "P"
            lbl_StatName(4).FontBold = True
            lbl_StatValue(4).FontBold = True
            
        Case 6  'emp
            lbl_StatName(5).Caption = lbl_StatName(5).Caption
            lbl_PST(5).Caption = "P"
            lbl_StatName(5).FontBold = True
            lbl_StatValue(5).FontBold = True
            
        Case 7  'pie
            lbl_StatName(6).Caption = lbl_StatName(6).Caption
            lbl_PST(6).Caption = "P"
            lbl_StatName(6).FontBold = True
            lbl_StatValue(6).FontBold = True
            
        Case 8  'cha
            lbl_StatName(7).Caption = lbl_StatName(7).Caption
            lbl_PST(7).Caption = "P"
            lbl_StatName(7).FontBold = True
            lbl_StatValue(7).FontBold = True
    End Select
    
    Select Case TOON.sStat
        Case 1  'str
            lbl_StatName(0).Caption = lbl_StatName(0).Caption
            lbl_PST(0).Caption = "S"
            lbl_StatName(0).FontBold = True
            lbl_StatValue(0).FontBold = True
            
        Case 2  'con
            lbl_StatName(1).Caption = lbl_StatName(1).Caption
            lbl_PST(1).Caption = "S"
            lbl_StatName(1).FontBold = True
            lbl_StatValue(1).FontBold = True
            
        Case 3  'dex
            lbl_StatName(2).Caption = lbl_StatName(2).Caption
            lbl_PST(2).Caption = "S"
            lbl_StatName(2).FontBold = True
            lbl_StatValue(2).FontBold = True
            
        Case 4  'qui
            lbl_StatName(3).Caption = lbl_StatName(3).Caption
            lbl_PST(3).Caption = "S"
            lbl_StatName(3).FontBold = True
            lbl_StatValue(3).FontBold = True
            
        Case 5  'int
            lbl_StatName(4).Caption = lbl_StatName(4).Caption
            lbl_PST(4).Caption = "S"
            lbl_StatName(4).FontBold = True
            lbl_StatValue(4).FontBold = True
            
        Case 6  'emp
            lbl_StatName(5).Caption = lbl_StatName(5).Caption
            lbl_PST(5).Caption = "S"
            lbl_StatName(5).FontBold = True
            lbl_StatValue(5).FontBold = True
            
        Case 7  'pie
            lbl_StatName(6).Caption = lbl_StatName(6).Caption
            lbl_PST(6).Caption = "S"
            lbl_StatName(6).FontBold = True
            lbl_StatValue(6).FontBold = True
            
        Case 8  'cha
            lbl_StatName(7).Caption = lbl_StatName(7).Caption
            lbl_PST(7).Caption = "S"
            lbl_StatName(7).FontBold = True
            lbl_StatValue(7).FontBold = True
    End Select
    
    Select Case TOON.tStat
        Case 1  'str
            lbl_StatName(0).Caption = lbl_StatName(0).Caption
            lbl_PST(0).Caption = "T"
            lbl_StatName(0).FontBold = True
            lbl_StatValue(0).FontBold = True
            
        Case 2  'con
            lbl_StatName(1).Caption = lbl_StatName(1).Caption
            lbl_PST(1).Caption = "T"
            lbl_StatName(1).FontBold = True
            lbl_StatValue(1).FontBold = True
            
        Case 3  'dex
            lbl_StatName(2).Caption = lbl_StatName(2).Caption
            lbl_PST(2).Caption = "T"
            lbl_StatName(2).FontBold = True
            lbl_StatValue(2).FontBold = True
            
        Case 4  'qui
            lbl_StatName(3).Caption = lbl_StatName(3).Caption
            lbl_PST(3).Caption = "T"
            lbl_StatName(3).FontBold = True
            lbl_StatValue(3).FontBold = True
            
        Case 5  'int
            lbl_StatName(4).Caption = lbl_StatName(4).Caption
            lbl_PST(4).Caption = "T"
            lbl_StatName(4).FontBold = True
            lbl_StatValue(4).FontBold = True
            
        Case 6  'emp
            lbl_StatName(5).Caption = lbl_StatName(5).Caption
            lbl_PST(5).Caption = "T"
            lbl_StatName(5).FontBold = True
            lbl_StatValue(5).FontBold = True
            
        Case 7  'pie
            lbl_StatName(6).Caption = lbl_StatName(6).Caption
            lbl_PST(6).Caption = "T"
            lbl_StatName(6).FontBold = True
            lbl_StatValue(6).FontBold = True
            
        Case 8  'cha
            lbl_StatName(7).Caption = lbl_StatName(7).Caption
            lbl_PST(7).Caption = "T"
            lbl_StatName(7).FontBold = True
            lbl_StatValue(7).FontBold = True
    End Select
    
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
    
    frame_Attributes.Caption = "Character Attributes :: (Points: " & TOON.CREATION_POINTS & ")"
End Sub

Private Function Calc_SpecPoints(Mult As Single, LEVEL As Single) As Long

    Dim Ctr As Long
    Dim Sum As Long
    Dim Level_LO As Single
    Dim Level_HI As Long
    
    Level_HI = Trunc(LEVEL)
    
    If (LEVEL <= 0) Or (LEVEL > 50) Then
        Sum = 0
    ElseIf LEVEL < 6 Then
    
        For Ctr = 0 To LEVEL
            Sum = Sum + Ctr
        Next Ctr
        
        Sum = Sum - 1
    Else
        For Ctr = 6 To Level_HI
        
            If Ctr > 40 Then
                Sum = Sum + Trunc(Mult * Trunc(Ctr - 1) / 2)
            End If
            
            Sum = Sum + Trunc(Mult * Ctr)
            
        Next Ctr
        
        Level_LO = LEVEL - Level_HI
        
        If Level_LO > 0 Then
        
            Sum = Sum + Trunc(Mult * Level_HI / 2)
            
        End If
        
        Sum = Sum + 14
    End If
    
    Calc_SpecPoints = Sum
        
End Function

Private Sub cmd_addSkillLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_addSkillLevel(Index).Picture = LoadResPicture("DOUBLE_RIGHT_DOWN", vbResBitmap)
    
End Sub

Private Sub cmd_addSkillLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim TotalSpecPoints As Long
    Dim SpentSpecPoints As Long
    
    Set cmd_addSkillLevel(Index).Picture = LoadResPicture("DOUBLE_RIGHT_UP", vbResBitmap)
    
    If Val(lbl_SkillLevel(Index)) + 1 <= TOON.LEVEL Then
        TotalSpecPoints = Calc_SpecPoints(TOON.MULTIPLIER, TOON.LEVEL)
        
        lbl_SkillLevel(Index).Caption = Val(lbl_SkillLevel(Index).Caption) + 1
        SpentSpecPoints = ((Val(lbl_SkillLevel(Index).Caption) * (Val(lbl_SkillLevel(Index).Caption) + 1)) / 2) - 1
        lbl_SpecSpent(Index).Caption = "Spent: " & SpentSpecPoints
        
        frm_Skill.Caption = "Skills Trainer :: Spec Points: " & TotalSpecPoints & " :: Remaining: " & TotalSpecPoints - SpentSpecPoints
    End If
    
End Sub

Private Sub cmd_AddStat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_AddStat(Index).Picture = LoadResPicture("SRIGHT_DOWN", vbResBitmap)
    
End Sub

Public Sub ClickAddStat(Index As Integer, lCount As Long)

    Dim lCnt As Long
    
    If lCount <> 0 Then
                       
        For lCnt = 1 To lCount
            Call cmd_AddStat_MouseUp(Index, 0, 0, 0, 0)
        Next lCnt
        
    End If
    
End Sub

Private Sub cmd_AddStat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lCode As Long
    
    Set cmd_AddStat(Index).Picture = LoadResPicture("SRIGHT_UP", vbResBitmap)
    
    'add to selected stat
    If (Val(lbl_StatValue(Index).Tag) < 10) And TOON.CREATION_POINTS - 1 >= 0 Then
        lbl_StatValue(Index).Tag = Val(lbl_StatValue(Index).Tag) + 1
        TOON.CREATION_POINTS = TOON.CREATION_POINTS - 1
        lbl_StatValue(Index).Caption = Val(lbl_StatValue(Index).Caption) + 1
        lbl_StatValue(Index).ForeColor = &H80&
    ElseIf (Val(lbl_StatValue(Index).Tag) >= 10) And (Val(lbl_StatValue(Index).Tag) < 15) And TOON.CREATION_POINTS - 2 >= 0 Then
        lbl_StatValue(Index).Tag = Val(lbl_StatValue(Index).Tag) + 1
        TOON.CREATION_POINTS = TOON.CREATION_POINTS - 2
        lbl_StatValue(Index).Caption = Val(lbl_StatValue(Index).Caption) + 1
        lbl_StatValue(Index).ForeColor = &HC0&
    ElseIf (Val(lbl_StatValue(Index).Tag) >= 15) And (Val(lbl_StatValue(Index).Tag) < 18) And TOON.CREATION_POINTS - 3 >= 0 Then
        lbl_StatValue(Index).Tag = Val(lbl_StatValue(Index).Tag) + 1
        TOON.CREATION_POINTS = TOON.CREATION_POINTS - 3
        lbl_StatValue(Index).Caption = Val(lbl_StatValue(Index).Caption) + 1
        lbl_StatValue(Index).ForeColor = &HFF&
    End If
    
    frame_Attributes.Caption = "Character Attributes :: (Points: " & TOON.CREATION_POINTS & ")"
    
    Select Case Index
        Case 0  'str
            lCode = 1
        Case 1  'con
            lCode = 2
        Case 2  'dex
            lCode = 3
        Case 3  'qui
            lCode = 4
        Case 4  'int
            lCode = 5
        Case 5  'emp
            lCode = 6
        Case 6  'pie
            lCode = 7
        Case 7  'cha
            lCode = 8
    End Select
    
    TOON.STAT_MATRIX(lCode, 2) = Val(lbl_StatValue(Index).Tag)
    
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
    
End Sub


Private Sub cmd_subSkillLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Set cmd_subSkillLevel(Index).Picture = LoadResPicture("DOUBLE_LEFT_DOWN", vbResBitmap)
    
End Sub

Private Sub cmd_subSkillLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_subSkillLevel(Index).Picture = LoadResPicture("DOUBLE_LEFT_UP", vbResBitmap)
    
    Dim TotalSpecPoints As Long
    Dim SpentSpecPoints As Long
    
    If (Val(lbl_SkillLevel(Index)) - 1) > 0 Then
        TotalSpecPoints = Calc_SpecPoints(TOON.MULTIPLIER, TOON.LEVEL)
        
        lbl_SkillLevel(Index).Caption = Val(lbl_SkillLevel(Index).Caption) - 1
        SpentSpecPoints = ((Val(lbl_SkillLevel(Index).Caption) * (Val(lbl_SkillLevel(Index).Caption) + 1)) / 2) - 1
        lbl_SpecSpent(Index).Caption = "Spent: " & SpentSpecPoints
        
        frm_Skill.Caption = "Skills Trainer :: Spec Points: " & TotalSpecPoints & " :: Remaining: " & TotalSpecPoints - SpentSpecPoints
    End If
    
End Sub

Private Sub cmd_SubStat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_SubStat(Index).Picture = LoadResPicture("SLEFT_DOWN", vbResBitmap)

End Sub

Private Sub cmd_SubStat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lCode As Long
    
    Set cmd_SubStat(Index).Picture = LoadResPicture("SLEFT_UP", vbResBitmap)
    
    'subtract from selected stat if there were points added
    If (Val(lbl_StatValue(Index).Tag) > 0) And (Val(lbl_StatValue(Index).Tag) <= 10) And (TOON.CREATION_POINTS + 1 <= 30) Then
        lbl_StatValue(Index).Tag = Val(lbl_StatValue(Index).Tag) - 1
        TOON.CREATION_POINTS = TOON.CREATION_POINTS + 1
        lbl_StatValue(Index).Caption = Val(lbl_StatValue(Index).Caption) - 1
        lbl_StatValue(Index).ForeColor = &H80&
    ElseIf (Val(lbl_StatValue(Index).Tag) > 10) And (Val(lbl_StatValue(Index).Tag) <= 15) And (TOON.CREATION_POINTS + 2 <= 30) Then
        lbl_StatValue(Index).Tag = Val(lbl_StatValue(Index).Tag) - 1
        TOON.CREATION_POINTS = TOON.CREATION_POINTS + 2
        lbl_StatValue(Index).Caption = Val(lbl_StatValue(Index).Caption) - 1
        lbl_StatValue(Index).ForeColor = &HC0&
    ElseIf (Val(lbl_StatValue(Index).Tag) > 15) And (Val(lbl_StatValue(Index).Tag) <= 18) And (TOON.CREATION_POINTS + 3 <= 30) Then
        lbl_StatValue(Index).Tag = Val(lbl_StatValue(Index).Tag) - 1
        TOON.CREATION_POINTS = TOON.CREATION_POINTS + 3
        lbl_StatValue(Index).Caption = Val(lbl_StatValue(Index).Caption) - 1
        lbl_StatValue(Index).ForeColor = &HFF&
    End If
    
    If Val(lbl_StatValue(Index).Tag) = 0 Then lbl_StatValue(Index).ForeColor = &H8000000D
    
    frame_Attributes.Caption = "Character Attributes :: (Points: " & TOON.CREATION_POINTS & ")"
    
    Select Case Index
        Case 0  'str
            lCode = 1
        Case 1  'con
            lCode = 2
        Case 2  'dex
            lCode = 3
        Case 3  'qui
            lCode = 4
        Case 4  'int
            lCode = 5
        Case 5  'emp
            lCode = 6
        Case 6  'pie
            lCode = 7
        Case 7  'cha
            lCode = 8
    End Select
    
    TOON.STAT_MATRIX(lCode, 2) = Val(lbl_StatValue(Index).Tag)
    
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
End Sub

Private Sub cmb_Realm_Click()

    Call Reset_Attribute_Labels
    
    TOON.REALM = cmb_Realm.ListIndex
    If Not TOON.R_RANK = 0 Then lbl_RealmTitle.Caption = GetRealmTitle(TOON.R_RANK, TOON.REALM, TOON.GENDER)
    
    cmb_Class.Clear
    Call Populate_Races(cmb_Race, TOON.REALM)
    
    Call Set_Doll_Image(TOON.REALM)
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
    
    Call SetNewValue(HKEY_LOCAL_MACHINE, REGKEY, NUM_VAL, "REALM", STR(TOON.REALM))
    
End Sub

Private Sub cmb_Race_Click()
       
    Call Reset_Attribute_Labels
    
    TOON.RACE = cmb_Race.ListIndex
    Call Populate_Classes(cmb_Class, TOON.REALM, TOON.RACE)
    Call Assign_Race_Attributes(TOON.REALM, TOON.RACE, TOON)
        
    Call Refresh_Attribute_Labels
    
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
    
End Sub

Private Sub Form_Load()
    
    Dim Ctr As Long
    
    mnuShowSCAttrib.Enabled = False
    
    TOON.LEVEL = 5
    txt_Level.Text = TOON.LEVEL
        
    'level up and down arrow graphics
    Set cmd_DecreaseLevel.Picture = LoadResPicture("LARGE_LEFT_UP", vbResBitmap)
    Set cmd_IncreaseLevel.Picture = LoadResPicture("LARGE_RIGHT_UP", vbResBitmap)
    
    Set cmd_subSkillLevel(0).Picture = LoadResPicture("DOUBLE_LEFT_UP", vbResBitmap)
    Set cmd_addSkillLevel(0).Picture = LoadResPicture("DOUBLE_RIGHT_UP", vbResBitmap)
    
    For Ctr = 0 To 7
        Set cmd_SubStat(Ctr).Picture = LoadResPicture("SLEFT_UP", vbResBitmap)
        Set cmd_AddStat(Ctr).Picture = LoadResPicture("SRIGHT_UP", vbResBitmap)
    Next Ctr
    
    opt_Male.Value = True
    
    cmb_Realm.ListIndex = TOON.REALM
    cmb_Race.ListIndex = TOON.RACE
    
    If TOON.GENDER = 0 Then opt_Male.Value = True
    If TOON.GENDER = 1 Then opt_Female.Value = True
    
    cmb_Class.ListIndex = GetClassComboID(TOON.REALM, TOON.CLASS, cmb_Class)
    
End Sub

Private Sub lbl_StatValue_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    lbl_StatValue(Index).ToolTipText = vbNullString
    If TOON.STAT_MATRIX(Index + 1, 1) <> 0 Then
        lbl_StatValue(Index).ToolTipText = TOON.STAT_MATRIX((Index + 1), 1) & " Base"
        
        If TOON.STAT_MATRIX(Index + 1, 2) <> 0 Then
            lbl_StatValue(Index).ToolTipText = lbl_StatValue(Index).ToolTipText & _
                                        " + " & TOON.STAT_MATRIX((Index + 1), 2) & " Creation"
        End If
        
        If TOON.STAT_MATRIX(Index + 1, 3) <> 0 Then
            lbl_StatValue(Index).ToolTipText = lbl_StatValue(Index).ToolTipText & _
                                        " + " & TOON.STAT_MATRIX((Index + 1), 3) & " Level"
        End If
    End If
          
End Sub

Private Sub mnuExitCB_Click()

    Me.Hide
    
End Sub

Private Sub opt_Female_Click()
    
    If Not TOON.GENDER = 1 Then
        TOON.GENDER = 1
        Call txt_RealmPoints_Change
        Call Populate_Classes(cmb_Class, TOON.REALM, TOON.RACE)
    End If
    
End Sub

Private Sub opt_Male_Click()
    
    If Not TOON.GENDER = 0 Then
        TOON.GENDER = 0
        Call txt_RealmPoints_Change
        Call Populate_Classes(cmb_Class, TOON.REALM, TOON.RACE)
    End If
    
End Sub

Private Sub txt_Level_Change()

    Static bHereAlready As Boolean
    Static sBuffer As String
    
    If Not bHereAlready = True Then
        bHereAlready = True
        
        If Val(txt_Level.Text) < 51 And Val(txt_Level.Text) > 0 Then TOON.LEVEL = Val(txt_Level.Text)
        
        If cmb_Class.Text <> vbNullString Then
            frm_Skill.Caption = "Skills Trainer :: Spec Points: " & Calc_SpecPoints(TOON.MULTIPLIER, TOON.LEVEL)
            Call Calculate_PSTValues(TOON.LEVEL)
            Call Refresh_Attribute_Labels
        End If
    End If
    
    bHereAlready = False
    
End Sub

Private Sub txt_RealmPoints_Change()
    Dim lResult As Long
    
    Dim iLTemp As Long
    Dim iRTemp As Long
    
    txt_RealmPoints.Text = Format(txt_RealmPoints.Text, "###,###,###")
    txt_RealmPoints.SelStart = Len(txt_RealmPoints.Text)
    
    TOON.REALM_POINTS = Val(Format(txt_RealmPoints.Text, "#########"))  'assign realmpoints to toon
    
    lResult = GetRealmRank(TOON.REALM_POINTS, TOON.R_RANK, TOON.R_LEVEL)
    
    If lResult = 0 Then 'if it's all good then
        TOON.REALM_TITLE = GetRealmTitle(TOON.R_RANK, TOON.REALM, TOON.GENDER)    'give toon its title
        lbl_RealmTitle.Caption = TOON.REALM_TITLE                   'show the label its title
        
        TOON.REALM_ABILITY_POINTS = RealmSkillPoints(TOON.R_RANK, TOON.R_LEVEL)
        lbl_RealmAbilityPoints.Caption = TOON.REALM_ABILITY_POINTS
    End If
    
    iRTemp = TOON.R_RANK
    iLTemp = TOON.R_LEVEL
    
    If iLTemp < 9 Then
        iLTemp = iLTemp + 1
    Else
        iRTemp = iRTemp + 1
    End If
       
    lbl_Next_RL.Caption = Format(GetRealmPoints(iRTemp, iLTemp) - TOON.REALM_POINTS, "###,###,###")
    
    iRTemp = TOON.R_RANK
    iLTemp = TOON.R_LEVEL
    
    If iRTemp < 10 Then
        iRTemp = iRTemp + 1
    ElseIf iRTemp = 11 Then
        iRTemp = 12
        iLTemp = 0
    End If
    
    lbl_Next_RR.Caption = Format(GetRealmPoints(iRTemp, iLTemp) - TOON.REALM_POINTS, "###,###,###")
        
End Sub

Private Sub txt_RealmLevel_Change()

    If Val(txt_RealmRank.Text) = 0 Then txt_RealmRank.Text = "1"
    
    If Val(txt_RealmRank.Text) = 13 Then txt_RealmLevel.Text = "0"
    
    If (Val(txt_RealmLevel.Text) >= 0 And Val(txt_RealmLevel.Text) <= 9) Then
    
        TOON.R_RANK = Val(txt_RealmRank.Text)
        TOON.R_LEVEL = Val(txt_RealmLevel.Text)
        
        TOON.REALM_ABILITY_POINTS = RealmSkillPoints(TOON.R_RANK, TOON.R_LEVEL)
        
        lbl_RealmAbilityPoints.Caption = TOON.REALM_ABILITY_POINTS
        
        TOON.REALM_POINTS = GetRealmPoints(TOON.R_RANK, TOON.R_LEVEL)
        
        txt_RealmPoints.Text = TOON.REALM_POINTS
        
    End If


End Sub

Private Sub txt_RealmRank_Change()

    If Len(txt_RealmRank.Text) > 2 Then txt_RealmRank.Text = vbNullString
    If (Val(txt_RealmRank.Text) >= 1 And Val(txt_RealmRank.Text) <= 13) Then
        
        TOON.R_RANK = Val(txt_RealmRank.Text)
        
        If TOON.R_RANK = 13 Then txt_RealmLevel.Text = 0
        
        TOON.R_LEVEL = Val(txt_RealmLevel.Text)
        
        TOON.REALM_ABILITY_POINTS = RealmSkillPoints(TOON.R_RANK, TOON.R_LEVEL) 'give toon its RSPs
        
        lbl_RealmAbilityPoints.Caption = TOON.REALM_ABILITY_POINTS
        
        TOON.REALM_POINTS = GetRealmPoints(TOON.R_RANK, TOON.R_LEVEL)
        
        txt_RealmPoints.Text = TOON.REALM_POINTS
    Else
        txt_RealmRank.Text = vbNullString
    End If

End Sub

Private Sub cmd_DecreaseLevel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_DecreaseLevel.Picture = LoadResPicture("LARGE_LEFT_DOWN", vbResBitmap)
    
End Sub

Private Sub cmd_DecreaseLevel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
       
    Dim LVL As Long
    If (TOON.LEVEL > 5) Then
        If (TOON.LEVEL <= 40) Then
            LVL = TOON.LEVEL
            TOON.LEVEL = LVL
            TOON.LEVEL = TOON.LEVEL - 1
        Else
            TOON.LEVEL = TOON.LEVEL - 0.5
        End If
        
        If TOON.LEVEL - 1 = 4 Then
            TOON.STAT_MATRIX(TOON.pStat, 3) = 0
            TOON.STAT_MATRIX(TOON.sStat, 3) = 0
            TOON.STAT_MATRIX(TOON.tStat, 3) = 0
        End If
        
        If cmb_Class.Text <> vbNullString Then
            frm_Skill.Caption = "Skills Trainer :: Spec Points: " & Calc_SpecPoints(TOON.MULTIPLIER, TOON.LEVEL)
            Call Calculate_PSTValues(TOON.LEVEL)
            Call Refresh_Attribute_Labels
        End If
    End If
    
    txt_Level.Text = TOON.LEVEL
    
    Set cmd_DecreaseLevel.Picture = LoadResPicture("LARGE_LEFT_UP", vbResBitmap)
    
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
    
End Sub

Private Sub cmd_IncreaseLevel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Set cmd_IncreaseLevel.Picture = LoadResPicture("LARGE_RIGHT_DOWN", vbResBitmap)
    
End Sub

Private Sub cmd_IncreaseLevel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    Dim LVL As Long
    
    If (TOON.LEVEL >= 5 And TOON.LEVEL < 50) Then
        If (TOON.LEVEL < 40) Then
            LVL = TOON.LEVEL
            TOON.LEVEL = LVL
            TOON.LEVEL = TOON.LEVEL + 1
        Else
            TOON.LEVEL = TOON.LEVEL + 0.5
        End If
        
        If cmb_Class.Text <> vbNullString Then
            frm_Skill.Caption = "Skills Trainer :: Spec Points: " & Calc_SpecPoints(TOON.MULTIPLIER, TOON.LEVEL)
            Call Calculate_PSTValues(TOON.LEVEL)
            Call Refresh_Attribute_Labels
        End If
    End If

    txt_Level.Text = TOON.LEVEL
        
    Set cmd_IncreaseLevel.Picture = LoadResPicture("LARGE_RIGHT_UP", vbResBitmap)
    
    Call SC.Refresh_SC(SC_SETTINGS.BONUS_OPTION)
    
End Sub

Private Sub txt_RealmLevel_KeyPress(KeyAscii As Integer)

    If KeyAscii > 47 And KeyAscii < 58 Then
        txt_RealmLevel.Text = vbNullString
    End If
    
End Sub
