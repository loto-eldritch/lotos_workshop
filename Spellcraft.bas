Attribute VB_Name = "Spellcraft"
Option Explicit

'Private FileExtensions(19) As String

Private OCStartPercentages(5) As Long
Private ImbueMultipliers(6) As Double
Private ImbuePoints(50, 4) As Long
Private GemQualityOCMODS(6) As Long
Private ItemQualityOCMODS(6) As Long

Private Type SPELLCRAFT_SETTINGS
    REALM As Long
    CRAFT_SKILL As Long
    BONUS_OPTION As Integer
End Type
'#  0 = total bonus: normal spellcraft mode
'#  1 = distance to cap: normal spellcraft mode

Public Const BO_SCBONUS As Long = 0
Public Const BO_TOCAP As Long = 1

Public tGemType(6) As String    'gem types

Public nStat(7) As String       'stat names
Public nResist(8) As String     'resist names
Public nSkill(2, 45) As String  'skill names
Public nFocus(2, 12) As String  'focus names

Public GemInfo(3, 9) As String   '[0] = container, [1] = cut, [2] = cost, [3] = remake

Private gResist(9) As Long      'resist amounts
Private gStat(9) As Long        'stat amounts
Private gHIT(9) As Long         'hit amounts
Private gPower(9) As Long       'power amounts
Private gFocus(9) As Long       'focus amounts
Private gSkill(9) As Long       'skill amounts

Public GemMaterial_Dust(11) As String
Public GemMaterial_Liquid(10) As String

Public ServerCodes(1, 50) As String

Public Type Gem
    DUST_QUAN As Long
    DUST_NAME As String
    DUST_Index As Long
    
    LIQUID1_QUAN As Long
    LIQUID2_QUAN As Long
    LIQUID3_QUAN As Long
    LIQUID1_NAME As String
    LIQUID2_NAME As String
    LIQUID3_NAME As String
    LIQUID1_Index As Long
    LIQUID2_Index As Long
    LIQUID3_Index As Long
    
    CONTAINER_QUAN As Long
    CONTAINER_NAME As String
    CONTAINER_Index As Long
    
    GEM_PRICE As Long
End Type

Public SC_SETTINGS As SPELLCRAFT_SETTINGS

Public Const SC_NSKILL_ALL_ARCHERY As Long = 0
Public Const SC_NSKILL_ALL_DUALWIELD As Long = 1
Public Const SC_NSKILL_ALL_MAGIC As Long = 2
Public Const SC_NSKILL_ALL_MELEE As Long = 3

Public Const SC_NSKILL_ALB_BODYMAGIC As Long = 4
Public Const SC_NSKILL_ALB_CHANTS As Long = 5
Public Const SC_NSKILL_ALB_COLDMAGIC As Long = 6
Public Const SC_NSKILL_ALB_CRITICAL As Long = 7
Public Const SC_NSKILL_ALB_CROSSBOW As Long = 8
Public Const SC_NSKILL_ALB_CRUSH As Long = 9
Public Const SC_NSKILL_ALB_DEATHSERVANT As Long = 10
Public Const SC_NSKILL_ALB_DEATHSIGHT As Long = 11
Public Const SC_NSKILL_ALB_DUALWIELD As Long = 12
Public Const SC_NSKILL_ALB_EARTHMAGIC As Long = 13
Public Const SC_NSKILL_ALB_ENHANCEMENT As Long = 14
Public Const SC_NSKILL_ALB_ENVENOM As Long = 15
Public Const SC_NSKILL_ALB_FIREMAGIC As Long = 16
Public Const SC_NSKILL_ALB_FLEXIBLE As Long = 17
Public Const SC_NSKILL_ALB_INSTRUMENTS As Long = 18
Public Const SC_NSKILL_ALB_LONGBOW As Long = 19
Public Const SC_NSKILL_ALB_MATTERMAGIC As Long = 20
Public Const SC_NSKILL_ALB_MINDMAGIC As Long = 21
Public Const SC_NSKILL_ALB_PAINWORKING As Long = 22
Public Const SC_NSKILL_ALB_PARRY As Long = 23
Public Const SC_NSKILL_ALB_POLEARM As Long = 24
Public Const SC_NSKILL_ALB_REJUVENATION As Long = 25
Public Const SC_NSKILL_ALB_SHIELD As Long = 26
Public Const SC_NSKILL_ALB_SLASH As Long = 27
Public Const SC_NSKILL_ALB_SMITE As Long = 28
Public Const SC_NSKILL_ALB_SOULRENDING As Long = 29
Public Const SC_NSKILL_ALB_SPIRITMAGIC As Long = 30
Public Const SC_NSKILL_ALB_STAFF As Long = 31
Public Const SC_NSKILL_ALB_STEALTH As Long = 32
Public Const SC_NSKILL_ALB_THRUST As Long = 33
Public Const SC_NSKILL_ALB_TWOHANDED As Long = 34
Public Const SC_NSKILL_ALB_WINDMAGIC As Long = 35
Public Const SC_NSKILL_ALB_AURAMANIP As Long = 36
Public Const SC_NSKILL_ALB_FISTWRAP As Long = 37
Public Const SC_NSKILL_ALB_MAGNETISM As Long = 38
Public Const SC_NSKILL_ALB_MAULERSTAFF As Long = 39
Public Const SC_NSKILL_ALB_POWERSTRIKES As Long = 40

Public Const SC_NSKILL_HIB_ARBOREAL As Long = 4
Public Const SC_NSKILL_HIB_BLADES As Long = 5
Public Const SC_NSKILL_HIB_BLUNT As Long = 6
Public Const SC_NSKILL_HIB_CELTICDUAL As Long = 7
Public Const SC_NSKILL_HIB_CELTICSPEAR As Long = 8
Public Const SC_NSKILL_HIB_CREEPING As Long = 9
Public Const SC_NSKILL_HIB_CRITICAL As Long = 10
Public Const SC_NSKILL_HIB_DEMENTIA As Long = 11
Public Const SC_NSKILL_HIB_ENCHANTMENTS As Long = 12
Public Const SC_NSKILL_HIB_ENVENOM As Long = 13
Public Const SC_NSKILL_HIB_ETHEREAL As Long = 14
Public Const SC_NSKILL_HIB_LARGEWEAP As Long = 15
Public Const SC_NSKILL_HIB_LIGHT As Long = 16
Public Const SC_NSKILL_HIB_MANA As Long = 17
Public Const SC_NSKILL_HIB_MENT As Long = 18
Public Const SC_NSKILL_HIB_MUSIC As Long = 19
Public Const SC_NSKILL_HIB_NATURE As Long = 20
Public Const SC_NSKILL_HIB_NURTURE As Long = 21
Public Const SC_NSKILL_HIB_PARRY As Long = 22
Public Const SC_NSKILL_HIB_PHANTASMAL As Long = 23
Public Const SC_NSKILL_HIB_PIERCING As Long = 24
Public Const SC_NSKILL_HIB_RECURVE As Long = 25
Public Const SC_NSKILL_HIB_REGROWTH As Long = 26
Public Const SC_NSKILL_HIB_SCYTHE As Long = 27
Public Const SC_NSKILL_HIB_SHADOW As Long = 28
Public Const SC_NSKILL_HIB_SHIELD As Long = 29
Public Const SC_NSKILL_HIB_SPECTRAL As Long = 30
Public Const SC_NSKILL_HIB_STEALTH As Long = 31
Public Const SC_NSKILL_HIB_VALOR As Long = 32
Public Const SC_NSKILL_HIB_VAMPIIRIC As Long = 33
Public Const SC_NSKILL_HIB_VERD As Long = 34
Public Const SC_NSKILL_HIB_VOID As Long = 35
Public Const SC_NSKILL_HIB_AURAMANIP As Long = 36
Public Const SC_NSKILL_HIB_FISTWRAP As Long = 37
Public Const SC_NSKILL_HIB_MAGNETISM As Long = 38
Public Const SC_NSKILL_HIB_MAULERSTAFF As Long = 39
Public Const SC_NSKILL_HIB_POWERSTRIKES As Long = 40

Public Const SC_NSKILL_MID_AUGMENTATION As Long = 4
Public Const SC_NSKILL_MID_AXE As Long = 5
Public Const SC_NSKILL_MID_BATTLESONGS As Long = 6
Public Const SC_NSKILL_MID_BEASTCRAFT As Long = 7
Public Const SC_NSKILL_MID_BONEARMY As Long = 8
Public Const SC_NSKILL_MID_CAVE As Long = 9
Public Const SC_NSKILL_MID_COMPOSITE As Long = 10
Public Const SC_NSKILL_MID_CRITICAL As Long = 11
Public Const SC_NSKILL_MID_CURSING As Long = 12
Public Const SC_NSKILL_MID_DARKNESS As Long = 13
Public Const SC_NSKILL_MID_ENVENOM As Long = 14
Public Const SC_NSKILL_MID_HAMMER As Long = 15
Public Const SC_NSKILL_MID_HANDTOHAND As Long = 16
Public Const SC_NSKILL_MID_LEFTAXE As Long = 17
Public Const SC_NSKILL_MID_MENDING As Long = 18
Public Const SC_NSKILL_MID_ODIN As Long = 19
Public Const SC_NSKILL_MID_PAC As Long = 20
Public Const SC_NSKILL_MID_PARRY As Long = 21
Public Const SC_NSKILL_MID_RUNECARVING As Long = 22
Public Const SC_NSKILL_MID_SHIELD As Long = 23
Public Const SC_NSKILL_MID_SPEAR As Long = 24
Public Const SC_NSKILL_MID_STEALTH As Long = 25
Public Const SC_NSKILL_MID_STORMCALLING As Long = 26
Public Const SC_NSKILL_MID_SUMMONING As Long = 27
Public Const SC_NSKILL_MID_SUPPRESSION As Long = 28
Public Const SC_NSKILL_MID_SWORD As Long = 29
Public Const SC_NSKILL_MID_THROWN As Long = 30
Public Const SC_NSKILL_MID_WITCHCRAFT As Long = 31
Public Const SC_NSKILL_MID_HEXING As Long = 32
Public Const SC_NSKILL_MID_AURAMANIP As Long = 33
Public Const SC_NSKILL_MID_FISTWRAP As Long = 34
Public Const SC_NSKILL_MID_MAGNETISM As Long = 35
Public Const SC_NSKILL_MID_MAULERSTAFF As Long = 36
Public Const SC_NSKILL_MID_POWERSTRIKES As Long = 37

Public Const SC_NFOCUS_ALB_ALLSPELL As Long = 0
Public Const SC_NFOCUS_ALB_BODY As Long = 1
Public Const SC_NFOCUS_ALB_COLD As Long = 2
Public Const SC_NFOCUS_ALB_SERVANT As Long = 3
Public Const SC_NFOCUS_ALB_SIGHT As Long = 4
Public Const SC_NFOCUS_ALB_EARTH As Long = 5
Public Const SC_NFOCUS_ALB_FIRE As Long = 6
Public Const SC_NFOCUS_ALB_MATTER As Long = 7
Public Const SC_NFOCUS_ALB_MIND As Long = 8
Public Const SC_NFOCUS_ALB_PAIN As Long = 9
Public Const SC_NFOCUS_ALB_SPIRIT As Long = 10
Public Const SC_NFOCUS_ALB_WIND As Long = 11

Public Const SC_NFOCUS_HIB_ALLSPELL As Long = 0
Public Const SC_NFOCUS_HIB_ARBOREAL As Long = 1
Public Const SC_NFOCUS_HIB_CREEPING As Long = 2
Public Const SC_NFOCUS_HIB_ENCHANTMENTS As Long = 3
Public Const SC_NFOCUS_HIB_ETHEREAL As Long = 4
Public Const SC_NFOCUS_HIB_LIGHT As Long = 5
Public Const SC_NFOCUS_HIB_MANA As Long = 6
Public Const SC_NFOCUS_HIB_MENT As Long = 7
Public Const SC_NFOCUS_HIB_PHANTASMAL As Long = 8
Public Const SC_NFOCUS_HIB_SPECTRAL As Long = 9
Public Const SC_NFOCUS_HIB_VERD As Long = 10
Public Const SC_NFOCUS_HIB_VOID As Long = 11

Public Const SC_NFOCUS_MID_ALLSPELL As Long = 0
Public Const SC_NFOCUS_MID_BONEARMY As Long = 1
Public Const SC_NFOCUS_MID_CURSING As Long = 2
Public Const SC_NFOCUS_MID_DARKNESS As Long = 3
Public Const SC_NFOCUS_MID_RUNECARVING As Long = 4
Public Const SC_NFOCUS_MID_SUMMONING As Long = 5
Public Const SC_NFOCUS_MID_SUPPRESSION As Long = 6


Public Function DupeGemCheck(Gem1 As ComboBox, Effect1 As ComboBox, _
                             Gem2 As ComboBox, Effect2 As ComboBox, _
                             Gem3 As ComboBox, Effect3 As ComboBox, _
                             Gem4 As ComboBox, Effect4 As ComboBox) As Long
                             
'check the value of gem1 against all others

    Dim lRet As Long
    
    lRet = 0
    
    If Gem1.Text <> vbNullString Then
        If (Gem1.Text = Gem2.Text) And (Effect1.Text = Effect2.Text) Then lRet = 1
        If (Gem1.Text = Gem3.Text) And (Effect1.Text = Effect3.Text) Then lRet = 1
        If (Gem1.Text = Gem4.Text) And (Effect1.Text = Effect4.Text) Then lRet = 1
    End If
    
    DupeGemCheck = lRet
    
End Function

Public Function GetGemMats(REALM As Long, GemIndex As Long, GemType As String, GemName As String) As Gem
'working as intended :: 9/10/06
    Dim newGem As Gem
    Dim sName As String
    
    sName = LCase$(GemName)
    
    newGem.CONTAINER_NAME = GemInfo(0, GemIndex)
    newGem.CONTAINER_QUAN = 1
    newGem.CONTAINER_Index = GemIndex
    'lets first check to see whether or not its a +all skill or +all focus
    If InStr(sName, "brilliant") <> 0 Or InStr(sName, "finesse") <> 0 Then
        'brilliant and finesse gems use the same dust per realm
        
        If InStr(sName, "brilliant") <> 0 Then
            newGem.CONTAINER_QUAN = 3
        Else
            newGem.CONTAINER_QUAN = 1
        End If
            
        newGem.LIQUID1_NAME = "Draconic Fire"
        newGem.LIQUID1_Index = 1
        newGem.LIQUID2_NAME = "Mystic Energy"
        newGem.LIQUID2_Index = 6
        newGem.LIQUID3_NAME = "Treat Blood"
        newGem.LIQUID3_Index = 9
        newGem.LIQUID1_QUAN = (GemIndex * 6) + 2
        newGem.LIQUID2_QUAN = (GemIndex * 6) + 2
        newGem.LIQUID3_QUAN = (GemIndex * 6) + 2
        Select Case REALM
            Case 0  'alb
                'figure out what dust to use
                If InStr(sName, "fervor sigil") <> 0 Then
                    newGem.DUST_NAME = "Ground Blessed Undead Bone"
                    newGem.DUST_QUAN = 1
                    newGem.DUST_Index = 3
                ElseIf InStr(sName, "war sigil") <> 0 Then
                    newGem.DUST_NAME = "Ground Caer Stone"
                    newGem.DUST_QUAN = 1
                    newGem.DUST_Index = 4
                Else
                    newGem.DUST_NAME = "Ground Draconic Scales"
                    newGem.DUST_QUAN = (GemIndex * 5) + 1
                    newGem.DUST_Index = 6
                End If
            Case 1  'hib
                'figure out what dust to use
                If InStr(sName, "nature spell") <> 0 Then
                    newGem.DUST_NAME = "Fairy Dust"
                    newGem.DUST_QUAN = 1
                    newGem.DUST_Index = 2
                ElseIf InStr(sName, "war spell") <> 0 Then
                    newGem.DUST_NAME = "Unseelie Dust"
                    newGem.DUST_QUAN = 1
                    newGem.DUST_Index = 11
                Else
                    newGem.DUST_NAME = "Ground Draconic Scales"
                    newGem.DUST_QUAN = (GemIndex * 5) + 1
                    newGem.DUST_Index = 6
                End If
            Case 2  'mid
                'figure out what dust to use
                If InStr(sName, "primal rune") <> 0 Then
                    newGem.DUST_NAME = "Ground Vendo Bone"
                    newGem.DUST_QUAN = 1
                    newGem.DUST_Index = 8
                ElseIf InStr(sName, "war rune") <> 0 Then
                    newGem.DUST_NAME = "Ground Giant Bone"
                    newGem.DUST_QUAN = 1
                    newGem.DUST_Index = 7
                Else
                    newGem.DUST_NAME = "Ground Draconic Scales"
                    newGem.DUST_QUAN = (GemIndex * 5) + 1
                    newGem.DUST_Index = 6
                End If
        End Select
    Else
        'it's not all skill/focus so what is it?
        If LCase$(GemType) = "resist" Or LCase$(GemType) = "focus" Then
            newGem.DUST_QUAN = (GemIndex * 5) + 1
            newGem.LIQUID1_QUAN = GemIndex + 1
        Else
            newGem.DUST_QUAN = (GemIndex * 4) + 1
            newGem.LIQUID1_QUAN = GemIndex + 1
        End If
  
        'determine the liquid used
        If (InStr(sName, "fire") <> 0) Or (InStr(sName, "fiery") <> 0) Or (InStr(sName, "spectral spell") <> 0) Or (InStr(sName, "phantasmal arcane") <> 0) Then
            newGem.LIQUID1_NAME = "Draconic Fire"
            newGem.LIQUID1_Index = 1
            
        ElseIf (InStr(sName, "earth") <> 0) Or (InStr(sName, "oozing") <> 0) Or (InStr(sName, "aberrant") <> 0) Or (InStr(sName, "cinder") <> 0) Or (InStr(sName, "radiant") <> 0) Then
            newGem.LIQUID1_NAME = "Treant Blood"
            newGem.LIQUID1_Index = 9
            
        ElseIf (InStr(sName, "vapor") <> 0) Or (InStr(sName, "vacuous") <> 0) Or (InStr(sName, "steaming spell") <> 0) Or (InStr(sName, "steaming nature") <> 0) Or (InStr(sName, "ethereal spell") <> 0) Or (InStr(sName, "shadowy") <> 0) Or (InStr(sName, "valiant") <> 0) Or (InStr(sName, "magnetic") <> 0) Then
            newGem.LIQUID1_NAME = "Swamp Fog"
            newGem.LIQUID1_Index = 8
            
        ElseIf (InStr(sName, "air") <> 0) Or (InStr(sName, "spectral arcane") <> 0) Or (InStr(sName, "blighted primal") <> 0) Or (InStr(sName, "unholy") <> 0) Then
            newGem.LIQUID1_NAME = "Air Elemental Essence"
            newGem.LIQUID1_Index = 0
            
        ElseIf (InStr(sName, "heat") <> 0) Or (InStr(sName, "steaming fervor") <> 0) Or (InStr(sName, "mineral") <> 0) Or (InStr(sName, "clout") <> 0) Or (InStr(sName, "glacial") <> 0) Then
            newGem.LIQUID1_NAME = "Heat From an Unearthly Pyre"
            newGem.LIQUID1_Index = 4
            
        ElseIf (InStr(sName, "icy") <> 0) Or (InStr(sName, "ice") <> 0) Or (InStr(sName, "embracing") <> 0) Then
            newGem.LIQUID1_NAME = "Frost From a Wasteland"
            newGem.LIQUID1_Index = 2
            
        ElseIf (InStr(sName, "water") <> 0) Or (InStr(sName, "lightning") <> 0) Or (InStr(sName, "molten") <> 0) Or (InStr(sName, "phantasmal spell") <> 0) Or (InStr(sName, "ethereal arcane") <> 0) Then
            newGem.LIQUID1_NAME = "Leviathan Blood"
            newGem.LIQUID1_Index = 5
            
        ElseIf (InStr(sName, "dust") <> 0) Or (InStr(sName, "ashen") <> 0) Or (InStr(sName, "blighted rune") <> 0) Then
            newGem.LIQUID1_NAME = "Undead Ash and Holy Water"
            newGem.LIQUID1_Index = 10
            
        ElseIf (InStr(sName, "salt") <> 0) Or (InStr(sName, "mystic") <> 0) Then
            newGem.LIQUID1_NAME = "Mystic Energy"
            newGem.LIQUID1_Index = 6
            
        ElseIf (InStr(sName, "light") <> 0) Then
            newGem.LIQUID1_NAME = "Sun Light"
            newGem.LIQUID1_Index = 7
            
        ElseIf (InStr(sName, "blood") <> 0) Then
            newGem.LIQUID1_NAME = "Giant Blood"
            newGem.LIQUID1_Index = 3
        End If
            
        'determine the dust used
        If (InStr(sName, "essence") <> 0) Then
        
            newGem.DUST_NAME = "Essence of Life"
            newGem.DUST_Index = 1
            
        ElseIf (InStr(sName, "shielding") <> 0) Or (InStr(sName, "spell stone") <> 0) Or (InStr(sName, "sigil") <> 0) Or (InStr(sName, "rune") <> 0) Then
        
            If (InStr(sName, "chaos rune") <> 0) Then
                newGem.DUST_NAME = "Soot From Niflheim"
                newGem.DUST_Index = 10
                
            ElseIf (InStr(sName, "war rune") <> 0) Then
                newGem.DUST_NAME = "Ground Giant Bone"
                newGem.DUST_Index = 7
                
            ElseIf (InStr(sName, "primal rune") <> 0) Then
                newGem.DUST_NAME = "Ground Vendo Bone"
                newGem.DUST_Index = 8
                
            ElseIf (InStr(sName, "evocation sigil") <> 0) Then
                newGem.DUST_NAME = "Ground Cave Crystal"
                newGem.DUST_Index = 5
                
            ElseIf (InStr(sName, "fervor sigil") <> 0) Then
                newGem.DUST_NAME = "Ground Blessed Udead Bone"
                newGem.DUST_Index = 3
                
            ElseIf (InStr(sName, "war sigil") <> 0) Then
                newGem.DUST_NAME = "Ground Caer Stone"
                newGem.DUST_Index = 4
                
            ElseIf (InStr(sName, "nature spell stone") <> 0) Then
                newGem.DUST_NAME = "Fairy Dust"
                newGem.DUST_Index = 2
                
            ElseIf (InStr(sName, "war spell stone") <> 0) Then
                newGem.DUST_NAME = "Unseelie Dust"
                newGem.DUST_Index = 11
                
            ElseIf (InStr(sName, "arcane spell stone") <> 0) Then
                newGem.DUST_NAME = "Other Worldly Dust"
                newGem.DUST_Index = 9
                
            Else
                newGem.DUST_NAME = "Ground Draconic Scales"
                newGem.DUST_Index = 6
            End If
        End If
    End If
    
    GetGemMats = newGem
    
End Function

Private Sub Init_Material_Liquid()

    GemMaterial_Liquid(0) = "Air Elemental Essence"
    GemMaterial_Liquid(1) = "Draconic Fire"
    GemMaterial_Liquid(2) = "Frost From a Wasteland"
    GemMaterial_Liquid(3) = "Giant Blood"
    GemMaterial_Liquid(4) = "Heat From an Unearthly Pyre"
    GemMaterial_Liquid(5) = "Leviathan Blood"
    GemMaterial_Liquid(6) = "Mystic Energy"
    GemMaterial_Liquid(7) = "Sun Light"
    GemMaterial_Liquid(8) = "Swamp Fog"
    GemMaterial_Liquid(9) = "Treant Blood"
    GemMaterial_Liquid(10) = "Undead Ash and Holy Water"
    
End Sub

Private Sub Init_Material_Dust()

    GemMaterial_Dust(0) = "Bloodied Battlefield Dirt"
    GemMaterial_Dust(1) = "Essence of Life"
    GemMaterial_Dust(2) = "Fairy Dust"
    GemMaterial_Dust(3) = "Ground Blessed Undead Bone"
    GemMaterial_Dust(4) = "Ground Caer Stone"
    GemMaterial_Dust(5) = "Ground Cave Crystal"
    GemMaterial_Dust(6) = "Ground Draconic Scales"
    GemMaterial_Dust(7) = "Ground Giant Bone"
    GemMaterial_Dust(8) = "Ground Vendo Bone"
    GemMaterial_Dust(9) = "Other Worldy Dust"
    GemMaterial_Dust(10) = "Soot From Niflheim"
    GemMaterial_Dust(11) = "Unseelie Dust"
    
End Sub

Public Function GetContainerGem(Index As Long) As String
'working as intended :: 9/10/06
    Dim sGem As String
    
    sGem = vbNullString
    
    If Index >= 0 And Index <= 9 Then
        sGem = GemInfo(0, Index)
    End If
    
    GetContainerGem = sGem
    
End Function

Public Sub InitSpellcraftArrays(ByRef ProgressBar As Label, ByRef Status As Label)

    Status.Caption = "Initializing Spellcraft Matricies: Imbue Points"
    DoEvents
    Call Init_ImbuePoints
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Status.Caption = "Initializing Spellcraft Matricies: Multipliers"
    DoEvents
    Call Init_ImbueMultipliers
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Call Init_OCStartPercentages
    Call Init_GemQualityOCMODS
    Call Init_ItemQualityOCMODS
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
        
    Status.Caption = "Initializing Spellcraft Matricies: Gem Types"
    DoEvents
    Call Init_GemTypes
    Call Init_Resist
    Call Init_Stat
    Call Init_Hit
    Call Init_Power
    Call Init_Focus
    Call Init_Skill
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Status.Caption = "Initializing Spellcraft Matricies: Stat Points"
    DoEvents
    Call Init_nStat
    Call Init_nResist
    Call Init_nSkill
    Call Init_nFocus
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Call Init_GemInfo
    
    Status.Caption = "Initializing Spellcraft Matricies: Materials"
    DoEvents
    Call Init_Material_Dust
    Call Init_Material_Liquid
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Status.Caption = "Initializing Spellcraft Matricies: Servers"
    DoEvents
    Call Init_ServerCodes
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
End Sub

Public Function FormatCost(Cost As Long) As String
'working as intended :: 9/10/06
    Dim lPlat As Long
    Dim lGold As Long
    Dim lSilver As Long
    Dim lCopper As Long
    
    Dim sPrice As String
    
    lPlat = (Cost \ 10000000) Mod 1000
    lGold = (Cost \ 10000) Mod 1000
    lSilver = (Cost \ 100) Mod 100
    lCopper = Cost Mod 100
    
    If lPlat > 0 Then sPrice = sPrice & lPlat & "p "
       
    If lGold > 0 Then sPrice = sPrice & lGold & "g "
    
    If lSilver > 0 Then sPrice = sPrice & lSilver & "s "
    
    If lCopper > 0 Then sPrice = sPrice & lCopper & "c"
    
    FormatCost = sPrice
    
End Function

Public Function SumColumn(iValueIndex As Long, BonusOption As Integer)
'working as intended :: 9/10/06
    '# ---------------------------------------------------------------------------
    '# J VALUE Index
    '# ---------------------------------------------------------------------------
    '# core toon:
    '# 1 base, 2 creation, 3 level
    '# ---------------------------------------------------------------------------
    '# armor locations:
    '# 4 head, 5 chest, 6 arms, 7 gloves, 8 pants, 9 boots
    '# ---------------------------------------------------------------------------
    '# accessory locations:
    '# 10 cloak, 11 neck, 12 gem, 13 belt, 14 r-ring, 15 l-ring
    '# 16 r-bracer, 17 l-bracer
    '# ---------------------------------------------------------------------------
    '# weapon locations:
    '# 18 r-hand, 19 l-hand, 20 2-hand, 21 ranged
    '# ---------------------------------------------------------------------------
    '# BONUSOPTION Flag values:
    '#  0 = total bonus: normal spellcraft mode
    '#  1 = distance to cap: normal spellcraft mode
    
    Dim Ctr As Long
    
    Dim Lo As Long
    Dim Hi As Long
    
    Dim lResult As Long
    
    Select Case BonusOption
        Case BO_SCBONUS
            Lo = SM_LOC_HEAD
            Hi = SM_LOC_MYTHICAL
        Case BO_TOCAP
            Lo = SM_LOC_HEAD
            Hi = SM_LOC_MYTHICAL
    End Select
    
    For Ctr = Lo To Hi
        lResult = lResult + TOON.STAT_MATRIX(iValueIndex, Ctr)
    Next Ctr
    
    SumColumn = lResult
            
End Function

Public Function TranslateLocationName(sName As String) As Integer
'working as intended :: 12/27/06
    Dim lCode As Integer
    
    Select Case LCase$(sName)
        Case "chest"
            lCode = WS_DOLL_CHEST
        Case "arms"
            lCode = WS_DOLL_ARMS
        Case "jewel"
            lCode = WS_DOLL_GEM
        Case "left ring"
            lCode = WS_DOLL_LRING
        Case "left wrist"
            lCode = WS_DOLL_LWRIST
        Case "legs"
            lCode = WS_DOLL_LEGS
        Case "right hand"
            lCode = WS_DOLL_RHAND
        Case "left hand"
            lCode = WS_DOLL_LHAND
        Case "2 handed", "two-handed"
            lCode = WS_DOLL_2HAND
        Case "ranged"
            lCode = WS_DOLL_RANGED
        Case "feet"
            lCode = WS_DOLL_FEET
        Case "right wrist"
            lCode = WS_DOLL_RWRIST
        Case "right ring"
            lCode = WS_DOLL_RRING
        Case "belt"
            lCode = WS_DOLL_WAIST
        Case "hands"
            lCode = WS_DOLL_HANDS
        Case "head"
            lCode = WS_DOLL_HEAD
        Case "cloak"
            lCode = WS_DOLL_CLOAK
        Case "neck"
            lCode = WS_DOLL_NECK
        Case "spare"
            lCode = WS_DOLL_RIGHTSPARE
    End Select
    
    TranslateLocationName = lCode
    
End Function

Public Function TranslateGemSelectName(GemType As String, DPSlot As ComboBox) As Long
'working as intended :: 9/10/06

    Dim lCode As Long
    
    Dim sTranslated As String
        
    Select Case LCase$(GemType)
        Case "unused"
            sTranslated = ""
        Case "stat"
            sTranslated = "Stat"
        Case "resist"
            sTranslated = "Resist"
        Case "hits"
            sTranslated = "Hits"
        Case "power"
            sTranslated = "Power"
        Case "focus"
            sTranslated = "Focus"
        Case "skill"
            sTranslated = "Skill"
        Case "cap increase"
            sTranslated = "Cap Increase"
        Case "other bonus"  'korts and loki's equiv of loto's toa bonus
            sTranslated = "ToA Bonus"
        Case "pve bonus", "pve"
            sTranslated = "PvE Bonus"
        Case "charged effect", "charge"
            sTranslated = "Charged Effect"
        Case "reactive effect", "reactive"
            sTranslated = "Reactive Effect"
        Case "offensive effect", "proc"
            sTranslated = "Offensive Effect"
    End Select
    
    lCode = GetListindexByString(sTranslated, DPSlot)
    
    TranslateGemSelectName = lCode
    
End Function

Public Function TranslateGemEffectName(GemEffect As String, EffectList As ComboBox) As Long
'working as intended :: 9/10/06
    Dim Ctr As Long
    Dim lCode As Long
    Dim bFound As Boolean
    Dim sTrans As String
    Dim sTemp As String
    
    'first we strip some stuff to make the matching work better
    sTemp = LCase$(GemEffect)
    sTemp = Trim(RemoveString(sTemp, "(pve)"))
    sTemp = Trim(RemoveString(sTemp, "(non-frontier)"))
    sTemp = Trim(RemoveString(sTemp, "focus"))
    sTemp = Trim(RemoveString(sTemp, "resist"))
    sTemp = Trim(RemoveString(sTemp, "bonus to"))
    sTemp = Trim(RemoveString(sTemp, "attribute"))
    sTemp = Trim(RemoveString(sTemp, "bonus cap"))
    sTemp = Trim(RemoveString(sTemp, "(af)"))
    sTemp = Trim(RemoveString(sTemp, "affinity"))
    sTemp = Trim(RemoveString(sTemp, "bonus"))
        
    For Ctr = 0 To EffectList.ListCount - 1
        If sTemp = LCase$(EffectList.list(Ctr)) Then
            bFound = True
            Exit For
        End If
    Next Ctr
    
    If bFound = True Then
        lCode = Ctr
    Else
        lCode = LISTINDEX_NOT_FOUND
    
        'we try matching based on a bruteforce list of known differences
        
        Select Case sTemp
            Case "af", "armour factor"
                sTrans = "armor factor"
                
            Case "all dual wielding skill", "all dual wield", "all dual wield skill", "all dual wielding"
                sTrans = "all dual wield skills"
                
            Case "all spell lines focus"
                sTrans = "all spell lines"
                
            Case "all magic skill", "all magic", "all magic skills"
                sTrans = "all magic skills"
                
            Case "all melee weapon skills", "all melee weapon skill", "all melee skill"
                sTrans = "all melee skills"
                
            Case "hitpoints"
                sTrans = "hits"
                
            Case "casting speed"
                sTrans = "spell haste"
                
            Case "duration of spells"
                sTrans = "spell duration"
            
            Case "healing"
                sTrans = "healing effectiveness"
                
            Case "magic damage"
                sTrans = "spell damage"
                
            Case "melee speed"
                sTrans = "melee combat speed"
            
            Case "pierce"
                If LCase$(GemEffect) = "resist pierce" Then
                    sTrans = "spell pierce"
                Else
                    sTrans = sTemp
                End If
                
            Case "stat debuff spell effectiveness", "debuff"
                sTrans = "stat debuff effectiveness"
                
            Case "stat enhancement spell effectiveness", "buff"
                sTrans = "stat buff effectiveness"
                
            Case "style damage"
                sTrans = "melee style damage"
                
            Case "negative effect duration reduction"
                sTrans = "negative effect duration"
                
            Case "power pool", "power percentage"
                sTrans = "% power pool"
                
            Case "light magic"
                sTrans = "light"
                
            Case "mana magic"
                sTrans = "mana"
                
            Case "void magic"
                sTrans = "void"
                
            Case "enchantments"
                sTrans = "enchantment"
        End Select
        
        If Len(sTrans) = 0 Then
            'try one more time before giving up
            Select Case LCase$(GemEffect)
                Case "all focus bonus"
                    sTrans = "all spell lines"
            End Select
        End If
        
        If Len(sTrans) <> 0 Then lCode = GetListindexByString(sTrans, EffectList)
        
    End If
        
    TranslateGemEffectName = lCode
    
End Function

Public Function TranslateKortCap(Effect As String, DPSlot As ComboBox) As Long
'working as intended :: 9/10/06
    Dim lCode As Long
    
    lCode = GetListindexByString(Effect, DPSlot)
  
    TranslateKortCap = lCode
    
End Function

Public Function TranslateKortProc(Effect As String, DPSlot As ComboBox) As Long
'working as intended :: 9/10/06
    Dim lCode As Long
    
    Dim sTranslated As String
    
    Select Case LCase$(Effect)
        Case "damage over time"
            sTranslated = "Damage Over Time"
        Case "dex/qui debuff"
            sTranslated = "Dex/Qui Debuff"
        Case "direct damage (cold)"
            sTranslated = "Direct Damage (Cold)"
        Case "direct damage (energy)"
            sTranslated = "Direct Damage (Energy)"
        Case "direct damage (fire)"
            sTranslated = "Direct Damage (Fire"
        Case "direct damage (spirit)"
            sTranslated = "Direct Damage (Spirit)"
        Case "lifedrain"
            sTranslated = "Lifedrain"
        Case "self af buff"
            sTranslated = "Self AF Buff"
        Case "self acuity buff"
            sTranslated = "Self Acuity Buff"
        Case "self damage add buff"
            sTranslated = "Self Damage Add"
        Case "self damage shield buff"
            sTranslated = "Self Damage Shield"
        Case "self melee haste buff"
            sTranslated = "Self Melee Haste"
        Case "self melee health buffer"
            sTranslated = "Self Melee Health Buffer"
        Case "str/con debuff"
            sTranslated = "Str/Con Debuff"
        Case "unique effect..."
            sTranslated = "Unique Effect"
    End Select

    If sTranslated <> vbNullString Then
        lCode = GetListindexByString(sTranslated, DPSlot)
    End If
    
    TranslateKortProc = lCode
    
End Function

Public Function TranslateKortPve(Effect As String, DPSlot As ComboBox) As Long
'working as intended :: 9/10/06
   
    Dim lCode As Long
    
    Dim sBuf As String
    Dim sTranslated As String
    
    sBuf = RemoveString(Effect, ".")
    sBuf = RemoveString(sBuf, "\")
    sBuf = RemoveString(sBuf, "/")
        
    Select Case LCase$(sBuf)
        Case "arrow recovery"
            sTranslated = "Arcane Siphon" 'Arrow Recovery"
        
        Case "arcane siphon"
            sTranslated = "Arcane Siphon"
            
        Case "bladeturn reinforcement"
            sTranslated = "Bladeturn Reinforcement"
            
        Case "block"
            sTranslated = "Block"
            
        Case "concentration"
            sTranslated = "Concentration"
            
        Case "damage reduction"
            sTranslated = "Damage Reduction"
            
        Case "death experience loss reduction"
            sTranslated = "Death Experience Loss Reduction"
            
        Case "defensive"
            sTranslated = "Defensive"
            
        Case "evade"
            sTranslated = "Evade"
            
        Case "negative effect duration reduction"
            sTranslated = "Negative Effect Duration"
            
        Case "parry"
            sTranslated = "Parry"
            
        Case "piece ablative"
            sTranslated = "Piece Ablative"
            
        Case "reactionary style damage"
            sTranslated = "Reactionary Style Damage"
            
        Case "spell power cost reduction"
            sTranslated = "Spell Power Cost Reduction"
            
        Case "style cost reduction"
            sTranslated = "Style Cost Reduction"
            
        Case "to hit"
            sTranslated = "To-Hit"
            
        Case "unique pve bonus"
            sTranslated = "Unique Bonus"
    End Select
    
    If sTranslated <> vbNullString Then
        lCode = GetListindexByString(sTranslated, DPSlot)
    End If
    
    TranslateKortPve = lCode
            
End Function

Public Function TranslateKortToa(Effect As String, DPSlot As ComboBox) As Long
'working as intended :: 9/10/06
    Dim lCode As Long
    
    Dim sTranslated As String
    
    Select Case LCase$(Effect)
        Case "% power pool"
            sTranslated = "% Power Pool"
        Case "af"
            sTranslated = "Armor Factor"
        Case "archery damage"
            sTranslated = "Archery Damage"
        Case "archery range"
            sTranslated = "Archery Range"
        Case "archery speed"
            sTranslated = "Archery Speed"
        Case "casting speed"
            sTranslated = "Spell Haste"
        Case "duration of spells"
            sTranslated = "Spell Duration"
        Case "fatigue"
            sTranslated = "Fatigue"
        Case "healing effectiveness"
            sTranslated = "Healing Effectiveness"
        Case "melee combat speed"
            sTranslated = "Melee Combat Speed"
        Case "melee damage"
            sTranslated = "Melee Damage"
        Case "spell damage"
            sTranslated = "Spell Damage"
        Case "spell piercing"
            sTranslated = "Spell Pierce"
        Case "spell range"
            sTranslated = "Spell Range"
        Case "stat buff effectiveness"
            sTranslated = "Stat Buff Effectiveness"
        Case "stat debuff effectiveness"
            sTranslated = "Stat Debuff Effectiveness"
        Case "style damage"
            sTranslated = "Melee Style Damage"
        Case "unique bonus..."
            sTranslated = "Unique Bonus"
    End Select
    
    If sTranslated <> vbNullString Then
        lCode = GetListindexByString(sTranslated, DPSlot)
    End If
    
    TranslateKortToa = lCode
    
End Function

Public Function TranslateGemAmountValue(AmountList As ComboBox, GemAmount As String) As Long
'working as intended :: 9/10/06
    Dim Ctr As Long
    
    Dim lCode As Long
    
    For Ctr = 0 To AmountList.ListCount - 1
        If Val(GemAmount) = Val(AmountList.list(Ctr)) Then Exit For
    Next Ctr
    
    lCode = Ctr
    
    TranslateGemAmountValue = lCode
    
End Function

Public Function TranslateLocationToMatrix(Location As Integer) As Long
'working as intended :: 9/10/06
    Dim jValueIndex As Long

    Select Case Location
        Case WS_DOLL_CHEST  'chest
            jValueIndex = SM_LOC_CHEST
        Case WS_DOLL_ARMS  'arms
            jValueIndex = SM_LOC_ARMS
        Case WS_DOLL_GEM  'gem
            jValueIndex = SM_LOC_GEM
        Case WS_DOLL_LRING  'left ring
            jValueIndex = SM_LOC_LRING
        Case WS_DOLL_LWRIST  'left wrist
            jValueIndex = SM_LOC_LBRACER
        Case WS_DOLL_LEGS  'legs
            jValueIndex = SM_LOC_LEGS
        Case WS_DOLL_RHAND  'right hand
            jValueIndex = SM_LOC_RHAND
        Case WS_DOLL_LHAND  'left hand
            jValueIndex = SM_LOC_LHAND
        Case WS_DOLL_2HAND  '2handed
            jValueIndex = SM_LOC_2HAND
        Case WS_DOLL_RANGED  'ranged
            jValueIndex = SM_LOC_RANGED
        Case WS_DOLL_FEET 'feet
            jValueIndex = SM_LOC_FEET
        Case WS_DOLL_RWRIST 'right wrist
            jValueIndex = SM_LOC_RBRACER
        Case WS_DOLL_RRING 'right ring
            jValueIndex = SM_LOC_RRING
        Case WS_DOLL_WAIST 'waist
            jValueIndex = SM_LOC_BELT
        Case WS_DOLL_HANDS 'hands
            jValueIndex = SM_LOC_HANDS
        Case WS_DOLL_HEAD 'head
            jValueIndex = SM_LOC_HEAD
        Case WS_DOLL_CLOAK 'cloak
            jValueIndex = SM_LOC_CLOAK
        Case WS_DOLL_NECK 'neck
            jValueIndex = SM_LOC_NECK
        Case WS_DOLL_RIGHTSPARE 'right hand spare
            jValueIndex = SM_LOC_RHAND
        Case WS_DOLL_LEFTSPARE 'left hand spare
            jValueIndex = SM_LOC_LHAND
        Case WS_DOLL_2HANDSPARE '2hand spare
            jValueIndex = SM_LOC_2HAND
        Case WS_DOLL_RANGEDSPARE 'ranged spare
            jValueIndex = SM_LOC_RANGED
        Case WS_DOLL_MYTHICAL 'mythical
            jValueIndex = SM_LOC_MYTHICAL
        Case (WS_DOLL_CHEST + 100)
            jValueIndex = SM_LOC_CHEST_5
        Case (WS_DOLL_ARMS + 100)
            jValueIndex = SM_LOC_ARMS_5
        Case (WS_DOLL_LEGS + 100)
            jValueIndex = SM_LOC_LEGS_5
        Case (WS_DOLL_FEET + 100)
            jValueIndex = SM_LOC_FEET_5
        Case (WS_DOLL_HANDS + 100)
            jValueIndex = SM_LOC_HANDS_5
        Case (WS_DOLL_HEAD + 100)
            jValueIndex = SM_LOC_HEAD_5
        Case (WS_DOLL_RHAND + 100)
            jValueIndex = SM_LOC_RHAND_5
        Case (WS_DOLL_LHAND + 100)
            jValueIndex = SM_LOC_LHAND_5
        Case (WS_DOLL_2HAND + 100)
            jValueIndex = SM_LOC_2HAND_5
        Case (WS_DOLL_RANGED + 100)
            jValueIndex = SM_LOC_RANGED_5
        Case (WS_DOLL_RIGHTSPARE + 100)
            jValueIndex = SM_LOC_RHAND_5
        Case (WS_DOLL_LEFTSPARE + 100)
            jValueIndex = SM_LOC_LHAND_5
        Case (WS_DOLL_2HANDSPARE + 100)
            jValueIndex = SM_LOC_2HAND_5
        Case (WS_DOLL_RANGEDSPARE + 100)
            jValueIndex = SM_LOC_RANGED_5
    End Select

    TranslateLocationToMatrix = jValueIndex
    
End Function

Public Function TranslateEffectToMatrix(GemType As String, GemEffect As String, ByRef Utility As Single) As Long
'working as intended :: 9/10/06
    Dim iValueIndex As Long
    
    Select Case LCase$(GemType)
        Case "stat"
            Utility = (2 / 3)
            Select Case LCase$(GemEffect)
                Case "strength"
                    iValueIndex = SM_STR
                Case "constitution"
                    iValueIndex = SM_CON
                Case "dexterity"
                    iValueIndex = SM_DEX
                Case "quickness"
                    iValueIndex = SM_QUI
                Case "intelligence"
                    iValueIndex = SM_INT
                Case "empathy"
                    iValueIndex = SM_EMP
                Case "piety"
                    iValueIndex = SM_PIE
                Case "charisma"
                    iValueIndex = SM_CHA
                Case "acuity"
                    iValueIndex = SM_ACU
            End Select
        Case "power"
            Utility = 2#
            iValueIndex = SM_POW
        Case "hits"
            Utility = 0.25
            iValueIndex = SM_HIT
        Case "resist"
            Utility = 2#
            Select Case LCase$(GemEffect)
                Case "crush"
                    iValueIndex = SM_CRUSH_RESIST
                Case "slash"
                    iValueIndex = SM_SLASH_RESIST
                Case "thrust"
                    iValueIndex = SM_THRUST_RESIST
                Case "body"
                    iValueIndex = SM_BODY_RESIST
                Case "cold"
                    iValueIndex = SM_COLD_RESIST
                Case "energy"
                    iValueIndex = SM_ENERGY_RESIST
                Case "heat"
                    iValueIndex = SM_HEAT_RESIST
                Case "matter"
                    iValueIndex = SM_MATTER_RESIST
                Case "spirit"
                    iValueIndex = SM_SPIRIT_RESIST
            End Select
        Case "skill"
            Utility = 5#
            If LCase$(GemEffect) = "all melee skills" Then
                iValueIndex = SM_ALLMELEE
            ElseIf LCase$(GemEffect) = "all magic skills" Then
                iValueIndex = SM_ALLMAGIC
            ElseIf LCase$(GemEffect) = "all dual wield skills" Then
                iValueIndex = SM_ALLDUALWIELD
            ElseIf LCase$(GemEffect) = "all archery skills" Then
                iValueIndex = SM_ALLARCHERY
            Else
                Select Case TOON.REALM
                    Case REALM_ALBION  'alb
                        Select Case LCase$(GemEffect)
                            Case "crush"
                                iValueIndex = SM_CRUSH
                            Case "slash"
                                iValueIndex = SM_SLASH
                            Case "thrust"
                                iValueIndex = SM_THRUST
                            Case "dual wield"
                                iValueIndex = SM_DUALWIELD
                            Case "crossbow"
                                iValueIndex = SM_CROSSBOW
                            Case "polearm"
                                iValueIndex = SM_POLEARM
                            Case "two handed"
                                iValueIndex = SM_TWOHANDED
                            Case "staff"
                                iValueIndex = SM_STAFF
                            Case "flexible"
                                iValueIndex = SM_FLEXIBLE
                            Case "archery"
                                iValueIndex = SM_LONGBOW
                            Case "cold magic"
                                iValueIndex = SM_COLD
                            Case "earth magic"
                                iValueIndex = SM_EARTH
                            Case "fire magic"
                                iValueIndex = SM_FIRE
                            Case "wind magic"
                                iValueIndex = SM_WIND
                            Case "matter magic"
                                iValueIndex = SM_MATTER
                            Case "mind magic"
                                iValueIndex = SM_MIND
                            Case "body magic"
                                iValueIndex = SM_BODY
                            Case "spirit magic"
                                iValueIndex = SM_SPIRIT
                            Case "soulrending"
                                iValueIndex = SM_SOULRENDING
                            Case "death servant"
                                iValueIndex = SM_DEATHSERVANT
                            Case "deathsight"
                                iValueIndex = SM_DEATHSIGHT
                            Case "painworking"
                                iValueIndex = SM_PAINWORKING
                            Case "instruments"
                                iValueIndex = SM_INSTRUMENTS
                            Case "enhancement"
                                iValueIndex = SM_ENHANCEMENT
                            Case "rejuvenation"
                                iValueIndex = SM_REJUVENATION
                            Case "smite"
                                iValueIndex = SM_SMITE
                            Case "chants"
                                iValueIndex = SM_CHANTS
                            Case "stealth"
                                iValueIndex = SM_STEALTH
                            Case "critical strike"
                                iValueIndex = SM_CRITICALSTRIKE
                            Case "envenom"
                                iValueIndex = SM_ENVENOM
                            Case "shield"
                                iValueIndex = SM_SHIELD
                            Case "parry"
                                iValueIndex = SM_PARRY
                            Case "aura manipulation"
                                iValueIndex = SM_LOTM_AURAMANIP
                            Case "fist wraps"
                                iValueIndex = SM_LOTM_FISTWRAP
                            Case "magnetism"
                                iValueIndex = SM_LOTM_MAGNETISM
                            Case "mauler staff"
                                iValueIndex = SM_LOTM_MAULERSTAFF
                            Case "power strikes"
                                iValueIndex = SM_LOTM_POWERSTRIKES
                        End Select
                    Case REALM_HIBERNIA  'hib
                        Select Case LCase$(GemEffect)
                            Case "blades"
                                iValueIndex = SM_BLADES
                            Case "blunt"
                                iValueIndex = SM_BLUNT
                            Case "piercing"
                                iValueIndex = SM_PIERCE
                            Case "large weaponry"
                                iValueIndex = SM_LARGEWEAP
                            Case "celtic spear"
                                iValueIndex = SM_CELTICSPEAR
                            Case "scythe"
                                iValueIndex = SM_SCYTHE
                            Case "celtic dual"
                                iValueIndex = SM_CELTICDUAL
                            Case "archery"
                                iValueIndex = SM_RECURVE
                            Case "arboreal path"
                                iValueIndex = SM_ARBOREAL
                            Case "creeping path"
                                iValueIndex = SM_CREEPING
                            Case "verdant path"
                                iValueIndex = SM_VERDANT
                            Case "light"
                                iValueIndex = SM_LIGHT
                            Case "mana"
                                iValueIndex = SM_MANA
                            Case "void"
                                iValueIndex = SM_VOID
                            Case "mentalism"
                                iValueIndex = SM_MENTALISM
                            Case "enchantments"
                                iValueIndex = SM_ENCHANTMENTS
                            Case "dementia"
                                iValueIndex = SM_DEMENTIA
                            Case "vampiiric embrace"
                                iValueIndex = SM_VAMPEMBRACE
                            Case "shadow mastery"
                                iValueIndex = SM_SHADOWMASTERY
                            Case "phantasmal wail"
                                iValueIndex = SM_PHANTASMALWAIL
                            Case "spectral guard"
                                iValueIndex = SM_SPECTRALGUARD
                            Case "ethereal shriek"
                                iValueIndex = SM_ETHEREALSHRIEK
                            Case "nurture"
                                iValueIndex = SM_NURTURE
                            Case "nature"
                                iValueIndex = SM_NATURE
                            Case "regrowth"
                                iValueIndex = SM_REGROWTH
                            Case "music"
                                iValueIndex = SM_MUSIC
                            Case "valor"
                                iValueIndex = SM_VALOR
                            Case "stealth"
                                iValueIndex = SM_STEALTH
                            Case "critical strike"
                                iValueIndex = SM_CRITICALSTRIKE
                            Case "envenom"
                                iValueIndex = SM_ENVENOM
                            Case "shield"
                                iValueIndex = SM_SHIELD
                            Case "parry"
                                iValueIndex = SM_PARRY
                            Case "aura manipulation"
                                iValueIndex = SM_LOTM_AURAMANIP
                            Case "fist wraps"
                                iValueIndex = SM_LOTM_FISTWRAP
                            Case "magnetism"
                                iValueIndex = SM_LOTM_MAGNETISM
                            Case "mauler staff"
                                iValueIndex = SM_LOTM_MAULERSTAFF
                            Case "power strikes"
                                iValueIndex = SM_LOTM_POWERSTRIKES
                        End Select
                    Case REALM_MIDGARD  'mid
                        Select Case LCase$(GemEffect)
                            Case "hammer"
                                iValueIndex = SM_HAMMER
                            Case "axe"
                                iValueIndex = SM_AXE
                            Case "sword"
                                iValueIndex = SM_SWORD
                            Case "spear"
                                iValueIndex = SM_SPEAR
                            Case "hand to hand"
                                iValueIndex = SM_HANDTOHAND
                            Case "archery"
                                iValueIndex = SM_COMPOSITEBOW
                            Case "thrown weapons"
                                iValueIndex = SM_THROWN
                            Case "left axe"
                                iValueIndex = SM_LEFTAXE
                            Case "bone army"
                                iValueIndex = SM_BONEARMY
                            Case "darkness"
                                iValueIndex = SM_DARKNESS
                            Case "suppression"
                                iValueIndex = SM_SUPPRESSION
                            Case "mending"
                                iValueIndex = SM_MENDING
                            Case "augmentation"
                                iValueIndex = SM_AUGMENTATION
                            Case "beastcraft"
                                iValueIndex = SM_BEASTCRAFT
                            Case "runecarving"
                                iValueIndex = SM_RUNECARVING
                            Case "cave magic"
                                iValueIndex = SM_CAVEMAGIC
                            Case "battlesongs"
                                iValueIndex = SM_BATTLESONGS
                            Case "summoning"
                                iValueIndex = SM_SUMMONING
                            Case "stormcalling"
                                iValueIndex = SM_STORMCALLING
                            Case "odin's will"
                                iValueIndex = SM_ODINSWILL
                            Case "cursing"
                                iValueIndex = SM_CURSING
                            Case "hexing"
                                iValueIndex = SM_HEXING
                            Case "pacification"
                                iValueIndex = SM_PACIFICATION
                            Case "witchcraft"
                                iValueIndex = SM_WITCHCRAFT
                            Case "stealth"
                                iValueIndex = SM_STEALTH
                            Case "critical strike"
                                iValueIndex = SM_CRITICALSTRIKE
                            Case "envenom"
                                iValueIndex = SM_ENVENOM
                            Case "shield"
                                iValueIndex = SM_SHIELD
                            Case "parry"
                                iValueIndex = SM_PARRY
                            Case "aura manipulation"
                                iValueIndex = SM_LOTM_AURAMANIP
                            Case "fist wraps"
                                iValueIndex = SM_LOTM_FISTWRAP
                            Case "magnetism"
                                iValueIndex = SM_LOTM_MAGNETISM
                            Case "mauler staff"
                                iValueIndex = SM_LOTM_MAULERSTAFF
                            Case "power strikes"
                                iValueIndex = SM_LOTM_POWERSTRIKES
                        End Select
                End Select
            End If
        Case "focus"
            Utility = -1
            If LCase$(GemEffect) = "all spell lines" Then
                iValueIndex = SM_ALLFOCUS
            Else
                Select Case TOON.REALM
                    Case REALM_ALBION  'alb
                        Select Case LCase$(GemEffect)
                            Case "body magic"
                                iValueIndex = SM_BODY_FOCUS
                            Case "cold magic"
                                iValueIndex = SM_COLD_FOCUS
                            Case "death servant"
                                iValueIndex = SM_DEATHSERVANT_FOCUS
                            Case "deathsight"
                                iValueIndex = SM_DEATHSIGHT_FOCUS
                            Case "earth magic"
                                iValueIndex = SM_EARTH_FOCUS
                            Case "fire magic"
                                iValueIndex = SM_FIRE_FOCUS
                            Case "matter magic"
                                iValueIndex = SM_MATTER_FOCUS
                            Case "mind magic"
                                iValueIndex = SM_MIND_FOCUS
                            Case "painworking"
                                iValueIndex = SM_PAINWORKING_FOCUS
                            Case "spirit magic"
                                iValueIndex = SM_SPIRIT_FOCUS
                            Case "wind magic"
                                iValueIndex = SM_WIND_FOCUS
                        End Select
                    Case REALM_HIBERNIA  'hib
                         Select Case LCase$(GemEffect)
                            Case "arboreal path"
                                iValueIndex = SM_ARBOREAL_FOCUS
                            Case "creeping path"
                                iValueIndex = SM_CREEPING_FOCUS
                            Case "verdant path"
                                iValueIndex = SM_VERDANT_FOCUS
                            Case "light"
                                iValueIndex = SM_LIGHT_FOCUS
                            Case "mana"
                                iValueIndex = SM_MANA_FOCUS
                            Case "void"
                                iValueIndex = SM_VOID_FOCUS
                            Case "enchantments"
                                iValueIndex = SM_ENCHANTMENT_FOCUS
                            Case "mentalism"
                                iValueIndex = SM_MENTALISM_FOCUS
                            Case "ethereal shriek"
                                iValueIndex = SM_ETHEREAL_FOCUS
                            Case "phantasmal wail"
                                iValueIndex = SM_PHANTASMAL_FOCUS
                            Case "spectral guard"
                                iValueIndex = SM_SPECTRAL_FOCUS
                        End Select
                    Case REALM_MIDGARD  'mid
                        Select Case LCase$(GemEffect)
                            Case "bone army"
                                iValueIndex = SM_BONE_FOCUS
                            Case "cursing"
                                iValueIndex = SM_CURSING_FOCUS
                            Case "darkness"
                                iValueIndex = SM_DARKNESS_FOCUS
                            Case "runecarving"
                                iValueIndex = SM_RUNECARVING_FOCUS
                            Case "summoning"
                                iValueIndex = SM_SUMMONING_FOCUS
                            Case "suppression"
                                iValueIndex = SM_SUPPRESSION_FOCUS
                            Case "hexing"
                                iValueIndex = SM_HEXING_FOCUS
                            Case "witchcraft"
                                iValueIndex = SM_WITCHCRAFT_FOCUS
                        End Select
                End Select
            End If
        Case "cap increase"
            Utility = 2#
            Select Case LCase$(GemEffect)
                Case "strength"
                    iValueIndex = SM_STR_CAP
                Case "constitution"
                    iValueIndex = SM_CON_CAP
                Case "dexterity"
                    iValueIndex = SM_DEX_CAP
                Case "quickness"
                    iValueIndex = SM_QUI_CAP
                Case "intelligence"
                    iValueIndex = SM_INT_CAP
                Case "empathy"
                    iValueIndex = SM_EMP_CAP
                Case "piety"
                    iValueIndex = SM_PIE_CAP
                Case "charisma"
                    iValueIndex = SM_CHA_CAP
                Case "hits"
                    iValueIndex = SM_HIT_CAP
                Case "power"
                    iValueIndex = SM_POW_CAP
                Case "acuity"
                    iValueIndex = SM_ACU_CAP
            End Select
        Case "lotm bonus"
            Select Case LCase$(GemEffect)
                Case "encumberence increase"
                    iValueIndex = SM_LOTM_ENC
                Case "essence resist"
                    iValueIndex = SM_LOTM_ESSENCE_RESIST
                Case "coin drop increase"
                    iValueIndex = SM_LOTM_COIN_DROP_INCREASE
                Case "rez sick reduction"
                    iValueIndex = SM_LOTM_REZ_SICK_REDUCTION
                Case "safe fall"
                    iValueIndex = SM_LOTM_SAFE_FALL
                Case "arcane siphon" 'arrow recovery"
                    iValueIndex = SM_LOTM_ARROW_RECOVERY
                Case "str cap"
                    iValueIndex = SM_LOTM_STR_CAP
                Case "str cap (+str)"
                    iValueIndex = SM_LOTM_STR_CAP_PLUS
                Case "con cap"
                    iValueIndex = SM_LOTM_CON_CAP
                Case "con cap (+con)"
                    iValueIndex = SM_LOTM_CON_CAP_PLUS
                Case "dex cap"
                    iValueIndex = SM_LOTM_DEX_CAP
                Case "dex cap (+dex)"
                    iValueIndex = SM_LOTM_DEX_CAP_PLUS
                Case "qui cap"
                    iValueIndex = SM_LOTM_QUI_CAP
                Case "qui cap (+qui)"
                    iValueIndex = SM_LOTM_QUI_CAP_PLUS
                Case "acuity cap)"
                    iValueIndex = SM_LOTM_ACU_CAP
                Case "acuity cap (+acu)"
                    iValueIndex = SM_LOTM_ACU_CAP_PLUS
                Case "crush cap"
                    iValueIndex = SM_LOTM_CRUSH_RESIST_CAP
                Case "crush cap (+crush)"
                    iValueIndex = SM_LOTM_CRUSH_RESIST_CAP_PLUS
                Case "slash cap"
                    iValueIndex = SM_LOTM_SLASH_RESIST_CAP
                Case "slash cap (+slash)"
                    iValueIndex = SM_LOTM_SLASH_RESIST_CAP_PLUS
                Case "thrust cap"
                    iValueIndex = SM_LOTM_THRUST_RESIST_CAP
                Case "thrust cap (+thrust)"
                    iValueIndex = SM_LOTM_THRUST_RESIST_CAP_PLUS
                Case "body cap"
                    iValueIndex = SM_LOTM_BODY_RESIST_CAP
                Case "body cap (+body)"
                    iValueIndex = SM_LOTM_BODY_RESIST_CAP_PLUS
                Case "cold cap"
                    iValueIndex = SM_LOTM_COLD_RESIST_CAP
                Case "cold cap (+cold)"
                    iValueIndex = SM_LOTM_COLD_RESIST_CAP_PLUS
                Case "energy cap"
                    iValueIndex = SM_LOTM_ENERGY_RESIST_CAP
                Case "energy cap (+energy)"
                    iValueIndex = SM_LOTM_ENERGY_RESIST_CAP_PLUS
                Case "heat cap"
                    iValueIndex = SM_LOTM_HEAT_RESIST_CAP
                Case "heat cap (+heat)"
                    iValueIndex = SM_LOTM_HEAT_RESIST_CAP_PLUS
                Case "matter cap"
                    iValueIndex = SM_LOTM_MATTER_RESIST_CAP
                Case "matter cap (+matter)"
                    iValueIndex = SM_LOTM_MATTER_RESIST_CAP_PLUS
                Case "spirit cap"
                    iValueIndex = SM_LOTM_SPIRIT_RESIST_CAP
                Case "spirit cap (+spirit)"
                    iValueIndex = SM_LOTM_SPIRIT_RESIST_CAP_PLUS
                Case "realm point bonus"
                    iValueIndex = SM_LOTM_REALMPOINT_INCREASE
                Case "parry"
                    iValueIndex = SM_LOTM_PARRY
                Case "evade"
                    iValueIndex = SM_LOTM_EVADE
                Case "block"
                    iValueIndex = SM_LOTM_BLOCK
                Case "siege damage decrease"
                    iValueIndex = SM_LOTM_SIEGE_DAMAGE_REDUCTION
                Case "spell level increase"
                    iValueIndex = SM_LOTM_SPELL_LEVEL_INCREASE
                Case "crowd control decrease"
                    iValueIndex = SM_LOTM_CROWDCONTROL_REDUCTION
                Case "damage increase"
                    iValueIndex = SM_LOTM_DAMAGE_INCREASE
                Case "physical damage decrease"
                    iValueIndex = SM_LOTM_PHYSICAL_DAMAGE_DECREASE
                Case "health regen"
                    iValueIndex = SM_LOTM_HEALTH_REGEN
                Case "power regen"
                    iValueIndex = SM_LOTM_POWER_REGEN
                Case "endurance regen"
                    iValueIndex = SM_LOTM_ENDURANCE_REGEN
                Case "siege speed increase"
                    iValueIndex = SM_LOTM_SIEGE_SPEED_INCREASE
                Case "water breathing"
                    iValueIndex = SM_LOTM_WATERBREATHING
            End Select
        Case "toa bonus"
            Select Case LCase$(GemEffect)
                Case "armor factor"
                    Utility = 1#
                    iValueIndex = SM_AFBONUS
                Case "% power pool"
                    Utility = 2
                    iValueIndex = SM_PERCPOWER
                Case "archery damage"
                    Utility = 5
                    iValueIndex = SM_ARCHERYDAMAGE
                Case "archery range"
                    Utility = 5
                    iValueIndex = SM_ARCHERYRANGE
                Case "archery speed"
                    Utility = 5
                    iValueIndex = SM_ARCHERYSPEED
                Case "fatigue"
                    Utility = 2
                    iValueIndex = SM_FATIGUE
                Case "healing effectiveness"
                    Utility = 2
                    iValueIndex = SM_HEALINGBONUS
                Case "melee combat speed"
                    Utility = 5
                    iValueIndex = SM_MELEESPEED
                Case "melee damage"
                    Utility = 5
                    iValueIndex = SM_MELEEDAMAGE
                Case "melee style damage"
                    Utility = 5
                    iValueIndex = SM_MELEESTYLE
                Case "spell damage"
                    Utility = 5
                    iValueIndex = SM_SPELLDAMAGE
                Case "spell duration"
                    Utility = 2
                    iValueIndex = SM_SPELLDURATION
                Case "spell pierce"
                    Utility = 5
                    iValueIndex = SM_SPELLPIERCE
                Case "spell haste"
                    Utility = 5
                    iValueIndex = SM_SPELLSPEED
                Case "spell range"
                    Utility = 5
                    iValueIndex = SM_SPELLRANGE
                Case "stat buff effectiveness"
                    Utility = 2
                    iValueIndex = SM_BUFFBONUS
                Case "stat debuff effectiveness"
                    Utility = 2
                    iValueIndex = SM_DEBUFFBONUS
                Case "unique bonus"
                    Utility = 0
                    iValueIndex = SM_TOAUNIQUE
            End Select
        Case "pve bonus"
            Select Case LCase$(GemEffect)
                Case "arcane siphon" 'arrow recovery"
                    Utility = 2
                    iValueIndex = SM_ARROWRECOVERY
                Case "bladeturn reinforcement"
                    Utility = 5
                    iValueIndex = SM_BLADETURN
                Case "block"
                    Utility = 5
                    iValueIndex = SM_BLOCKBONUS
                Case "concentration"
                    Utility = 5
                    iValueIndex = SM_CONCENTRATION
                Case "damage reduction"
                    Utility = 5
                    iValueIndex = SM_DAMAGEREDUCTION
                Case "death experience loss reduction"
                    Utility = -2
                    iValueIndex = SM_EXPLOSSREDUCTION
                Case "defensive"
                    Utility = 5
                    iValueIndex = SM_DEFENSIVEBONUS
                Case "evade"
                    Utility = 5
                    iValueIndex = SM_EVADEBONUS
                Case "negative effect duration"
                    Utility = 5
                    iValueIndex = SM_NEGEFFECTDURATION
                Case "parry"
                    Utility = 5
                    iValueIndex = SM_PARRYBONUS
                Case "piece ablative"
                    Utility = 5
                    iValueIndex = SM_PIECEABLATIVE
                Case "reactionary style damage"
                    Utility = 5
                    iValueIndex = SM_REACTIONARYBONUS
                Case "spell power cost reduction"
                    Utility = 5
                    iValueIndex = SM_SPELLCOSTBONUS
                Case "style cost reduction"
                    Utility = 5
                    iValueIndex = SM_STYLECOSTBONUS
                Case "to-hit"
                    Utility = 5
                    iValueIndex = SM_TOHITBONUS
                Case "unique bonus"
                    Utility = 0
                    iValueIndex = SM_PVEUNIQUE
            End Select
    End Select

    TranslateEffectToMatrix = iValueIndex
    
End Function

Public Sub AssignEffect(ItemLocation As Integer, GemType As ComboBox, GemEffect As ComboBox, _
                            GemAmount As Control, ByRef tToon As TOON_TYPE)
'working as intended :: 9/10/06
    '# STAT_MATRIX (I,J)
    '# I=Stat | J=Location
    
    Dim iValueIndex As Long     'stat being affected
    Dim jValueIndex As Long     'location of the effect being applied
    
    Dim Utility_Modifier As Single
    
    Dim Gem_Utility As Single
   
    'translate the spellcraft locations to the toon stat matrix coordinates
    
    If GemAmount.Name = "txt_GemAmountSC5" Then
        jValueIndex = TranslateLocationToMatrix(ItemLocation + 100)
    Else
        jValueIndex = TranslateLocationToMatrix(ItemLocation)
    End If
    
    If GemEffect.Tag <> vbNullString Then
        tToon.STAT_MATRIX(Val(GemEffect.Tag), jValueIndex) = 0
        GemEffect.Tag = vbNullString
    End If
    
    'check to see if unused type is selected and reset field accordingly
    If LCase$(GemType.Text) = vbNullString Then
        tToon.STAT_MATRIX(Val(GemEffect.Tag), jValueIndex) = 0
        GemEffect.Tag = vbNullString
    Else
        iValueIndex = TranslateEffectToMatrix(GemType.Text, GemEffect.Text, Utility_Modifier)
        If Utility_Modifier <> -1 Then  '-1 denotes that the gem was focus
            GemType.Tag = Val(GemAmount.Text) * Utility_Modifier
        Else
            GemType.Tag = 1 'so if it's focus the utility is 1
        End If
        
    End If
    
    'set the effect column Index to the tag so we know wtf is there
    GemEffect.Tag = iValueIndex
    'assign the value to the matrix
    tToon.STAT_MATRIX(iValueIndex, jValueIndex) = Val(GemAmount.Text)
    
End Sub

'Public Function CalcOvercharge(Item_Level As Long, Item_Quality_Index As Long, Total_Imbue As Single, Crafter_Skill As Long, _
'                                GemQ1 As Long, GemQ2 As Long, GemQ3 As Long, GemQ4 As Long) As Single
'working as intended :: 9/10/06
'
'    Dim OC_Val As Long
'    Dim OC_SkillMod As Long
'
'    Dim OC_Amount As Single
'    Dim OC_TEMP As Long
'
'    Dim ItemCapacity As Long
'
'    ItemCapacity = GetImbuePoints(Item_Level, Item_Quality_Index)
'
'    OC_Amount = Trunc(Total_Imbue - ItemCapacity)
'
'    If OC_Amount < 0 Then OC_Amount = 0
'
'    If OC_Amount > 5.5 Then
'        OC_Val = -1
'    Else
'
'        OC_TEMP = OC_Amount
'
'        'chance = Initial modifier + crafter skill mod + gem quality mods + item quality mod
'
'        OC_Val = OCStartPercentages(OC_Amount)
'        OC_Val = OC_Val + ItemQualityOCMODS(Item_Quality_Index)
'
'        If GemQ1 <> 0 Then OC_Val = OC_Val + GemQualityOCMODS(GemQ1 - 94)
'        If GemQ2 <> 0 Then OC_Val = OC_Val + GemQualityOCMODS(GemQ2 - 94)
'        If GemQ3 <> 0 Then OC_Val = OC_Val + GemQualityOCMODS(GemQ3 - 94)
'        If GemQ4 <> 0 Then OC_Val = OC_Val + GemQualityOCMODS(GemQ4 - 94)
'
'        If Crafter_Skill > 50 Then OC_SkillMod = -45
'        If Crafter_Skill > 100 Then OC_SkillMod = -40
'        If Crafter_Skill > 150 Then OC_SkillMod = -35
'        If Crafter_Skill > 200 Then OC_SkillMod = -30
'        If Crafter_Skill > 250 Then OC_SkillMod = -25
'        If Crafter_Skill > 300 Then OC_SkillMod = -20
'        If Crafter_Skill > 350 Then OC_SkillMod = -15
'        If Crafter_Skill > 400 Then OC_SkillMod = -10
'        If Crafter_Skill > 450 Then OC_SkillMod = -5
'        If Crafter_Skill > 500 Then OC_SkillMod = 0
'        If Crafter_Skill > 550 Then OC_SkillMod = 5
'        If Crafter_Skill > 600 Then OC_SkillMod = 10
'        If Crafter_Skill > 650 Then OC_SkillMod = 15
'        If Crafter_Skill > 700 Then OC_SkillMod = 20
'        If Crafter_Skill > 750 Then OC_SkillMod = 25
'        If Crafter_Skill > 800 Then OC_SkillMod = 30
'        If Crafter_Skill > 850 Then OC_SkillMod = 35
'        If Crafter_Skill > 900 Then OC_SkillMod = 40
'        If Crafter_Skill > 950 Then OC_SkillMod = 45
'        If Crafter_Skill > 1000 Then OC_SkillMod = 50
'
'        OC_Val = OC_Val + OC_SkillMod
'    End If
'
'    CalcOvercharge = OC_Val
'
'End Function

Public Sub SetGemCost(Tier As ComboBox, GemNameTag As Label)
'working as intended :: 10/23/06
'passes the tier and quality to GetGemCost, the value returned from there gets assigned
'to the GemName.Tag

    Dim lCost As Long
    
    lCost = GetGemCost(Tier.ListIndex)
    
    GemNameTag.Tag = lCost
    
End Sub

Private Function GetGemCost(Tier As Long) As Long
'working as intended :: 10/23/06
'this function should never be called directly except by SetGemCost

    'GemInfo(2 3 (make or remake), Tier)   'Tier = 0-9 (gem tier Indexes)
    
    GetGemCost = Val(GemInfo(2, Tier))
    
End Function

Public Function ItemLevel(itemType As Long, AFDPS As Single) As Long
'working as intended :: 9/10/06

    'calculate item level
    If itemType = 1 Then
        ItemLevel = (AFDPS / 2)
    ElseIf itemType = 2 Then
        ItemLevel = (AFDPS)
    ElseIf itemType = 3 Then
        ItemLevel = ((AFDPS - 1.2) / 0.3)
    ElseIf itemType = 4 Then
        If AFDPS = 0 Then
            ItemLevel = 2
        Else
            ItemLevel = (AFDPS * 5)
        End If
    ElseIf itemType = 5 Then
        ItemLevel = ((AFDPS + 5) * 5) + 1
    End If
    
End Function

Public Function CalcGemPoints(GemType As String, GemAmount As Long) As Single
'working as intended :: 9/10/06
    Dim mVal As Single
    
    If GemAmount = 0 Then
        mVal = 0#
    Else
        Select Case LCase$(GemType)
            Case "unused"
                mVal = 0#
            Case "focus"
                mVal = 1#
            Case "hits"
                mVal = GemAmount / 4#
            Case "power"
                mVal = (GemAmount - 1) * 2#
            Case "resist"
                mVal = (GemAmount - 1) * 2#
            Case "skill"
                mVal = (GemAmount - 1) * 5#
            Case "stat"
                mVal = (((GemAmount - 1) / 3#) * 2) + 1
        End Select
        
        If mVal = 0 Then mVal = 1#
    End If
    
    CalcGemPoints = mVal
    
End Function

Public Sub SetDollImage(tToon As TOON_TYPE)
'working as intended :: 9/10/06
    If tToon.REALM = REALM_ALBION Then
        Set WS.imgDoll.Picture = LoadResPicture("DOLL_ALB", vbResBitmap)
    ElseIf tToon.REALM = REALM_HIBERNIA Then
        Set WS.imgDoll.Picture = LoadResPicture("DOLL_HIB", vbResBitmap)
    ElseIf tToon.REALM = REALM_MIDGARD Then
        Set WS.imgDoll.Picture = LoadResPicture("DOLL_MID", vbResBitmap)
    End If
    
End Sub

Public Sub LoadProcEffectDP(CTRL_EFFECT As ComboBox)

    With CTRL_EFFECT
        .Clear
        
        .AddItem "Damage Over Time"
        .AddItem "Dex/Qui Debuff"
        .AddItem "Direct Damage (Cold)"
        .AddItem "Direct Damage (Energy)"
        .AddItem "Direct Damage (Fire"
        .AddItem "Direct Damage (Spirit)"
        .AddItem "Lifedrain"
        .AddItem "Self AF Buff"
        .AddItem "Self Acuity Buff"
        .AddItem "Self Damage Add"
        .AddItem "Self Damage Shield"
        .AddItem "Self Melee Haste"
        .AddItem "Self Melee Health Buffer"
        .AddItem "Self Melee and Magic Health Buffer"
        .AddItem "Str/Con Debuff"
        .AddItem "Accuracy Boost"
        .AddItem "Arcane Leadership"
        .AddItem "Arch Magery"
        .AddItem "Arrogance (Invocation)"
        .AddItem "Attack Speed Decrease"
        .AddItem "Aura of Kings"
        .AddItem "Bedazzled"
        .AddItem "Bolt"
        .AddItem "Boon of Kings"
        .AddItem "Cheat Death"
        .AddItem "Create Item"
        .AddItem "Cure disease"
        .AddItem "Cure mesmerize"
        .AddItem "Cure poison"
        .AddItem "Damage Conversion"
        .AddItem "Damage Shield"
        .AddItem "Defensive Proc"
        .AddItem "Direct Damage"
        .AddItem "Disoriented"
        .AddItem "Dispel"
        .AddItem "Efficient Endurance"
        .AddItem "Efficient Healing"
        .AddItem "Endurance Heal"
        .AddItem "Heal"
        .AddItem "Heal All Over Time"
        .AddItem "Heal Group"
        .AddItem "Illusion"
        .AddItem "Improved Stat Enhancement"
        .AddItem "Lore Debuff"
        .AddItem "Melee Absorption Debuff"
        .AddItem "Mesmerization Feedback"
        .AddItem "Nearsight"
        .AddItem "Omni-Lifedrain"
        .AddItem "Pet Command"
        .AddItem "Raise Dead"
        .AddItem "Realm Lore"
        .AddItem "Recovery"
        .AddItem "Replenish Power"
        .AddItem "Replenish Power (Group)"
        .AddItem "Replenish Power Over Time"
        .AddItem "Resistance Decrease"
        .AddItem "Resistance Enhancement"
        .AddItem "Self Absorb Buff"
        .AddItem "Speed Decrease"
        .AddItem "Speed Enhancement"
        .AddItem "Stat Drain"
        .AddItem "Stealth Lore"
        .AddItem "Stun Feedback"
        .AddItem "Style Damage Shield"
        .AddItem "Summon Elemental"
        .AddItem "Tempest"
        .AddItem "Water Breathing"
        .AddItem "Wave of Healing"
        .AddItem "Weight of a Feather"
        .AddItem "Unique Effect"
    End With
    
End Sub

Public Sub LoadLotmBonusDP(CTRL_EFFECT As ComboBox)

    With CTRL_EFFECT
        .Clear
        
        .AddItem "Encumberence Increase"
        .AddItem "Essence Resist"
        .AddItem "Coin Drop Increase"
        .AddItem "Rez Sick Reduction"
        .AddItem "Safe Fall"
        .AddItem "Arcane Siphon"
        .AddItem "Str Cap"
        .AddItem "Str Cap (+Str)"
        .AddItem "Con Cap"
        .AddItem "Con Cap (+Con)"
        .AddItem "Dex Cap"
        .AddItem "Dex Cap (+Dex)"
        .AddItem "Qui Cap"
        .AddItem "Qui Cap (+Qui)"
        .AddItem "Acuity Cap"
        .AddItem "Acuity Cap (+Acu)"
        .AddItem "Crush Cap"
        .AddItem "Crush Cap (+Crush)"
        .AddItem "Slash Cap"
        .AddItem "Slash Cap (+Slash)"
        .AddItem "Thrust Cap"
        .AddItem "Thrust Cap (+Thrust)"
        .AddItem "Body Cap"
        .AddItem "Body Cap (+Body)"
        .AddItem "Cold Cap"
        .AddItem "Cold Cap (+Cold)"
        .AddItem "Energy Cap"
        .AddItem "Energy Cap (+Energy)"
        .AddItem "Heat Cap"
        .AddItem "Heat Cap (+Heat)"
        .AddItem "Matter Cap"
        .AddItem "Matter Cap (+Matter)"
        .AddItem "Spirit Cap"
        .AddItem "Spirit Cap (+Spirit)"
        .AddItem "Realm Point Bonus"
        .AddItem "Parry"
        .AddItem "Evade"
        .AddItem "Block"
        .AddItem "Siege Damage Decrease"
        .AddItem "Spell Level Increase"
        .AddItem "Crowd Control Decrease"
        .AddItem "Damage Increase"
        .AddItem "Physical Damage Decrease"
        .AddItem "Health Regen"
        .AddItem "Power Regen"
        .AddItem "Endurance Regen"
        .AddItem "Siege Speed Increase"
        .AddItem "Water Breathing"
    End With
    
End Sub

Public Sub LoadToaBonusDP(CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    CTRL_EFFECT.Clear
    
    CTRL_EFFECT.AddItem "% Power Pool"
    CTRL_EFFECT.AddItem "Armor Factor"
    
    'CTRL_EFFECT.AddItem "Archery Damage"
    'CTRL_EFFECT.AddItem "Archery Range"
    'CTRL_EFFECT.AddItem "Archery Speed"
    
    CTRL_EFFECT.AddItem "Fatigue"
    CTRL_EFFECT.AddItem "Healing Effectiveness"
    
    CTRL_EFFECT.AddItem "Melee Combat Speed"
    CTRL_EFFECT.AddItem "Melee Damage"
    CTRL_EFFECT.AddItem "Melee Style Damage"
    
    CTRL_EFFECT.AddItem "Spell Damage"
    CTRL_EFFECT.AddItem "Spell Duration"
    CTRL_EFFECT.AddItem "Spell Haste"
    CTRL_EFFECT.AddItem "Spell Pierce"
    CTRL_EFFECT.AddItem "Spell Range"
    
    CTRL_EFFECT.AddItem "Stat Buff Effectiveness"
    CTRL_EFFECT.AddItem "Stat Debuff Effectiveness"
    
    CTRL_EFFECT.AddItem "Unique Bonus"
    
End Sub

Public Sub LoadPveBonusDP(CTRL_EFFECT As ComboBox)

    CTRL_EFFECT.Clear
    
    CTRL_EFFECT.AddItem "Arcane Siphon"
    CTRL_EFFECT.AddItem "Bladeturn Reinforcement"
    CTRL_EFFECT.AddItem "Block"
    CTRL_EFFECT.AddItem "Concentration"
    CTRL_EFFECT.AddItem "Damage Reduction"
    CTRL_EFFECT.AddItem "Death Experience Loss Reduction"
    CTRL_EFFECT.AddItem "Defensive"
    CTRL_EFFECT.AddItem "Evade"
    CTRL_EFFECT.AddItem "Negative Effect Duration"
    CTRL_EFFECT.AddItem "Parry"
    CTRL_EFFECT.AddItem "Piece Ablative"
    CTRL_EFFECT.AddItem "Reactionary Style Damage"
    CTRL_EFFECT.AddItem "Spell Power Cost Reduction"
    CTRL_EFFECT.AddItem "Style Cost Reduction"
    CTRL_EFFECT.AddItem "To-Hit"
    CTRL_EFFECT.AddItem "Unique Bonus"
    
End Sub

Public Sub LoadCapIncreaseDP(CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    CTRL_EFFECT.Clear
    
    CTRL_EFFECT.AddItem "Strength"
    CTRL_EFFECT.AddItem "Constitution"
    CTRL_EFFECT.AddItem "Dexterity"
    CTRL_EFFECT.AddItem "Quickness"
    CTRL_EFFECT.AddItem "Intelligence"
    CTRL_EFFECT.AddItem "Empathy"
    CTRL_EFFECT.AddItem "Piety"
    CTRL_EFFECT.AddItem "Charisma"
    CTRL_EFFECT.AddItem "Acuity"
    CTRL_EFFECT.AddItem "Hits"
    CTRL_EFFECT.AddItem "Power"
    
End Sub
    
Public Sub LoadFocusDP(REALM As Long, CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    For idx = 0 To 12
        If nFocus(REALM, idx) = vbNullString Then Exit For
        CTRL_EFFECT.AddItem nFocus(REALM, idx)
    Next idx
    
End Sub

Public Sub LoadFocusSC(tToon As TOON_TYPE, CTRL_EFFECT As ComboBox, CTRL_Amount As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    With CTRL_EFFECT
    
        .Clear
        CTRL_Amount.Clear
            
        Select Case tToon.REALM
            Case REALM_ALBION  'alb
                Select Case tToon.CLASS
                    Case TCA_CABALIST  'cabalist
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_ALLSPELL)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_BODY)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_MATTER)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_SPIRIT)
                    Case TCA_NECROMANCER  'necromancer
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_ALLSPELL)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_SERVANT)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_SIGHT)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_PAIN)
                    Case TCA_SORCERER 'sorcerer
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_ALLSPELL)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_BODY)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_MATTER)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_MIND)
                    Case TCA_THEURGIST 'theurgist
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_ALLSPELL)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_COLD)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_EARTH)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_WIND)
                    Case TCA_WIZARD 'wizard
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_ALLSPELL)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_COLD)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_EARTH)
                        .AddItem nFocus(REALM_ALBION, SC_NFOCUS_ALB_FIRE)
                    Case Else
                        .AddItem "No Focus"
                End Select
            Case REALM_HIBERNIA  'hib
                Select Case tToon.CLASS
                    Case TCH_ANIMIST  'animist
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ALLSPELL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ARBOREAL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_CREEPING)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_VERD)
                    Case TCH_BAINSHEE  'bainshee
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ALLSPELL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ETHEREAL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_PHANTASMAL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_SPECTRAL)
                    Case TCH_ELDRITCH  'eldritch
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ALLSPELL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_LIGHT)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_MANA)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_VOID)
                    Case TCH_ENCHANTER  'enchanter
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ALLSPELL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ENCHANTMENTS)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_LIGHT)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_MANA)
                    Case TCH_MENTALIST  'mentalist
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ALLSPELL)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_LIGHT)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_MANA)
                        .AddItem nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_MENT)
                    Case Else
                        .AddItem "No Focus"
                End Select
            Case REALM_MIDGARD  'mid
                Select Case tToon.CLASS
                    Case TCM_BONEDANCER  'bonedancer
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_ALLSPELL)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_BONEARMY)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_DARKNESS)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_SUPPRESSION)
                    Case TCM_RUNEMASTER  'runemaster
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_ALLSPELL)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_DARKNESS)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_RUNECARVING)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_SUPPRESSION)
                    Case TCM_SPIRITMASTER  'spiritmaster
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_ALLSPELL)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_DARKNESS)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_SUMMONING)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_SUPPRESSION)
                    Case TCM_WARLOCK 'warlock
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_ALLSPELL)
                        .AddItem nFocus(REALM_MIDGARD, SC_NFOCUS_MID_CURSING)
                    Case Else
                        .AddItem "No Focus"
                End Select
        End Select
        
    End With
    
    For idx = 0 To 9
        CTRL_Amount.AddItem gFocus(idx)
    Next idx
    
End Sub

Public Sub LoadHitsDP(CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    CTRL_EFFECT.Clear
    CTRL_EFFECT.AddItem "Hits"
    CTRL_EFFECT.Text = "Hits"

End Sub

Public Sub LoadHitsSC(CTRL_EFFECT As ComboBox, CTRL_Amount As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    CTRL_Amount.Clear
    
    CTRL_EFFECT.AddItem "Hits"
        
    For idx = 0 To 9
        CTRL_Amount.AddItem gHIT(idx)
    Next idx
    
End Sub

Public Sub LoadPowerDP(CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    CTRL_EFFECT.Clear
    CTRL_EFFECT.AddItem "Power"
    CTRL_EFFECT.Text = "Power"
    
End Sub

Public Sub LoadPowerSC(CTRL_EFFECT As ComboBox, CTRL_Amount As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    CTRL_Amount.Clear
    
    CTRL_EFFECT.AddItem "Power"
    
    For idx = 0 To 9
        CTRL_Amount.AddItem gPower(idx)
    Next idx
    
End Sub

Public Sub LoadResistsDP(CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    
    For idx = 0 To 8
        CTRL_EFFECT.AddItem nResist(idx)
    Next idx
    
End Sub

Public Sub LoadResistsSC(CTRL_EFFECT As ComboBox, CTRL_Amount As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    CTRL_Amount.Clear
    
    For idx = 0 To 8
        CTRL_EFFECT.AddItem nResist(idx)
    Next idx
    
    For idx = 0 To 9
        CTRL_Amount.AddItem gResist(idx)
    Next idx
    
End Sub

Public Sub LoadSkillDP(REALM As Long, CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    
    For idx = 0 To 45
        If nSkill(REALM, idx) = vbNullString Then Exit For
        CTRL_EFFECT.AddItem nSkill(REALM, idx)
    Next idx
    
End Sub

Public Sub LoadSkillSC(tToon As TOON_TYPE, CTRL_EFFECT As ComboBox, CTRL_Amount As ComboBox)
'working as intended :: 11/16/06
    Dim idx As Long
    
    With CTRL_EFFECT
        .Clear
        CTRL_Amount.Clear
    
        Select Case tToon.REALM
            Case REALM_ALBION  'alb
                Select Case tToon.CLASS
                    Case TCA_ARMSMAN   'armsman
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CRUSH)    'crush
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH)   'slash
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST)   'thrust
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_POLEARM)   'polearm
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_TWOHANDED)   'two-handed
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_PARRY)   'parry
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SHIELD)   'shield
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CROSSBOW)    'crossbow
                    Case TCA_CABALIST  'cabalist
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_BODYMAGIC)    'body
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_MATTERMAGIC)   'matter
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SPIRITMAGIC)   'spirit
                    Case TCA_CLERIC 'cleric
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_ENHANCEMENT)   'enhance
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_REJUVENATION)   'rejuve
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SMITE)   'smite
                    Case TCA_FRIAR  'friar
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_STAFF)   'staff
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_PARRY)   'parry
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_ENHANCEMENT)   'enhance
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_REJUVENATION)   'rejuve
                    Case TCA_HERETIC  'heretic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CRUSH)    'crush
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_FLEXIBLE)   'flex
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SHIELD)   'shield
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_ENHANCEMENT)   'enhance
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_REJUVENATION)   'rejuve
                    Case TCA_INFILTRATOR  'infiltrator
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH)   'slash
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST)   'thrust
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CRITICAL)    'critical strike
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_DUALWIELD)   'dual wield
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_ENVENOM)   'envenom
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_STEALTH)   'stealth
                    Case TCA_MERCENARY  'mercenary
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CRUSH)    'crush
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH)   'slash
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST)   'thrust
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_DUALWIELD)   'dual wield
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_PARRY)   'parry
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SHIELD)   'shield
                    Case TCA_MINSTREL  'minstrel
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH)   'slash
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST)   'thrust
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_STEALTH)   'stealth
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_INSTRUMENTS)   'instruments
                    Case TCA_NECROMANCER  'necromancer
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_DEATHSIGHT)   'deathsight
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_DEATHSERVANT)   'death servant
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_PAINWORKING)    'painworking
                    Case TCA_PALADIN  'paladin
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CRUSH)    'crush
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH)   'slash
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST)   'thrust
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_TWOHANDED)   'two handed
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_PARRY)   'parry
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SHIELD)   'shield
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CHANTS)    'chants
                    Case TCA_REAVER 'reaver
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_CRUSH)    'crush
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH)   'slash
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST)   'thrust
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_FLEXIBLE)   'flex
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_PARRY)   'parry
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SHIELD)   'shield
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SOULRENDING)   'soulrending
                    Case TCA_SCOUT 'scout
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH)   'slash
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST)   'thrust
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_LONGBOW)    'long bow
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_SHIELD)   'shield
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_STEALTH)   'stealth
                    Case TCA_SORCERER 'sorcerer
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_BODYMAGIC)    'body
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_MATTERMAGIC)   'matter
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_MINDMAGIC)    'mind
                    Case TCA_THEURGIST 'theurgist
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_EARTHMAGIC)   'earth
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_COLDMAGIC)     'ice
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_WINDMAGIC)   'wind
                    Case TCA_WIZARD 'wizard
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_EARTHMAGIC)    'earth
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_FIREMAGIC)   'fire
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_COLDMAGIC)    'ice
                    Case TCA_MAULER
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE)
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_AURAMANIP)
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_FISTWRAP)
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_MAGNETISM)
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_MAULERSTAFF)
                        .AddItem nSkill(REALM_ALBION, SC_NSKILL_ALB_POWERSTRIKES)
                End Select
            Case REALM_HIBERNIA  'hib
                Select Case tToon.CLASS
                    Case TCH_ANIMIST  'animist
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ARBOREAL)    'arboreal
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CREEPING)     'creeping
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VERD)   'verdant
                    Case TCH_BAINSHEE  'bainshee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ETHEREAL)   'ethereal shriek
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PHANTASMAL)    'phantasmal wail
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SPECTRAL)   'spectral guard
                    Case TCH_BARD  'bard
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES)    'blades
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLUNT)    'blunt
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MUSIC)   'music
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_NURTURE)   'nurture
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_REGROWTH)   'regrowth
                    Case TCH_BLADEMASTER  'blademaster
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES)    'blades
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLUNT)    'blunt
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CELTICDUAL)    'cd
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PIERCING)   'pierce
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PARRY)   'parry
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SHIELD)   'shield
                    Case TCH_CHAMPION  'champion
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES)    'blades
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLUNT)    'blunt
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PIERCING)   'pierce
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_LARGEWEAP)   'large weaponry
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PARRY)   'parry
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SHIELD)   'shield
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VALOR)   'valor
                    Case TCH_DRUID  'druid
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_NATURE)   'nature
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_NURTURE)   'nurture
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_REGROWTH)   'regrowth
                    Case TCH_ELDRITCH  'eldritch
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_LIGHT)   'light
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MANA)   'mana
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VOID)   'void
                    Case TCH_ENCHANTER  'enchanter
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_LIGHT)   'light
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MANA)   'mana
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ENCHANTMENTS)   'enchantments
                    Case TCH_HERO  'hero
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES)    'blades
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLUNT)    'blunt
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CELTICSPEAR)     'celtic spear
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_LARGEWEAP)   'large weaponry
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PARRY)   'parry
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PIERCING)   'piercing
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SHIELD)   'shield
                    Case TCH_MENTALIST  'mentalist
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_LIGHT)   'light
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MANA)   'mana
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MENT)   'mentalism
                    Case TCH_NIGHTSHADE 'nightshade
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES)    'blades
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PIERCING)   'pierce
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CELTICDUAL)    'celtic dual
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CRITICAL)   'crit strike
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ENVENOM)   'envenom
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_STEALTH)   'stealth
                    Case TCH_RANGER 'ranger
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES)    'blades
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PIERCING)   'pierce
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CELTICDUAL)    'cd
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_RECURVE)   'recurve bow - now archery
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_STEALTH)   'stealth
                    Case TCH_VALEWALKER 'valewalker
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ARBOREAL)    'arboreal
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PARRY)   'parry
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SCYTHE)   'scythe
                    Case TCH_VAMPIIR 'vampiir
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_DEMENTIA)   'dementia
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SHADOW)   'shadow mastery
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VAMPIIRIC)   'vampiiric embrace
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PIERCING)   'piercing
                    Case TCH_WARDEN 'warden
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES)    'blades
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLUNT)    'blunt
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PARRY)   'parry
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SHIELD)   'shield
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_NURTURE)   'nurture
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_REGROWTH)   'regrowth
                    Case TCH_MAULER
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC)
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE)
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_AURAMANIP)
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_FISTWRAP)
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MAGNETISM)
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MAULERSTAFF)
                        .AddItem nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_POWERSTRIKES)
                End Select
            Case REALM_MIDGARD  'mid
                Select Case tToon.CLASS
                    Case TCM_BERSERKER  'berserker
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AXE)    'axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_HAMMER)   'hammer
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_LEFTAXE)   'left axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_PARRY)   'parry
                    Case TCM_BONEDANCER   'bonedancer
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_BONEARMY)    'bone
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_DARKNESS)   'dark
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SUPPRESSION)   'supp
                    Case TCM_HEALER   'healer
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AUGMENTATION)    'aug
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_MENDING)   'mending
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_PAC)   'pac
                    Case TCM_HUNTER  'hunter
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_COMPOSITE)   'bow
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SPEAR)   'spear
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_BEASTCRAFT)     'beastcraft
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_STEALTH)   'stealth
                    Case TCM_RUNEMASTER  'runemaster
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_DARKNESS)   'dark
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SUPPRESSION)   'supp
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_RUNECARVING)   'rune
                    Case TCM_SAVAGE  'savage
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AXE)    'axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_HAMMER)   'hammer
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_HANDTOHAND)   'hand to hand
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_PARRY)   'parry
                    Case TCM_SHADOWBLADE  'shadowblade
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AXE)    'axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_LEFTAXE)   'left axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_CRITICAL)   'critical strike
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_ENVENOM)   'envenom
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_STEALTH)   'stealth
                    Case TCM_SHAMAN  'shaman
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AUGMENTATION)    'aug
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_CAVE)    'cave
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_MENDING)   'mend
                    Case TCM_SKALD  'skald
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AXE)    'axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_HAMMER)   'hammer
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_PARRY)   'parry
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_BATTLESONGS)     'battlesongs
                    Case TCM_SPIRITMASTER  'spiritmaster
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_DARKNESS)   'dark
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SUMMONING)   'summ
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SUPPRESSION)    'supp
                    Case TCM_THANE 'thane
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AXE)    'axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_HAMMER)   'hammer
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_PARRY)   'parry
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SHIELD)   'shield
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_STORMCALLING)   'stormcalling
                    Case TCM_VALKYRIE 'Valkyrie
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SPEAR)   'spear
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_PARRY)   'parry
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SHIELD)   'shield
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_ODIN)   'odin's will
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_MENDING)   'mending
                    Case TCM_WARLOCK  'warlock
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_CURSING)   'cursing
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_HEXING)   'hexing
                    Case TCM_WARRIOR 'warrior
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AXE)    'axe
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_HAMMER)   'hammer
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD)   'sword
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_THROWN)   'thrown
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_PARRY)   'parry
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_SHIELD)   'shield
                    Case TCM_MAULER
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC)    'all magic
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE)    'all melee
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_AURAMANIP)
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_FISTWRAP)
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_MAGNETISM)
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_MAULERSTAFF)
                        .AddItem nSkill(REALM_MIDGARD, SC_NSKILL_MID_POWERSTRIKES)
                End Select
        End Select
    End With    'ctrl_effect
    
    For idx = 0 To 9
        CTRL_Amount.AddItem gSkill(idx)
    Next idx
    
End Sub

Public Sub LoadStatDP(CTRL_EFFECT As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    
    CTRL_EFFECT.AddItem "Acuity"
    CTRL_EFFECT.AddItem "Strength"
    CTRL_EFFECT.AddItem "Constitution"
    CTRL_EFFECT.AddItem "Dexterity"
    CTRL_EFFECT.AddItem "Quickness"
    CTRL_EFFECT.AddItem "Intelligence"
    CTRL_EFFECT.AddItem "Empathy"
    CTRL_EFFECT.AddItem "Piety"
    CTRL_EFFECT.AddItem "Charisma"
        
End Sub

Public Sub LoadStatSC(REALM As Long, CTRL_EFFECT As ComboBox, CTRL_Amount As ComboBox)
'working as intended :: 9/10/06
    Dim idx As Long
    
    CTRL_EFFECT.Clear
    CTRL_Amount.Clear
    
    'common to all
    CTRL_EFFECT.AddItem "Strength"
    CTRL_EFFECT.AddItem "Constitution"
    CTRL_EFFECT.AddItem "Dexterity"
    CTRL_EFFECT.AddItem "Quickness"
    
    Select Case REALM
        Case REALM_ALBION
            'alb
            CTRL_EFFECT.AddItem "Intelligence"
            CTRL_EFFECT.AddItem "Piety"
        Case REALM_HIBERNIA
            'hib
            CTRL_EFFECT.AddItem "Intelligence"
            CTRL_EFFECT.AddItem "Empathy"
        Case REALM_MIDGARD
            'mid
            CTRL_EFFECT.AddItem "Piety"
    End Select
    
    CTRL_EFFECT.AddItem "Charisma"
    
    For idx = 0 To 9
        CTRL_Amount.AddItem gStat(idx)
    Next idx
        
End Sub

Private Sub Init_GemTypes()

    tGemType(0) = ""
    tGemType(1) = "Focus"
    tGemType(2) = "Hits"
    tGemType(3) = "Power"
    tGemType(4) = "Resist"
    tGemType(5) = "Skill"
    tGemType(6) = "Stat"
    
End Sub

Private Sub Init_GemInfo()
'[0][X] = container, [1][X] = cut, [2][X] = cost, [3][X] = remake

    GemInfo(0, 0) = "Lo"
    GemInfo(0, 1) = "Um"
    GemInfo(0, 2) = "On"
    GemInfo(0, 3) = "Ee"
    GemInfo(0, 4) = "Pal"
    GemInfo(0, 5) = "Mon"
    GemInfo(0, 6) = "Ros"
    GemInfo(0, 7) = "Zo"
    GemInfo(0, 8) = "Kath"
    GemInfo(0, 9) = "Ra"
    
    GemInfo(1, 0) = "Raw"
    GemInfo(1, 1) = "Uncut"
    GemInfo(1, 2) = "Rough"
    GemInfo(1, 3) = "Flawed"
    GemInfo(1, 4) = "Imperfect"
    GemInfo(1, 5) = "Polished"
    GemInfo(1, 6) = "Faceted"
    GemInfo(1, 7) = "Precious"
    GemInfo(1, 8) = "Flawless"
    GemInfo(1, 9) = "Perfect"
    
    'cost for gem in copper pieces
    GemInfo(2, 0) = "160"
    GemInfo(2, 1) = "920"
    GemInfo(2, 2) = "3900"
    GemInfo(2, 3) = "13900"
    GemInfo(2, 4) = "40100"
    GemInfo(2, 5) = "88980"
    GemInfo(2, 6) = "133000"
    GemInfo(2, 7) = "198920"
    GemInfo(2, 8) = "258240"
    GemInfo(2, 9) = "296860"

    'cost for gem remake in copper pieces
    GemInfo(3, 0) = "120"
    GemInfo(3, 1) = "560"
    GemInfo(3, 2) = "1740"
    GemInfo(3, 3) = "5260"
    GemInfo(3, 4) = "14180"
    GemInfo(3, 5) = "30660"
    GemInfo(3, 6) = "45520"
    GemInfo(3, 7) = "67680"
    GemInfo(3, 8) = "87640"
    GemInfo(3, 9) = "100700"
    
End Sub

Public Function GetGemName(GEM_TYPE As String, GEM_AMOUNT_Index As Long, Gem_Effect As String, REALM As Long) As String
'working as intended :: 9/10/06
    Dim sCut As String
    Dim sPrefix As String
    Dim sSuffix As String
    
    sCut = vbNullString
    sPrefix = vbNullString
    sSuffix = vbNullString
    
    sCut = GemInfo(1, GEM_AMOUNT_Index)
    
    Select Case LCase$(GEM_TYPE)
        Case "resist"
            Select Case LCase$(Gem_Effect)
                Case "crush"
                    sPrefix = "Fiery"
                Case "slash"
                    sPrefix = "Watery"
                Case "thrust"
                    sPrefix = "Airy"
                Case "heat"
                    sPrefix = "Heated"
                Case "cold"
                    sPrefix = "Icy"
                Case "matter"
                    sPrefix = "Earthen"
                Case "body"
                    sPrefix = "Dusty"
                Case "spirit"
                    sPrefix = "Vapor"
                Case "energy"
                    sPrefix = "Light"
            End Select
                sSuffix = "Shielding"
        Case "stat"
            Select Case LCase$(Gem_Effect)
                Case "strength"
                    sPrefix = "Fiery"
                Case "constitution"
                    sPrefix = "Earthen"
                Case "dexterity"
                    sPrefix = "Vapor"
                Case "quickness"
                    sPrefix = "Airy"
                Case "intelligence"
                    sPrefix = "Dusty"
                Case "empathy"
                    sPrefix = "Heated"
                Case "piety"
                    sPrefix = "Watery"
                Case "charisma"
                    sPrefix = "Icy"
            End Select
                sSuffix = "Essence"
        Case "hits"
            sPrefix = "Blood"
            sSuffix = "Essence"
        Case "power"
            sPrefix = "Mystical"
            sSuffix = "Essence"
        Case "focus"
            Select Case REALM
                Case REALM_ALBION  'alb focus lines
                        'All Spell Lines' : 'Brilliant Sigil',
                        'Body Magic' :      'Heat Sigil',
                        'Cold Magic' :      'Ice Sigil',
                        'Death Servant' :   'Ashen Sigil',
                        'Deathsight' :      'Vacuous Sigil',
                        'Earth Magic' :     'Earth Sigil',
                        'Fire Magic' :      'Fire Sigil',
                        'Matter Magic' :    'Dust Sigil',
                        'Mind Magic' :      'Water Sigil',
                        'Painworking' :     'Salt Crusted Sigil',
                        'Spirit Magic' :    'Vapor Sigil',
                        'Wind Magic' :      'Air Sigil',
                    Select Case LCase$(Gem_Effect)
                        Case "all spell lines"
                            sPrefix = "Brilliant"
                        Case "body magic"
                            sPrefix = "Heat"
                        Case "cold magic"
                            sPrefix = "Ice"
                        Case "death servant"
                            sPrefix = "Ashen"
                        Case "deathsight"
                            sPrefix = "Vacuous"
                        Case "earth magic"
                            sPrefix = "Earth"
                        Case "fire magic"
                            sPrefix = "Fire"
                        Case "matter magic"
                            sPrefix = "Dust"
                        Case "mind magic"
                            sPrefix = "Water"
                        Case "painworking"
                            sPrefix = "Salt Crusted"
                        Case "spirit magic"
                            sPrefix = "Vapor"
                        Case "wind magic"
                            sPrefix = "Air"
                        Case Else
                            sPrefix = vbNullString
                    End Select
                        sSuffix = "Sigil"
                Case REALM_HIBERNIA  'hib focus lines
                        'All Spell Lines' : 'Brilliant Spell Stone',
                        'Arboreal Path' :   'Steaming Spell Stone',
                        'Creeping Path' :   'Oozing Spell Stone',
                        'Enchantments' :    'Vapor Spell Stone',
                        'Ethereal Shriek' : 'Ethereal Spell Stone',
                        'Light' :           'Fire Spell Stone',
                        'Mana' :            'Water Spell Stone',
                        'Mentalism' :       'Earth Spell Stone',
                        'Phantasmal Wail':  'Phantasmal Spell Stone',
                        'Spectral Guard' :  'Spectral Spell Stone',
                        'Verdant Path' :    'Mineral Encrusted Spell Stone',
                        'Void' :            'Ice Spell Stone',
                    Select Case LCase$(Gem_Effect)
                        Case "all spell lines"
                            sPrefix = "Brilliant"
                        Case "arboreal path"
                            sPrefix = "Steaming"
                        Case "creeping path"
                            sPrefix = "Oozing"
                        Case "enchantments"
                            sPrefix = "Vapor"
                        Case "ethereal shriek"
                            sPrefix = "Ethereal"
                        Case "light"
                            sPrefix = "Fire"
                        Case "mana"
                            sPrefix = "Water"
                        Case "mentalism"
                            sPrefix = "Earth"
                        Case "phantasmal wail"
                            sPrefix = "Phantasmal"
                        Case "spectral guard"
                            sPrefix = "Spectral"
                        Case "verdant path"
                            sPrefix = "Mineral Encrusted"
                        Case "void"
                            sPrefix = "Ice"
                        Case Else
                            sPrefix = vbNullString
                    End Select
                        sSuffix = "Spell Stone"
                Case REALM_MIDGARD  'mid spell lines
                        'All Spell Lines' : 'Brilliant Rune',
                        'Bone Army' :       'Ashen Rune',
                        'Cursing' :         'Blighted Rune',
                        'Darkness' :        'Ice Rune',
                        'Runecarving' :     'Heat Rune',
                        'Summoning' :       'Vapor Rune',
                        'Suppression' :     'Dust Rune',
                    Select Case LCase$(Gem_Effect)
                        Case "all spell lines"
                            sPrefix = "Brilliant"
                        Case "bone army"
                            sPrefix = "Ashen"
                        Case "cursing"
                            sPrefix = "Blighted"
                        Case "darkness"
                            sPrefix = "Ice"
                        Case "runecarving"
                            sPrefix = "Heat"
                        Case "summoning"
                            sPrefix = "Vapor"
                        Case "suppression"
                            sPrefix = "Dust"
                        Case Else
                            sPrefix = vbNullString
                    End Select
                        sSuffix = "Rune"
            End Select
        Case "skill"
            Select Case REALM
                Case REALM_ALBION  'alb skill sets
                        'All Magic Skills' :        'Finesse Fervor Sigil',
                        'All Melee Weapon Skills' : 'Finesse War Sigil',
                        'Body Magic' :        'Heated Evocation Sigil',
                        'Chants' :            'Earthen Fervor Sigil',
                        'Cold Magic' :        'Icy Evocation Sigil',
                        'Critical Strike' :   'Heated Battle Jewel',
                        'Crossbow' :          'Vapor War Sigil',
                        'Crush' :             'Fiery War Sigil',
                        'Death Servant' :     'Ashen Fervor Sigil',
                        'Deathsight' :        'Vacuous Fervor Sigil',
                        'Dual Wield' :        'Icy War Sigil',
                        'Earth Magic' :       'Earthen Evocation Sigil',
                        'Enhancement' :       'Airy Fervor Sigil',
                        'Envenom' :           'Dusty Battle Jewel',
                        'Flexible' :          'Molten Magma War Sigil',
                        'Fire Magic' :        'Fiery Evocation Sigil',
                        'Instruments' :       'Vapor Fervor Sigil',
                        'Longbow' :           'Airy War Sigil',
                        'Matter Magic' :      'Dusty Evocation Sigil',
                        'Mind Magic' :        'Watery Evocation Sigil',
                        'Painworking' :       'Salt Crusted Fervor Sigil',
                        'Parry' :             'Vapor Battle Jewel',
                        'Polearm' :           'Earthen War Sigil',
                        'Rejuvenation' :      'Watery Fervor Sigil',
                        'Shield' :            'Fiery Battle Jewel',
                        'Slash' :             'Watery War Sigil',
                        'Smite' :             'Fiery Fervor Sigil',
                        'Soulrending' :       'Steaming Fervor Sigil',
                        'Spirit Magic' :      'Vapor Evocation Sigil',
                        'Staff' :             'Earthen Battle Jewel',
                        'Stealth' :           'Airy Battle Jewel',
                        'Thrust' :            'Dusty War Sigil',
                        'Two Handed' :        'Heated War Sigil',
                        'Wind Magic' :        'Air Evocation Sigil',
                    Select Case LCase$(Gem_Effect)
                        Case "all magic skills"
                            sPrefix = "Finesse Fervor Sigil"
                        Case "all melee skills"
                            sPrefix = "Finesse War Sigil"
                        Case "body magic"
                            sPrefix = "Heated Evocation Sigil"
                        Case "chants"
                            sPrefix = "Earthen Fervor Sigil"
                        Case "cold magic"
                            sPrefix = "Icy Evocation Sigil"
                        Case "critical strike"
                            sPrefix = "Heated Battle Jewel"
                        Case "crossbow"
                            sPrefix = " Vapor War Sigil"
                        Case "crush"
                            sPrefix = "Fiery War Sigil"
                        Case "death servant"
                            sPrefix = "Ashen Fervor Sigil"
                        Case "deathsight"
                            sPrefix = "Vacuous Fervor Sigil"
                        Case "dual wield"
                            sPrefix = "Icy War Sigil"
                        Case "earth magic"
                            sPrefix = "Earthen Evocation Sigil"
                        Case "enhancement"
                            sPrefix = "Airy Fervor Sigil"
                        Case "envenom"
                            sPrefix = "Dusty Battle Jewel"
                        Case "flexible"
                            sPrefix = "Molten Magma War Sigil"
                        Case "fire magic"
                            sPrefix = "Fiery Evocation Sigil"
                        Case "instruments"
                            sPrefix = "Vapor Fervor Sigil"
                        Case "archery"
                            sPrefix = "Airy War Sigil"
                        Case "matter magic"
                            sPrefix = "Dusty Evocation Sigil"
                        Case "mind magic"
                            sPrefix = "Watery Evocation Sigil"
                        Case "painworking"
                            sPrefix = "Salt Crusted Fervor Sigil"
                        Case "parry"
                            sPrefix = "Vapor Battle Jewel"
                        Case "polearm"
                            sPrefix = "Earthen War Sigil"
                        Case "rejuvenation"
                            sPrefix = "Watery Fervor Sigil"
                        Case "shield"
                            sPrefix = "Fiery Battle Jewel"
                        Case "slash"
                            sPrefix = "Watery War Sigil"
                        Case "smite"
                            sPrefix = "Fiery Fervor Sigil"
                        Case "soulrending"
                            sPrefix = "Steaming Fervor Sigil"
                        Case "spirit magic"
                            sPrefix = "Vapor Evocation Sigil"
                        Case "staff"
                            sPrefix = "Earthen Battle Jewel"
                        Case "stealth"
                            sPrefix = "Airy Battle Jewel"
                        Case "thrust"
                            sPrefix = "Dusty War Sigil"
                        Case "two handed"
                            sPrefix = "Heated War Sigil"
                        Case "wind magic"
                            sPrefix = "Air Evocation Sigil"
                        Case "aura manipulation"
                            sPrefix = "Radiant Fervor Sigil"
                        Case "fist wraps"
                            sPrefix = "Glacial War Sigil"
                        Case "magnetism"
                            sPrefix = "Magnetic Fervor Sigil"
                        Case "mauler staff"
                            sPrefix = "Cinder War Sigil"
                        Case "power strikes"
                            sPrefix = "Clout Fervor Sigil"
                    End Select
                Case REALM_HIBERNIA  'hib
                    'All Magic Skills' :        'Finesse Nature Spell Stone',
                    'All Melee Weapon Skills' : 'Finesse War Spell Stone',
                    'Arboreal Path' :     'Steaming Nature Spell Stone',
                    'Blades' :            'Watery War Spell Stone',
                    'Blunt' :             'Fiery War Spell Stone',
                    'Celtic Dual' :       'Icy War Spell Stone',
                    'Celtic Spear' :      'Earthen War Spell Stone',
                    'Creeping Path' :     'Oozing Nature Spell Stone',
                    'Critical Strike' :   'Heated Battle Jewel',
                    'Dementia' :          'Aberrant Arcane Spell Stone',
                    'Enchantments' :      'Vapor Arcane Spell Stone',
                    'Envenom' :           'Dusty Battle Jewel',
                    'Ethereal Shriek' :   'Ethereal Arcane Spell Stone',
                    'Large Weaponry' :    'Heated War Spell Stone',
                    'Light' :             'Fiery Arcane Spell Stone',
                    'Mana' :              'Watery Arcane Spell Stone',
                    'Mentalism' :         'Earthen Arcane Spell Stone',
                    'Music' :             'Airy Nature Spell Stone',
                    'Nature' :            'Earthen Nature Spell Stone',
                    'Nurture' :           'Fiery Nature Spell Stone',
                    'Parry' :             'Vapor Battle Jewel',
                    'Phantasmal Wail' :   'Phantasmal Arcane Spell Stone',
                    'Piercing' :          'Dusty War Spell Stone',
                    'Recurve Bow' :       'Airy War Spell Stone',
                    'Regrowth' :          'Watery Nature Spell Stone',
                    'Scythe' :            'Light War Spell Stone',
                    'Shadow Mastery' :    'Shadowy Arcane Spell Stone',
                    'Shield' :            'Fiery Battle Jewel',
                    'Spectral Guard' :    'Spectral Arcane Spell Stone',
                    'Staff' :             'Earthen Battle Jewel',
                    'Stealth' :           'Airy Battle Jewel',
                    'Valor' :             'Airy Arcane Spell Stone',
                    'Vampiiric Embrace' : 'Embracing Arcane Spell Stone',
                    'Verdant Path' :      'Mineral Encrusted Nature Spell Stone',
                    'Void' :              'Icy Arcane Spell Stone',
                    Select Case LCase$(Gem_Effect)
                        Case "all magic skills"
                            sPrefix = "Finesse Nature Spell Stone"
                        Case "all melee skills"
                            sPrefix = "Finesse War Spell Stone"
                        Case "arboreal path"
                            sPrefix = "Steaming Nature Spell Stone"
                        Case "blades"
                            sPrefix = "Watery War Spell Stone"
                        Case "blunt"
                            sPrefix = "Fiery War Spell Stone"
                        Case "celtic dual"
                            sPrefix = "Icy War Spell Stone"
                        Case "celtic spear"
                            sPrefix = "Earthen War Spell Stone"
                        Case "creeping path"
                            sPrefix = "Oozing Nature Spell Stone"
                        Case "critical strike"
                            sPrefix = "Heated Battle Jewel"
                        Case "dementia"
                            sPrefix = "Aberrant Arcane Spell Stone"
                        Case "enchantments"
                            sPrefix = "Vapor Arcane Spell Stone"
                        Case "envenom"
                            sPrefix = "Dusty Battle Jewel"
                        Case "ethereal shriek"
                            sPrefix = "Ethereal Arcane Spell Stone"
                        Case "large weaponry"
                            sPrefix = "Heated War Spell Stone"
                        Case "light"
                            sPrefix = "Fiery Arcane Spell Stone"
                        Case "mana"
                            sPrefix = "Watery Arcane Spell Stone"
                        Case "mentalism"
                            sPrefix = "Earthen Arcane Spell Stone"
                        Case "music"
                            sPrefix = "Airy Nature Spell Stone"
                        Case "nature"
                            sPrefix = "Earthen Nature Spell Stone"
                        Case "nurture"
                            sPrefix = "Fiery Nature Spell Stone"
                        Case "parry"
                            sPrefix = "Vapor Battle Jewel"
                        Case "phantasmal wail"
                            sPrefix = "Phantasmal Arcane Spell Stone"
                        Case "piercing"
                            sPrefix = "Dusty War Spell Stone"
                        Case "archery"
                            sPrefix = "Airy War Spell Stone"
                        Case "regrowth"
                            sPrefix = "Watery Nature Spell Stone"
                        Case "scythe"
                            sPrefix = "Light War Spell Stone"
                        Case "shadow mastery"
                            sPrefix = "Shadowy Arcane Spell Stone"
                        Case "shield"
                            sPrefix = "Fiery Battle Jewel"
                        Case "spectral guard"
                            sPrefix = "Spectral Arcane Spell Stone"
                        Case "stealth"
                            sPrefix = "Airy Battle Jewel"
                        Case "valor"
                            sPrefix = "Airy Arcane Spell Stone"
                        Case "vampiiric embrace"
                            sPrefix = "Embracing Arcane Spell Stone"
                        Case "verdant path"
                            sPrefix = "Mineral Encrusted Nature Spell Stone"
                        Case "void"
                            sPrefix = "Icy Arcane Spell Stone"
                        Case "aura manipulation"
                            sPrefix = "Radiant Arcane Spell Stone"
                        Case "fist wraps"
                            sPrefix = "Glacial War Spell Stone"
                        Case "magnetism"
                            sPrefix = "Magnetic Arcane Spell Stone"
                        Case "mauler staff"
                            sPrefix = "Cinder War Spell Stone"
                        Case "power strikes"
                            sPrefix = "Clout Arcane Spell Stone"
                    End Select
                Case REALM_MIDGARD  'mid
                    'All Magic Skills' :        'Finesse Primal Rune',
                    'All Melee Weapon Skills' : 'Finesse War Rune',
                    'Augmentation' :      'Airy Chaos Rune',
                    'Axe' :               'Earthen War Rune',
                    'Battlesongs' :       'Airy Primal Rune',
                    'Beastcraft' :        'Earthen Primal Rune',
                    'Bone Army' :         'Ashen Primal Rune',
                    'Cave Magic' :        'Fiery Chaos Rune',
                    'Composite Bow' :     'Airy War Rune',
                    'Critical Strike' :   'Heated Battle Jewel',
                    'Cursing' :           'Blighted Primal Rune',
                    'Darkness' :          'Icy Chaos Rune',
                    'Envenom' :           'Dusty Battle Jewel',
                    'Hammer' :            'Fiery War Rune',
                    'Hand To Hand' :      'Lightning Charged War Rune',
                    'Hexing' :            'Unholy Primal Rune',
                    'Left Axe' :          'Icy War Rune',
                    'Mending' :           'Watery Chaos Rune',
                    'Odin\'s Will' :      'Valiant Primal Rune',
                    'Pacification' :      'Earthen Chaos Rune',
                    'Parry' :             'Vapor Battle Jewel',
                    'Runecarving' :       'Heated Chaos Rune',
                    'Shield' :            'Fiery Battle Jewel',
                    'Spear' :             'Heated War Rune',
                    'Staff' :             'Earthen Battle Jewel',
                    'Stealth' :           'Airy Battle Jewel',
                    'Stormcalling' :      'Fiery Primal Rune',
                    'Summoning' :         'Vapor Chaos Rune',
                    'Suppression' :       'Dusty Chaos Rune',
                    'Sword' :             'Watery War Rune',
                    'Thrown Weapons' :    'Vapor War Rune',
                    Select Case LCase$(Gem_Effect)
                        Case "all magic skills"
                            sPrefix = "Finesse Primal Rune"
                        Case "all melee skills"
                            sPrefix = "Finesse War Rune"
                        Case "augmentation"
                            sPrefix = "Airy Chaos Rune"
                        Case "axe"
                            sPrefix = "Earthen War Rune"
                        Case "battlesongs"
                            sPrefix = "Airy Primal Rune"
                        Case "beastcraft"
                            sPrefix = "Earthen Primal Rune"
                        Case "bone army"
                            sPrefix = "Ashen Primal Rune"
                        Case "cave magic"
                            sPrefix = "Fiery Chaos Rune"
                        Case "archery"
                            sPrefix = "Airy War Rune"
                        Case "critical strike"
                            sPrefix = "Heated Battle Jewel"
                        Case "cursing"
                            sPrefix = "Blighted Primal Rune"
                        Case "darkness"
                            sPrefix = "Icy Chaos Rune"
                        Case "envenom"
                            sPrefix = "Dusty Battle Jewel"
                        Case "hammer"
                            sPrefix = "Fiery War Rune"
                        Case "hand to hand"
                            sPrefix = "Lightning Charged War Rune"
                        Case "hexing"
                            sPrefix = "Unholy Primal Rune"
                        Case "left axe"
                            sPrefix = "Icy War Rune"
                        Case "mending"
                            sPrefix = "Watery Chaos Rune"
                        Case "odin's will"
                            sPrefix = "Valiant Primal Rune"
                        Case "pacification"
                            sPrefix = "Earthen Chaos Rune"
                        Case "parry"
                            sPrefix = "Vapor Battle Jewel"
                        Case "runecarving"
                            sPrefix = "Heated Chaos Rune"
                        Case "shield"
                            sPrefix = "Fiery Battle Jewel"
                        Case "spear"
                            sPrefix = "Heated War Rune"
                        Case "stealth"
                            sPrefix = "Airy Battle Jewel"
                        Case "stormcalling"
                            sPrefix = "Fiery Primal Rune"
                        Case "summoning"
                            sPrefix = "Vapor Chaos Rune"
                        Case "suppression"
                            sPrefix = "Dusty Chaos Rune"
                        Case "sword"
                            sPrefix = "Watery War Rune"
                        Case "thrown weapons"
                            sPrefix = "Vapor War Rune"
                        Case "aura manipulation"
                            sPrefix = "Radiant Primal Rune"
                        Case "fist wraps"
                            sPrefix = "Glacial War Rune"
                        Case "magnetism"
                            sPrefix = "Magnetic Primal Rune"
                        Case "mauler staff"
                            sPrefix = "Cinder War Rune"
                        Case "power strikes"
                            sPrefix = "Clout Primal Rune"
                    End Select
            End Select
    End Select
    
    GetGemName = sCut & " " & sPrefix & " " & sSuffix
End Function

Private Sub Init_nFocus()
'a0, h1, m2
            
    'alb:all spell lines, body magic, cold magic, death servant, deathsight, earth magic,
    'fire magic, matter magic, mind magic, painworking, spirit magic, wind magic
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_ALLSPELL) = "ALL Spell Lines"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_BODY) = "Body Magic"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_COLD) = "Cold Magic"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_SERVANT) = "Death Servant"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_SIGHT) = "Deathsight"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_EARTH) = "Earth Magic"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_FIRE) = "Fire Magic"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_MATTER) = "Matter Magic"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_MIND) = "Mind Magic"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_PAIN) = "Painworking"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_SPIRIT) = "Spirit Magic"
    nFocus(REALM_ALBION, SC_NFOCUS_ALB_WIND) = "Wind Magic"
            
    'hib:all spell lines: 50 lvls, arboreal path, creeping path, enchantments, ethereal shriek,
    'light, mana, mentalism, phantasmal wail, spectral guard, verdant path, void
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ALLSPELL) = "ALL Spell Lines"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ARBOREAL) = "Arboreal Path"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_CREEPING) = "Creeping Path"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ENCHANTMENTS) = "Enchantments"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_ETHEREAL) = "Ethereal Shriek"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_LIGHT) = "Light"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_MANA) = "Mana"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_MENT) = "Mentalism"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_PHANTASMAL) = "Phantasmal Wail"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_SPECTRAL) = "Spectral Guard"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_VERD) = "Verdant Path"
    nFocus(REALM_HIBERNIA, SC_NFOCUS_HIB_VOID) = "Void"
        
    'mid:all spell lines, bone army, cursing, darkness, runecarving, summoning, suppression
    nFocus(REALM_MIDGARD, SC_NFOCUS_MID_ALLSPELL) = "ALL Spell Lines"
    nFocus(REALM_MIDGARD, SC_NFOCUS_MID_BONEARMY) = "Bone Army"
    nFocus(REALM_MIDGARD, SC_NFOCUS_MID_CURSING) = "Cursing"
    nFocus(REALM_MIDGARD, SC_NFOCUS_MID_DARKNESS) = "Darkness"
    nFocus(REALM_MIDGARD, SC_NFOCUS_MID_RUNECARVING) = "Runecarving"
    nFocus(REALM_MIDGARD, SC_NFOCUS_MID_SUMMONING) = "Summoning"
    nFocus(REALM_MIDGARD, SC_NFOCUS_MID_SUPPRESSION) = "Suppression"

End Sub

Private Sub Init_nSkill()
'a0 , h1, m2
    '"all magic", "all melee", "all dual wield", "all archery", body magic, chants, cold magic, critical strike,
    'crossbow, crush, death servant, deathsight, dual wield, earth magic, enhancement, envenom, fire magic, flexible,
    'instruments, longbow, matter magic, mind magic, painworking, parry, polearm, rejuvenation, shield, slash, smite,
    'soulrending, spirit magic, staff, stealth, thrust, two handed, wind magic
        
    nSkill(REALM_ALBION, SC_NSKILL_ALL_ARCHERY) = "ALL Archery Skills"
    nSkill(REALM_ALBION, SC_NSKILL_ALL_DUALWIELD) = "ALL Dual Wield Skills"
 
    nSkill(REALM_ALBION, SC_NSKILL_ALL_MAGIC) = "ALL Magic Skills"
    nSkill(REALM_ALBION, SC_NSKILL_ALL_MELEE) = "ALL Melee Skills"
    
    nSkill(REALM_ALBION, SC_NSKILL_ALB_BODYMAGIC) = "Body Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_CHANTS) = "Chants"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_COLDMAGIC) = "Cold Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_CRITICAL) = "Critical Strike"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_CROSSBOW) = "Crossbow"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_CRUSH) = "Crush"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_DEATHSERVANT) = "Death Servant"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_DEATHSIGHT) = "Deathsight"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_DUALWIELD) = "Dual Wield"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_EARTHMAGIC) = "Earth Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_ENHANCEMENT) = "Enhancement"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_ENVENOM) = "Envenom"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_FIREMAGIC) = "Fire Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_FLEXIBLE) = "Flexible"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_INSTRUMENTS) = "Instruments"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_LONGBOW) = "Archery"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_MATTERMAGIC) = "Matter Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_MINDMAGIC) = "Mind Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_PAINWORKING) = "Painworking"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_PARRY) = "Parry"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_POLEARM) = "Polearm"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_REJUVENATION) = "Rejuvenation"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_SHIELD) = "Shield"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_SLASH) = "Slash"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_SMITE) = "Smite"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_SOULRENDING) = "Soulrending"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_SPIRITMAGIC) = "Spirit Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_STAFF) = "Staff"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_STEALTH) = "Stealth"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_THRUST) = "Thrust"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_TWOHANDED) = "Two Handed"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_WINDMAGIC) = "Wind Magic"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_AURAMANIP) = "Aura Manipulation"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_FISTWRAP) = "Fist Wraps"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_MAGNETISM) = "Magnetism"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_MAULERSTAFF) = "Mauler Staff"
    nSkill(REALM_ALBION, SC_NSKILL_ALB_POWERSTRIKES) = "Power Strikes"
    

    '"all magic", "all melee", "all dual wield", "all archery", arboreal path, blades, blunt, celtic dual,
    'celtic spear, creeping path, critical strike, dementia, enchantments, envenom, ethereal shriek, large weaponry,
    'light, mana, mentalism, music, nature, nurture, parry, phantasmal wail, piercing, recurve bow, regrowth, scythe,
    'shadow mastery, shield, spectral guard, staff, stealth, valor, vampiiric embrace, verdant path, void
    nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_ARCHERY) = "ALL Archery Skills"
    nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_DUALWIELD) = "ALL Dual Wield Skills"
 
    nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MAGIC) = "ALL Magic Skills"
    nSkill(REALM_HIBERNIA, SC_NSKILL_ALL_MELEE) = "ALL Melee Skills"
        
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ARBOREAL) = "Arboreal Path"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLADES) = "Blades"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_BLUNT) = "Blunt"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CELTICDUAL) = "Celtic Dual"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CELTICSPEAR) = "Celtic Spear"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CREEPING) = "Creeping Path"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_CRITICAL) = "Critical Strike"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_DEMENTIA) = "Dementia"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ENCHANTMENTS) = "Enchantments"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ENVENOM) = "Envenom"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_ETHEREAL) = "Ethereal Shriek"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_LARGEWEAP) = "Large Weaponry"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_LIGHT) = "Light"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MANA) = "Mana"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MENT) = "Mentalism"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MUSIC) = "Music"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_NATURE) = "Nature"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_NURTURE) = "Nurture"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PARRY) = "Parry"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PHANTASMAL) = "Phantasmal Wail"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_PIERCING) = "Piercing"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_RECURVE) = "Archery"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_REGROWTH) = "Regrowth"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SCYTHE) = "Scythe"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SHADOW) = "Shadow Mastery"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SHIELD) = "Shield"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_SPECTRAL) = "Spectral Guard"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_STEALTH) = "Stealth"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VALOR) = "Valor"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VAMPIIRIC) = "Vampiiric Embrace"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VERD) = "Verdant Path"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_VOID) = "Void"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_AURAMANIP) = "Aura Manipulation"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_FISTWRAP) = "Fist Wraps"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MAGNETISM) = "Magnetism"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_MAULERSTAFF) = "Mauler Staff"
    nSkill(REALM_HIBERNIA, SC_NSKILL_HIB_POWERSTRIKES) = "Power Strikes"
    '"all magic", "all melee", augmentation, axe, battlesongs, beastcraft, bone army, cave magic, composite bow,
    'critical strike, cursing, darkness, envenom, hammer, hand to hand, hexing, left axe, mending, odin's will, pacification,
    'parry, runecarving, shield, spear, staff, stealth, stormcalling, summoning, suppression, sword, thrown weapons, witchcraft
    nSkill(REALM_MIDGARD, SC_NSKILL_ALL_ARCHERY) = "ALL Archery Skills"
    nSkill(REALM_MIDGARD, SC_NSKILL_ALL_DUALWIELD) = "ALL Dual Wield Skills"
 
    nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MAGIC) = "ALL Magic Skills"
    nSkill(REALM_MIDGARD, SC_NSKILL_ALL_MELEE) = "ALL Melee Skills"
        
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_AUGMENTATION) = "Augmentation"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_AXE) = "Axe"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_BATTLESONGS) = "Battlesongs"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_BEASTCRAFT) = "Beastcraft"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_BONEARMY) = "Bone Army"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_CAVE) = "Cave Magic"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_COMPOSITE) = "Archery"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_CRITICAL) = "Critical Strike"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_CURSING) = "Cursing"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_DARKNESS) = "Darkness"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_ENVENOM) = "Envenom"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_HAMMER) = "Hammer"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_HANDTOHAND) = "Hand to Hand"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_LEFTAXE) = "Left Axe"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_MENDING) = "Mending"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_ODIN) = "Odin's Will"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_PAC) = "Pacification"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_PARRY) = "Parry"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_RUNECARVING) = "Runecarving"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_SHIELD) = "Shield"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_SPEAR) = "Spear"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_STEALTH) = "Stealth"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_STORMCALLING) = "Stormcalling"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_SUMMONING) = "Summoning"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_SUPPRESSION) = "Suppression"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_SWORD) = "Sword"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_THROWN) = "Thrown Weapons"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_WITCHCRAFT) = "Witchcraft"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_HEXING) = "Hexing"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_AURAMANIP) = "Aura Manipulation"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_FISTWRAP) = "Fist Wraps"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_MAGNETISM) = "Magnetism"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_MAULERSTAFF) = "Mauler Staff"
    nSkill(REALM_MIDGARD, SC_NSKILL_MID_POWERSTRIKES) = "Power Strikes"
  
End Sub

Private Sub Init_nResist()

    nResist(0) = "Crush"
    nResist(1) = "Slash"
    nResist(2) = "Thrust"
    nResist(3) = "Heat"
    nResist(4) = "Cold"
    nResist(5) = "Matter"
    nResist(6) = "Body"
    nResist(7) = "Spirit"
    nResist(8) = "Energy"
    
End Sub

Private Sub Init_nStat()

    nStat(0) = "Strength"
    nStat(1) = "Constitution"
    nStat(2) = "Dexterity"
    nStat(3) = "Quickness"
    nStat(4) = "Intelligence"
    nStat(5) = "Empathy"
    nStat(6) = "Piety"
    nStat(7) = "Charisma"
    
End Sub

Private Sub Init_Resist()
'1,2,3,5,7,9,11,13,15,17
    gResist(0) = 1
    gResist(1) = 2
    gResist(2) = 3
    gResist(3) = 5
    gResist(4) = 7
    gResist(5) = 9
    gResist(6) = 11
    gResist(7) = 13
    gResist(8) = 15
    gResist(9) = 17
        
End Sub

Private Sub Init_Stat()
'working as intended :: 9/10/06
'1,4,7,10,13,16,19,22,25,28
    gStat(0) = 1
    gStat(1) = 4
    gStat(2) = 7
    gStat(3) = 10
    gStat(4) = 13
    gStat(5) = 16
    gStat(6) = 19
    gStat(7) = 22
    gStat(8) = 25
    gStat(9) = 28
    
End Sub

Private Sub Init_Hit()
'working as intended :: 9/10/06
'4,12,20,28,36,44,52,60,68,76
    gHIT(0) = 4
    gHIT(1) = 12
    gHIT(2) = 20
    gHIT(3) = 28
    gHIT(4) = 36
    gHIT(5) = 44
    gHIT(6) = 52
    gHIT(7) = 60
    gHIT(8) = 68
    gHIT(9) = 76
    
End Sub

Private Sub Init_Power()
'working as intended :: 9/10/06
'1,2,3,5,7,9,11,13,15,17
    gPower(0) = 1
    gPower(1) = 2
    gPower(2) = 3
    gPower(3) = 5
    gPower(4) = 7
    gPower(5) = 9
    gPower(6) = 11
    gPower(7) = 13
    gPower(8) = 15
    gPower(9) = 17
    
End Sub

Private Sub Init_Focus()
'working as intended :: 9/10/06
'5,10,15,20,25,30,35,40,45,50
    gFocus(0) = 5
    gFocus(1) = 10
    gFocus(2) = 15
    gFocus(3) = 20
    gFocus(4) = 25
    gFocus(5) = 30
    gFocus(6) = 35
    gFocus(7) = 40
    gFocus(8) = 45
    gFocus(9) = 50
    
End Sub

Private Sub Init_Skill()
'working as intended :: 9/10/06
'1,2,3,4,5,6,7,8 9 and 10 Un-Imbueable
    gSkill(0) = 1
    gSkill(1) = 2
    gSkill(2) = 3
    gSkill(3) = 4
    gSkill(4) = 5
    gSkill(5) = 6
    gSkill(6) = 7
    gSkill(7) = 8
    gSkill(8) = 9
    gSkill(9) = 10
    
End Sub

Private Sub Init_ImbueMultipliers()
'working as intended :: 9/10/06
    
    ImbueMultipliers(0) = 1#    'Stat   1.0
    ImbueMultipliers(1) = 2#    'Resist 2.0
    ImbueMultipliers(2) = 5#    'Skill  5.0
    ImbueMultipliers(3) = 0.25  'Hits   0.25
    ImbueMultipliers(4) = 2#    'Power  2.0
    ImbueMultipliers(5) = 1#    'Focus  1.0
    ImbueMultipliers(6) = 0#    'Unused 0.0
    
End Sub

Private Sub Init_OCStartPercentages()
'working as intended :: 9/10/06

    OCStartPercentages(0) = 0
    OCStartPercentages(1) = -10
    OCStartPercentages(2) = -20
    OCStartPercentages(3) = -30
    OCStartPercentages(4) = -50
    OCStartPercentages(5) = -70
    
End Sub

Private Sub Init_GemQualityOCMODS()
'working as intended :: 9/10/06
    
    GemQualityOCMODS(0) = 0     '94
    GemQualityOCMODS(1) = 0     '95
    GemQualityOCMODS(2) = 1     '96
    GemQualityOCMODS(3) = 3     '97
    GemQualityOCMODS(4) = 5     '98
    GemQualityOCMODS(5) = 8     '99
    GemQualityOCMODS(6) = 11    '100
    
End Sub

Private Sub Init_ItemQualityOCMODS()
'working as intended :: 9/10/06
    
    ItemQualityOCMODS(0) = 0    '94
    ItemQualityOCMODS(1) = 0    '95
    ItemQualityOCMODS(2) = 6    '96
    ItemQualityOCMODS(3) = 8    '97
    ItemQualityOCMODS(4) = 10   '98
    ItemQualityOCMODS(5) = 18   '99
    ItemQualityOCMODS(6) = 26   '100
    
End Sub

Public Function GetImbuePoints(LEVEL As Long, Quality_Index As Long) As Long
'working as intended :: 9/10/06
    Dim lResult As Long
    
    If LEVEL - 1 < 0 Then
        lResult = 0
    ElseIf Quality_Index < 0 Then
        lResult = 0
    ElseIf Quality_Index > 6 Then
        lResult = 0
    Else
        lResult = ImbuePoints(LEVEL - 1, Quality_Index)
    End If
    
    GetImbuePoints = lResult
            
End Function

Private Sub Init_ImbuePoints()
'working as intended :: 9/10/06
'old method
'level 1
'ImbuePoints(0, 0) = 0   '94
'ImbuePoints(0, 1) = 1   '95
'ImbuePoints(0, 2) = 1   '96
'ImbuePoints(0, 3) = 1   '97
'ImbuePoints(0, 4) = 1   '98
'ImbuePoints(0, 5) = 1   '99
'ImbuePoints(0, 6) = 1   '100
'--------------------
'level 2
'ImbuePoints(1, 0) = 1
'ImbuePoints(1, 1) = 1
'ImbuePoints(1, 2) = 1
'ImbuePoints(1, 3) = 1
'ImbuePoints(1, 4) = 1
'ImbuePoints(1, 5) = 2
'ImbuePoints(1, 6) = 2
'--------------------
'level 3
'ImbuePoints(2, 0) = 1
'ImbuePoints(2, 1) = 1
'ImbuePoints(2, 2) = 1
'ImbuePoints(2, 3) = 2
'ImbuePoints(2, 4) = 2
'ImbuePoints(2, 5) = 2
'ImbuePoints(2, 6) = 2
'--------------------
'level 4
'ImbuePoints(3, 0) = 1
'ImbuePoints(3, 1) = 1
'ImbuePoints(3, 2) = 2
'ImbuePoints(3, 3) = 2
'ImbuePoints(3, 4) = 2
'ImbuePoints(3, 5) = 3
'ImbuePoints(3, 6) = 3
'--------------------
'level 5
'ImbuePoints(4, 0) = 1
'ImbuePoints(4, 1) = 2
'ImbuePoints(4, 2) = 2
'ImbuePoints(4, 3) = 2
'ImbuePoints(4, 4) = 3
'ImbuePoints(4, 5) = 3
'ImbuePoints(4, 6) = 4
'--------------------
'level 6
'ImbuePoints(5, 0) = 1
'ImbuePoints(5, 1) = 2
'ImbuePoints(5, 2) = 2
'ImbuePoints(5, 3) = 3
'ImbuePoints(5, 4) = 3
'ImbuePoints(5, 5) = 4
'ImbuePoints(5, 6) = 4
'--------------------
'level 7
'ImbuePoints(6, 0) = 2
'ImbuePoints(6, 1) = 2
'ImbuePoints(6, 2) = 3
'ImbuePoints(6, 3) = 3
'ImbuePoints(6, 4) = 4
'ImbuePoints(6, 5) = 4
'ImbuePoints(6, 6) = 5
'--------------------
'level 8
'ImbuePoints(7, 0) = 2
'ImbuePoints(7, 1) = 3
'ImbuePoints(7, 2) = 3
'ImbuePoints(7, 3) = 4
'ImbuePoints(7, 4) = 4
'ImbuePoints(7, 5) = 5
'ImbuePoints(7, 6) = 5
'--------------------
'level 9
'ImbuePoints(8, 0) = 2
'ImbuePoints(8, 1) = 3
'ImbuePoints(8, 2) = 3
'ImbuePoints(8, 3) = 4
'ImbuePoints(8, 4) = 5
'ImbuePoints(8, 5) = 5
'ImbuePoints(8, 6) = 6
'--------------------
'level 10
'ImbuePoints(9, 0) = 2
'ImbuePoints(9, 1) = 3
'ImbuePoints(9, 2) = 4
'ImbuePoints(9, 3) = 4
'ImbuePoints(9, 4) = 5
'ImbuePoints(9, 5) = 6
'ImbuePoints(9, 6) = 7
'--------------------
'level 11
'ImbuePoints(10, 0) = 2
'ImbuePoints(10, 1) = 3
'ImbuePoints(10, 2) = 4
'ImbuePoints(10, 3) = 5
'ImbuePoints(10, 4) = 6
'ImbuePoints(10, 5) = 6
'ImbuePoints(10, 6) = 7
'--------------------
'level 12
'ImbuePoints(11, 0) = 3
'ImbuePoints(11, 1) = 4
'ImbuePoints(11, 2) = 4
'ImbuePoints(11, 3) = 5
'ImbuePoints(11, 4) = 6
'ImbuePoints(11, 5) = 7
'ImbuePoints(11, 6) = 8
'--------------------
'level 13
'ImbuePoints(12, 0) = 3
'ImbuePoints(12, 1) = 4
'ImbuePoints(12, 2) = 5
'ImbuePoints(12, 3) = 6
'ImbuePoints(12, 4) = 6
'ImbuePoints(12, 5) = 7
'ImbuePoints(12, 6) = 9
'--------------------
'level 14
'ImbuePoints(13, 0) = 3
'ImbuePoints(13, 1) = 4
'ImbuePoints(13, 2) = 5
'ImbuePoints(13, 3) = 6
'ImbuePoints(13, 4) = 7
'ImbuePoints(13, 5) = 8
'ImbuePoints(13, 6) = 9
'--------------------
'level 15
'ImbuePoints(14, 0) = 3
'ImbuePoints(14, 1) = 4
'ImbuePoints(14, 2) = 5
'ImbuePoints(14, 3) = 6
'ImbuePoints(14, 4) = 7
'ImbuePoints(14, 5) = 8
'ImbuePoints(14, 6) = 10
'--------------------
'level 16
'ImbuePoints(15, 0) = 3
'ImbuePoints(15, 1) = 5
'ImbuePoints(15, 2) = 6
'ImbuePoints(15, 3) = 7
'ImbuePoints(15, 4) = 8
'ImbuePoints(15, 5) = 9
'ImbuePoints(15, 6) = 10
'--------------------
'level 17
'ImbuePoints(16, 0) = 4
'ImbuePoints(16, 1) = 5
'ImbuePoints(16, 2) = 6
'ImbuePoints(16, 3) = 7
'ImbuePoints(16, 4) = 8
'ImbuePoints(16, 5) = 10
'ImbuePoints(16, 6) = 11
'--------------------
'level 18
'ImbuePoints(17, 0) = 4
'ImbuePoints(17, 1) = 5
'ImbuePoints(17, 2) = 6
'ImbuePoints(17, 3) = 8
'ImbuePoints(17, 4) = 9
'ImbuePoints(17, 5) = 10
'ImbuePoints(17, 6) = 12
'--------------------
'level 19
'ImbuePoints(18, 0) = 4
'ImbuePoints(18, 1) = 6
'ImbuePoints(18, 2) = 7
'ImbuePoints(18, 3) = 8
'ImbuePoints(18, 4) = 9
'ImbuePoints(18, 5) = 11
'ImbuePoints(18, 6) = 12
'--------------------
'level 20
'ImbuePoints(19, 0) = 4
'ImbuePoints(19, 1) = 6
'ImbuePoints(19, 2) = 7
'ImbuePoints(19, 3) = 8
'ImbuePoints(19, 4) = 10
'ImbuePoints(19, 5) = 11
'ImbuePoints(19, 6) = 13
'--------------------
'level 21
'ImbuePoints(20, 0) = 4
'ImbuePoints(20, 1) = 6
'ImbuePoints(20, 2) = 7
'ImbuePoints(20, 3) = 9
'ImbuePoints(20, 4) = 10
'ImbuePoints(20, 5) = 12
'ImbuePoints(20, 6) = 13
'--------------------
'level 22
'ImbuePoints(21, 0) = 5
'ImbuePoints(21, 1) = 6
'ImbuePoints(21, 2) = 8
'ImbuePoints(21, 3) = 9
'ImbuePoints(21, 4) = 11
'ImbuePoints(21, 5) = 12
'ImbuePoints(21, 6) = 14
'--------------------
'level 23
'ImbuePoints(22, 0) = 5
'ImbuePoints(22, 1) = 7
'ImbuePoints(22, 2) = 8
'ImbuePoints(22, 3) = 10
'ImbuePoints(22, 4) = 11
'ImbuePoints(22, 5) = 13
'ImbuePoints(22, 6) = 15
'--------------------
'level 24
'ImbuePoints(23, 0) = 5
'ImbuePoints(23, 1) = 7
'ImbuePoints(23, 2) = 9
'ImbuePoints(23, 3) = 10
'ImbuePoints(23, 4) = 12
'ImbuePoints(23, 5) = 13
'ImbuePoints(23, 6) = 15
'--------------------
'level 25
'ImbuePoints(24, 0) = 5
'ImbuePoints(24, 1) = 7
'ImbuePoints(24, 2) = 9
'ImbuePoints(24, 3) = 10
'ImbuePoints(24, 4) = 12
'ImbuePoints(24, 5) = 14
'ImbuePoints(24, 6) = 16
'--------------------
'level 26
'ImbuePoints(25, 0) = 5
'ImbuePoints(25, 1) = 8
'ImbuePoints(25, 2) = 9
'ImbuePoints(25, 3) = 11
'ImbuePoints(25, 4) = 12
'ImbuePoints(25, 5) = 14
'ImbuePoints(25, 6) = 16
'--------------------
'level 27
'ImbuePoints(26, 0) = 6
'ImbuePoints(26, 1) = 8
'ImbuePoints(26, 2) = 10
'ImbuePoints(26, 3) = 11
'ImbuePoints(26, 4) = 13
'ImbuePoints(26, 5) = 15
'ImbuePoints(26, 6) = 17
'--------------------
'level 28
'ImbuePoints(27, 0) = 6
'ImbuePoints(27, 1) = 8
'ImbuePoints(27, 2) = 10
'ImbuePoints(27, 3) = 12
'ImbuePoints(27, 4) = 13
'ImbuePoints(27, 5) = 15
'ImbuePoints(27, 6) = 18
'--------------------
'level 29
'ImbuePoints(28, 0) = 6
'ImbuePoints(28, 1) = 8
'ImbuePoints(28, 2) = 10
'ImbuePoints(28, 3) = 12
'ImbuePoints(28, 4) = 14
'ImbuePoints(28, 5) = 16
'ImbuePoints(28, 6) = 18
'--------------------
'level 30
'ImbuePoints(29, 0) = 6
'ImbuePoints(29, 1) = 9
'ImbuePoints(29, 2) = 11
'ImbuePoints(29, 3) = 12
'ImbuePoints(29, 4) = 14
'ImbuePoints(29, 5) = 16
'ImbuePoints(29, 6) = 19
'--------------------
'level 31
'ImbuePoints(30, 0) = 6
'ImbuePoints(30, 1) = 9
'ImbuePoints(30, 2) = 11
'ImbuePoints(30, 3) = 13
'ImbuePoints(30, 4) = 15
'ImbuePoints(30, 5) = 17
'ImbuePoints(30, 6) = 20
'--------------------
'level 32
'ImbuePoints(31, 0) = 7
'ImbuePoints(31, 1) = 9
'ImbuePoints(31, 2) = 11
'ImbuePoints(31, 3) = 13
'ImbuePoints(31, 4) = 15
'ImbuePoints(31, 5) = 17
'ImbuePoints(31, 6) = 20
'--------------------
'level 33
'ImbuePoints(32, 0) = 7
'ImbuePoints(32, 1) = 10
'ImbuePoints(32, 2) = 12
'ImbuePoints(32, 3) = 14
'ImbuePoints(32, 4) = 16
'ImbuePoints(32, 5) = 18
'ImbuePoints(32, 6) = 21
'--------------------
'level 34
'ImbuePoints(33, 0) = 7
'ImbuePoints(33, 1) = 10
'ImbuePoints(33, 2) = 12
'ImbuePoints(33, 3) = 14
'ImbuePoints(33, 4) = 16
'ImbuePoints(33, 5) = 18
'ImbuePoints(33, 6) = 21
'--------------------
'level 35
'ImbuePoints(34, 0) = 7
'ImbuePoints(34, 1) = 10
'ImbuePoints(34, 2) = 12
'ImbuePoints(34, 3) = 14
'ImbuePoints(34, 4) = 17
'ImbuePoints(34, 5) = 19
'ImbuePoints(34, 6) = 22
'--------------------
'level 36
'ImbuePoints(35, 0) = 7
'ImbuePoints(35, 1) = 10
'ImbuePoints(35, 2) = 13
'ImbuePoints(35, 3) = 15
'ImbuePoints(35, 4) = 17
'ImbuePoints(35, 5) = 20
'ImbuePoints(35, 6) = 23
'--------------------
'level 37
'ImbuePoints(36, 0) = 8
'ImbuePoints(36, 1) = 11
'ImbuePoints(36, 2) = 13
'ImbuePoints(36, 3) = 15
'ImbuePoints(36, 4) = 17
'ImbuePoints(36, 5) = 20
'ImbuePoints(36, 6) = 23
'--------------------
'level 38
'ImbuePoints(37, 0) = 8
'ImbuePoints(37, 1) = 11
'ImbuePoints(37, 2) = 13
'ImbuePoints(37, 3) = 16
'ImbuePoints(37, 4) = 18
'ImbuePoints(37, 5) = 21
'ImbuePoints(37, 6) = 24
'--------------------
'level 39
'ImbuePoints(38, 0) = 8
'ImbuePoints(38, 1) = 11
'ImbuePoints(38, 2) = 14
'ImbuePoints(38, 3) = 16
'ImbuePoints(38, 4) = 18
'ImbuePoints(38, 5) = 21
'ImbuePoints(38, 6) = 24
'--------------------
'level 40
'ImbuePoints(39, 0) = 8
'ImbuePoints(39, 1) = 11
'ImbuePoints(39, 2) = 14
'ImbuePoints(39, 3) = 16
'ImbuePoints(39, 4) = 19
'ImbuePoints(39, 5) = 22
'ImbuePoints(39, 6) = 25
'--------------------
'level 41
'ImbuePoints(40, 0) = 8
'ImbuePoints(40, 1) = 12
'ImbuePoints(40, 2) = 14
'ImbuePoints(40, 3) = 17
'ImbuePoints(40, 4) = 19
'ImbuePoints(40, 5) = 22
'ImbuePoints(40, 6) = 26
'--------------------
'level 42
'ImbuePoints(41, 0) = 9
'ImbuePoints(41, 1) = 12
'ImbuePoints(41, 2) = 15
'ImbuePoints(41, 3) = 17
'ImbuePoints(41, 4) = 20
'ImbuePoints(41, 5) = 23
'ImbuePoints(41, 6) = 26
'--------------------
'level 43
'ImbuePoints(42, 0) = 9
'ImbuePoints(42, 1) = 12
'ImbuePoints(42, 2) = 15
'ImbuePoints(42, 3) = 18
'ImbuePoints(42, 4) = 20
'ImbuePoints(42, 5) = 23
'ImbuePoints(42, 6) = 27
'--------------------
'level 44
'ImbuePoints(43, 0) = 9
'ImbuePoints(43, 1) = 13
'ImbuePoints(43, 2) = 15
'ImbuePoints(43, 3) = 18
'ImbuePoints(43, 4) = 21
'ImbuePoints(43, 5) = 24
'ImbuePoints(43, 6) = 27
'--------------------
'level 45
'ImbuePoints(44, 0) = 9
'ImbuePoints(44, 1) = 13
'ImbuePoints(44, 2) = 16
'ImbuePoints(44, 3) = 18
'ImbuePoints(44, 4) = 21
'ImbuePoints(44, 5) = 24
'ImbuePoints(44, 6) = 28
'--------------------
'level 46
'ImbuePoints(45, 0) = 9
'ImbuePoints(45, 1) = 13
'ImbuePoints(45, 2) = 16
'ImbuePoints(45, 3) = 19
'ImbuePoints(45, 4) = 22
'ImbuePoints(45, 5) = 25
'ImbuePoints(45, 6) = 29
'--------------------
'level 47
'ImbuePoints(46, 0) = 10
'ImbuePoints(46, 1) = 13
'ImbuePoints(46, 2) = 16
'ImbuePoints(46, 3) = 19
'ImbuePoints(46, 4) = 22
'ImbuePoints(46, 5) = 25
'ImbuePoints(46, 6) = 29
'--------------------
'level 48
'ImbuePoints(47, 0) = 10
'ImbuePoints(47, 1) = 14
'ImbuePoints(47, 2) = 17
'ImbuePoints(47, 3) = 20
'ImbuePoints(47, 4) = 23
'ImbuePoints(47, 5) = 26
'ImbuePoints(47, 6) = 30
'--------------------
'level 49
'ImbuePoints(48, 0) = 10
'ImbuePoints(48, 1) = 14
'ImbuePoints(48, 2) = 17
'ImbuePoints(48, 3) = 20
'ImbuePoints(48, 4) = 23
'ImbuePoints(48, 5) = 27
'ImbuePoints(48, 6) = 31
'--------------------
'level 50
'ImbuePoints(49, 0) = 10
'ImbuePoints(49, 1) = 14
'ImbuePoints(49, 2) = 17
'ImbuePoints(49, 3) = 20
'ImbuePoints(49, 4) = 23
'ImbuePoints(49, 5) = 27
'ImbuePoints(49, 6) = 31
'--------------------
'level 51
'ImbuePoints(50, 0) = 10
'ImbuePoints(50, 1) = 15
'ImbuePoints(50, 2) = 18
'ImbuePoints(50, 3) = 21
'ImbuePoints(50, 4) = 24
'ImbuePoints(50, 5) = 28
'ImbuePoints(50, 6) = 32
'--------------------

'working as intended 1/5/2007
'new method
'level 1
ImbuePoints(0, 0) = 1   '96
ImbuePoints(0, 1) = 1   '97
ImbuePoints(0, 2) = 1   '98
ImbuePoints(0, 3) = 1   '99
ImbuePoints(0, 4) = 1   '100
'--------------------
'level 2
ImbuePoints(1, 0) = 2
ImbuePoints(1, 1) = 2
ImbuePoints(1, 2) = 2
ImbuePoints(1, 3) = 2
ImbuePoints(1, 4) = 2
'--------------------
'level 3
ImbuePoints(2, 0) = 2
ImbuePoints(2, 1) = 2
ImbuePoints(2, 2) = 2
ImbuePoints(2, 3) = 2
ImbuePoints(2, 4) = 2
'--------------------
'level 4
ImbuePoints(3, 0) = 3
ImbuePoints(3, 1) = 3
ImbuePoints(3, 2) = 3
ImbuePoints(3, 3) = 3
ImbuePoints(3, 4) = 3
'--------------------
'level 5
ImbuePoints(4, 0) = 4
ImbuePoints(4, 1) = 4
ImbuePoints(4, 2) = 4
ImbuePoints(4, 3) = 4
ImbuePoints(4, 4) = 4
'--------------------
'level 6
ImbuePoints(5, 0) = 4
ImbuePoints(5, 1) = 4
ImbuePoints(5, 2) = 4
ImbuePoints(5, 3) = 4
ImbuePoints(5, 4) = 4
'--------------------
'level 7
ImbuePoints(6, 0) = 5
ImbuePoints(6, 1) = 5
ImbuePoints(6, 2) = 5
ImbuePoints(6, 3) = 5
ImbuePoints(6, 4) = 5
'--------------------
'level 8
ImbuePoints(7, 0) = 5
ImbuePoints(7, 1) = 5
ImbuePoints(7, 2) = 5
ImbuePoints(7, 3) = 5
ImbuePoints(7, 4) = 5
'--------------------
'level 9
ImbuePoints(8, 0) = 6
ImbuePoints(8, 1) = 6
ImbuePoints(8, 2) = 6
ImbuePoints(8, 3) = 6
ImbuePoints(8, 4) = 6
'--------------------
'level 10
ImbuePoints(9, 0) = 7
ImbuePoints(9, 1) = 7
ImbuePoints(9, 2) = 7
ImbuePoints(9, 3) = 7
ImbuePoints(9, 4) = 7
'--------------------
'level 11
ImbuePoints(10, 0) = 7
ImbuePoints(10, 1) = 7
ImbuePoints(10, 2) = 7
ImbuePoints(10, 3) = 7
ImbuePoints(10, 4) = 7
'--------------------
'level 12
ImbuePoints(11, 0) = 8
ImbuePoints(11, 1) = 8
ImbuePoints(11, 2) = 8
ImbuePoints(11, 3) = 8
ImbuePoints(11, 4) = 8
'--------------------
'level 13
ImbuePoints(12, 0) = 9
ImbuePoints(12, 1) = 9
ImbuePoints(12, 2) = 9
ImbuePoints(12, 3) = 9
ImbuePoints(12, 4) = 9
'--------------------
'level 14
ImbuePoints(13, 0) = 9
ImbuePoints(13, 1) = 9
ImbuePoints(13, 2) = 9
ImbuePoints(13, 3) = 9
ImbuePoints(13, 4) = 9
'--------------------
'level 15
ImbuePoints(14, 0) = 10
ImbuePoints(14, 1) = 10
ImbuePoints(14, 2) = 10
ImbuePoints(14, 3) = 10
ImbuePoints(14, 4) = 10
'--------------------
'level 16
ImbuePoints(15, 0) = 10
ImbuePoints(15, 1) = 10
ImbuePoints(15, 2) = 10
ImbuePoints(15, 3) = 10
ImbuePoints(15, 4) = 10
'--------------------
'level 17
ImbuePoints(16, 0) = 11
ImbuePoints(16, 1) = 11
ImbuePoints(16, 2) = 11
ImbuePoints(16, 3) = 11
ImbuePoints(16, 4) = 11
'--------------------
'level 18
ImbuePoints(17, 0) = 12
ImbuePoints(17, 1) = 12
ImbuePoints(17, 2) = 12
ImbuePoints(17, 3) = 12
ImbuePoints(17, 4) = 12
'--------------------
'level 19
ImbuePoints(18, 0) = 12
ImbuePoints(18, 1) = 12
ImbuePoints(18, 2) = 12
ImbuePoints(18, 3) = 12
ImbuePoints(18, 4) = 12
'--------------------
'level 20
ImbuePoints(19, 0) = 13
ImbuePoints(19, 1) = 13
ImbuePoints(19, 2) = 13
ImbuePoints(19, 3) = 13
ImbuePoints(19, 4) = 13
'--------------------
'level 21
ImbuePoints(20, 0) = 13
ImbuePoints(20, 1) = 13
ImbuePoints(20, 2) = 13
ImbuePoints(20, 3) = 13
ImbuePoints(20, 4) = 13
'--------------------
'level 22
ImbuePoints(21, 0) = 14
ImbuePoints(21, 1) = 14
ImbuePoints(21, 2) = 14
ImbuePoints(21, 3) = 14
ImbuePoints(21, 4) = 14
'--------------------
'level 23
ImbuePoints(22, 0) = 15
ImbuePoints(22, 1) = 15
ImbuePoints(22, 2) = 15
ImbuePoints(22, 3) = 15
ImbuePoints(22, 4) = 15
'--------------------
'level 24
ImbuePoints(23, 0) = 15
ImbuePoints(23, 1) = 15
ImbuePoints(23, 2) = 15
ImbuePoints(23, 3) = 15
ImbuePoints(23, 4) = 15
'--------------------
'level 25
ImbuePoints(24, 0) = 16
ImbuePoints(24, 1) = 16
ImbuePoints(24, 2) = 16
ImbuePoints(24, 3) = 16
ImbuePoints(24, 4) = 16
'--------------------
'level 26
ImbuePoints(25, 0) = 16
ImbuePoints(25, 1) = 16
ImbuePoints(25, 2) = 16
ImbuePoints(25, 3) = 16
ImbuePoints(25, 4) = 16
'--------------------
'level 27
ImbuePoints(26, 0) = 17
ImbuePoints(26, 1) = 17
ImbuePoints(26, 2) = 17
ImbuePoints(26, 3) = 17
ImbuePoints(26, 4) = 17
'--------------------
'level 28
ImbuePoints(27, 0) = 18
ImbuePoints(27, 1) = 18
ImbuePoints(27, 2) = 18
ImbuePoints(27, 3) = 18
ImbuePoints(27, 4) = 18
'--------------------
'level 29
ImbuePoints(28, 0) = 18
ImbuePoints(28, 1) = 18
ImbuePoints(28, 2) = 18
ImbuePoints(28, 3) = 18
ImbuePoints(28, 4) = 18
'--------------------
'level 30
ImbuePoints(29, 0) = 19
ImbuePoints(29, 1) = 19
ImbuePoints(29, 2) = 19
ImbuePoints(29, 3) = 19
ImbuePoints(29, 4) = 19
'--------------------
'level 31
ImbuePoints(30, 0) = 20
ImbuePoints(30, 1) = 20
ImbuePoints(30, 2) = 20
ImbuePoints(30, 3) = 20
ImbuePoints(30, 4) = 20
'--------------------
'level 32
ImbuePoints(31, 0) = 20
ImbuePoints(31, 1) = 20
ImbuePoints(31, 2) = 20
ImbuePoints(31, 3) = 20
ImbuePoints(31, 4) = 20
'--------------------
'level 33
ImbuePoints(32, 0) = 21
ImbuePoints(32, 1) = 21
ImbuePoints(32, 2) = 21
ImbuePoints(32, 3) = 21
ImbuePoints(32, 4) = 21
'--------------------
'level 34
ImbuePoints(33, 0) = 21
ImbuePoints(33, 1) = 21
ImbuePoints(33, 2) = 21
ImbuePoints(33, 3) = 21
ImbuePoints(33, 4) = 21
'--------------------
'level 35
ImbuePoints(34, 0) = 22
ImbuePoints(34, 1) = 22
ImbuePoints(34, 2) = 22
ImbuePoints(34, 3) = 22
ImbuePoints(34, 4) = 22
'--------------------
'level 36
ImbuePoints(35, 0) = 23
ImbuePoints(35, 1) = 23
ImbuePoints(35, 2) = 23
ImbuePoints(35, 3) = 23
ImbuePoints(35, 4) = 23
'--------------------
'level 37
ImbuePoints(36, 0) = 23
ImbuePoints(36, 1) = 23
ImbuePoints(36, 2) = 23
ImbuePoints(36, 3) = 23
ImbuePoints(36, 4) = 23
'--------------------
'level 38
ImbuePoints(37, 0) = 24
ImbuePoints(37, 1) = 24
ImbuePoints(37, 2) = 24
ImbuePoints(37, 3) = 24
ImbuePoints(37, 4) = 24
'--------------------
'level 39
ImbuePoints(38, 0) = 24
ImbuePoints(38, 1) = 24
ImbuePoints(38, 2) = 24
ImbuePoints(38, 3) = 24
ImbuePoints(38, 4) = 24
'--------------------
'level 40
ImbuePoints(39, 0) = 25
ImbuePoints(39, 1) = 25
ImbuePoints(39, 2) = 25
ImbuePoints(39, 3) = 25
ImbuePoints(39, 4) = 25
'--------------------
'level 41
ImbuePoints(40, 0) = 26
ImbuePoints(40, 1) = 26
ImbuePoints(40, 2) = 26
ImbuePoints(40, 3) = 26
ImbuePoints(40, 4) = 26
'--------------------
'level 42
ImbuePoints(41, 0) = 26
ImbuePoints(41, 1) = 26
ImbuePoints(41, 2) = 26
ImbuePoints(41, 3) = 26
ImbuePoints(41, 4) = 26
'--------------------
'level 43
ImbuePoints(42, 0) = 27
ImbuePoints(42, 1) = 27
ImbuePoints(42, 2) = 27
ImbuePoints(42, 3) = 27
ImbuePoints(42, 4) = 27
'--------------------
'level 44
ImbuePoints(43, 0) = 27
ImbuePoints(43, 1) = 27
ImbuePoints(43, 2) = 27
ImbuePoints(43, 3) = 27
ImbuePoints(43, 4) = 27
'--------------------
'level 45
ImbuePoints(44, 0) = 28
ImbuePoints(44, 1) = 28
ImbuePoints(44, 2) = 28
ImbuePoints(44, 3) = 28
ImbuePoints(44, 4) = 28
'--------------------
'level 46
ImbuePoints(45, 0) = 29
ImbuePoints(45, 1) = 29
ImbuePoints(45, 2) = 29
ImbuePoints(45, 3) = 29
ImbuePoints(45, 4) = 29
'--------------------
'level 47
ImbuePoints(46, 0) = 29
ImbuePoints(46, 1) = 29
ImbuePoints(46, 2) = 29
ImbuePoints(46, 3) = 29
ImbuePoints(46, 4) = 29
'--------------------
'level 48
ImbuePoints(47, 0) = 30
ImbuePoints(47, 1) = 30
ImbuePoints(47, 2) = 30
ImbuePoints(47, 3) = 30
ImbuePoints(47, 4) = 30
'--------------------
'level 49
ImbuePoints(48, 0) = 31
ImbuePoints(48, 1) = 31
ImbuePoints(48, 2) = 31
ImbuePoints(48, 3) = 31
ImbuePoints(48, 4) = 31
'--------------------
'level 50
ImbuePoints(49, 0) = 31
ImbuePoints(49, 1) = 31
ImbuePoints(49, 2) = 31
ImbuePoints(49, 3) = 31
ImbuePoints(49, 4) = 31
'--------------------
'level 51
ImbuePoints(50, 0) = 32
ImbuePoints(50, 1) = 32
ImbuePoints(50, 2) = 32
ImbuePoints(50, 3) = 32
ImbuePoints(50, 4) = 32
'--------------------
End Sub

Public Sub Init_ServerCodes()
    'ServerCode(X,Y) x=code, y=name
    
    ServerCodes(0, 0) = "35"
    ServerCodes(1, 0) = "Akatsuki"
    
    ServerCodes(0, 1) = "16"
    ServerCodes(1, 1) = "Bedevere"
    
    ServerCodes(0, 2) = "19"
    ServerCodes(1, 2) = "Bors"
    
    ServerCodes(0, 3) = "34"
    ServerCodes(1, 3) = "Ector"
    
    ServerCodes(0, 4) = "23"
    ServerCodes(1, 4) = "Gaheris"
    
    ServerCodes(0, 5) = "10"
    ServerCodes(1, 5) = "Galahad"
    
    ServerCodes(0, 6) = "33"
    ServerCodes(1, 6) = "Gareth"
    
    ServerCodes(0, 7) = "18"
    ServerCodes(1, 7) = "Gawaine"
    
    ServerCodes(0, 8) = "15"
    ServerCodes(1, 8) = "Guinevere"
    
    ServerCodes(0, 9) = "40"
    ServerCodes(1, 9) = "Hector"
    
    ServerCodes(0, 10) = "28"
    ServerCodes(1, 10) = "Igraine"
    
    ServerCodes(0, 11) = "20"
    ServerCodes(1, 11) = "Iseult"
    
    ServerCodes(0, 12) = "26"
    ServerCodes(1, 12) = "Kay"
    
    ServerCodes(0, 13) = "32"
    ServerCodes(1, 13) = "Lamorak"
    
    ServerCodes(0, 14) = "11"
    ServerCodes(1, 14) = "Lancelot"
    
    ServerCodes(0, 15) = "14"
    ServerCodes(1, 15) = "Merlin"
    
    ServerCodes(0, 16) = "31"
    ServerCodes(1, 16) = "Mordred"
    
    ServerCodes(0, 17) = "17"
    ServerCodes(1, 17) = "Morgan Le Fey"
    
    ServerCodes(0, 18) = "22"
    ServerCodes(1, 18) = "Nimue"
    
    ServerCodes(0, 19) = "13"
    ServerCodes(1, 19) = "Palomides"
    
    ServerCodes(0, 20) = "21"
    ServerCodes(1, 20) = "Pellinor"
    
    ServerCodes(0, 21) = "5"
    ServerCodes(1, 21) = "Pendragon"
    
    ServerCodes(0, 22) = "12"
    ServerCodes(1, 22) = "Percival"
    
    ServerCodes(0, 23) = "27"
    ServerCodes(1, 23) = "Tristan"
    
    ServerCodes(0, 24) = "18"
    ServerCodes(1, 24) = "Glastonbury"
    
    ServerCodes(0, 25) = "19"
    ServerCodes(1, 25) = "Salisbury"
    
    ServerCodes(0, 26) = "1"
    ServerCodes(1, 26) = "Excalibur"
    
    ServerCodes(0, 27) = "6"
    ServerCodes(1, 27) = "Prydwen"
    
    ServerCodes(0, 28) = "2"
    ServerCodes(1, 28) = "Broceliande"
    
    ServerCodes(0, 29) = "5"
    ServerCodes(1, 29) = "Ys"
    
    ServerCodes(0, 30) = "13"
    ServerCodes(1, 30) = "Carnac"
    
    ServerCodes(0, 31) = "10"
    ServerCodes(1, 31) = "Orcanie"
    
    ServerCodes(0, 32) = "3"
    ServerCodes(1, 32) = "Avalon"
    
    ServerCodes(0, 33) = "4"
    ServerCodes(1, 33) = "Lyonesse"
    
    ServerCodes(0, 34) = "12"
    ServerCodes(1, 34) = "Dartmoor"
    
    ServerCodes(0, 35) = "8"
    ServerCodes(1, 35) = "Logres"
    
    ServerCodes(0, 36) = "7"
    ServerCodes(1, 36) = "Stonehenge"
    
    ServerCodes(0, 37) = "17"
    ServerCodes(1, 37) = "Cumbria"
    
    ServerCodes(0, 38) = "16"
    ServerCodes(1, 38) = "Deira"
    
    ServerCodes(0, 39) = "11"
    ServerCodes(1, 39) = "Camlann"
    
    ServerCodes(0, 40) = "41"
    ServerCodes(1, 40) = "Ywain 1"
    
    ServerCodes(0, 41) = "49"
    ServerCodes(1, 41) = "Ywain 2"
    
    ServerCodes(0, 42) = "50"
    ServerCodes(1, 42) = "Ywain 3"
    
    ServerCodes(0, 43) = "51"
    ServerCodes(1, 43) = "Ywain 4"
    
    ServerCodes(0, 44) = "52"
    ServerCodes(1, 44) = "Ywain 5"
    
    ServerCodes(0, 45) = "53"
    ServerCodes(1, 45) = "Ywain 6"
    
    ServerCodes(0, 46) = "54"
    ServerCodes(1, 46) = "Ywain 7"
    
    ServerCodes(0, 47) = "55"
    ServerCodes(1, 47) = "Ywain 8"
    
    ServerCodes(0, 48) = "56"
    ServerCodes(1, 48) = "Ywain 9"
    
    ServerCodes(0, 49) = "57"
    ServerCodes(1, 49) = "Ywain 10"
    
End Sub

Public Function GetServerName(lServer As Long) As String

    Dim lCnt As Long
    Dim sServer As String
    
    For lCnt = 0 To 49
        If lServer = Val(ServerCodes(0, lCnt)) Then
            sServer = ServerCodes(1, lCnt)
            Exit For
        End If
    Next lCnt
    
    GetServerName = sServer
    
End Function

Public Function GetServerCode(sServer As String) As Long

    Dim lCnt As Long
    Dim lServer As Long
    
    lServer = -1
    
    For lCnt = 0 To 49
        If sServer = ServerCodes(1, lCnt) Then
            lServer = Val(ServerCodes(0, lCnt))
            Exit For
        End If
    Next lCnt
            
    GetServerCode = lServer
    
End Function

Public Function GetGemCode(lRealm As Long, sGem As String) As Long

'spellcraft gem hotkey code breakdown
'45,13xxxyy,,-1
'
'xxx = gem type code padded with leading zeros for values less than 100
'yy = bonus amount padded with leading zeros for values less than 10 [0-9] valid values
'
'spellcraft gem list icon is 44,13,,xxxx
'ie. Hotkey_79=44,13,,2667
'
'kort 's looks for this code to determine if the character is a spellcrafter

'1,300,000
'base number = 1,300,000
'example blood essence jewel = 80 * 100 = 8000
'28 hits = gem level 4 (4-1) = 3
'target code = 1308003 = 1300000 + 8000 + 3 = 1308003
'now for a 3 digit base gem code
'example Clout Nature Spell Stone = 140 * 100 = 14000
'+5 skill = gem level 5 (5-1) = 4 (zero based array code)
'target code = 1314004 = 1300000 + 14000 + 4 = 1314004
'now for a 1 digit base gem code
'example Airy Essence Jewel = 6 * 100 = 600
'28 quickness = gem level 10 (10-1) = 9
'target code = 1300609 = 1300000 + 600 + 9 = 1300609

    Dim lCnt As Long
    
    Dim lGemLevel As Long
    Dim lGemRootCode As Long
    Dim lGemMasterCode As Long
    
    Dim sGemCut As String
    Dim sGemRoot As String
    
    sGemCut = Trim$(Left$(sGem, InStr(sGem, " ") - 1))
    sGemRoot = Trim$(Right$(sGem, Len(sGem) - Len(sGemCut) - 1))
    
    'inititalize the gem level to -1 for error checking
    lGemLevel = -1
    lGemRootCode = -1
    
    'get the gem level code, it should not be -1
    For lCnt = 0 To 9
        If sGemCut = GemInfo(1, lCnt) Then
            lGemLevel = lCnt
            Exit For
        End If
    Next lCnt
       
    Select Case lRealm
        Case REALM_ALBION
            Select Case sGemRoot
                Case "Fiery Essence"
                    lGemRootCode = 0
                Case "Earthen Essence"
                    lGemRootCode = 2
                Case "Vapor Essence"
                    lGemRootCode = 4
                Case "Airy Essence"
                    lGemRootCode = 6
                Case "Watery Essence"
                    lGemRootCode = 8
                Case "Heated Essence"
                    lGemRootCode = 10
                Case "Dusty Essence"
                    lGemRootCode = 12
                Case "Icy Essence"
                    lGemRootCode = 14
                Case "Earthen Shielding"
                    lGemRootCode = 16
                Case "Icy Shielding"
                    lGemRootCode = 18
                Case "Heated Shielding"
                    lGemRootCode = 20
                Case "Light Shielding"
                    lGemRootCode = 22
                Case "Airy Shielding"
                    lGemRootCode = 24
                Case "Vapor Shielding"
                    lGemRootCode = 26
                Case "Dusty Shielding"
                    lGemRootCode = 28
                Case "Fiery Shielding"
                    lGemRootCode = 30
                Case "Watery Shielding"
                    lGemRootCode = 32
                Case "Vapor Battle Jewel"
                    lGemRootCode = 34
                Case "Fiery Battle Jewel"
                    lGemRootCode = 36
                Case "Earthen Battle Jewel"
                    lGemRootCode = 38
                Case "Airy Battle Jewel"
                    lGemRootCode = 40
                Case "Dusty Battle Jewel"
                    lGemRootCode = 42
                Case "Heated Battle Jewel"
                    lGemRootCode = 44
                Case "Watery War Sigil"
                    lGemRootCode = 46
                Case "Fiery War Sigil"
                    lGemRootCode = 48
                Case "Dusty War Sigil"
                    lGemRootCode = 50
                Case "Heated War Sigil"
                    lGemRootCode = 52
                Case "Earthen War Sigil"
                    lGemRootCode = 54
                Case "Airy War Sigil"
                    lGemRootCode = 56
                Case "Vapor War Sigil"
                    lGemRootCode = 58
                Case "Icy War Sigil"
                    lGemRootCode = 60
                Case "Fiery Fervor Sigil"
                    lGemRootCode = 62
                Case "Airy Fervor Sigil"
                    lGemRootCode = 64
                Case "Watery Fervor Sigil"
                    lGemRootCode = 66
                Case "Earthen Fervor Sigil"
                    lGemRootCode = 68
                Case "Vapor Fervor Sigil"
                    lGemRootCode = 70
                Case "Earthen Evocation Sigil"
                    lGemRootCode = 72
                Case "Icy Evocation Sigil"
                    lGemRootCode = 74
                Case "Fiery Evocation Sigil"
                    lGemRootCode = 76
                Case "Airy Evocation Sigil"
                    lGemRootCode = 78
                Case "Heated Evocation Sigil"
                    lGemRootCode = 80
                Case "Dusty Evocation Sigil"
                    lGemRootCode = 82
                Case "Vapor Evocation Sigil"
                    lGemRootCode = 84
                Case "Watery Evocation Sigil"
                    lGemRootCode = 86
                Case "Blood Essence"
                    lGemRootCode = 88
                Case "Mystical Essence"
                    lGemRootCode = 90
                Case "Earth Sigil"
                    lGemRootCode = 92
                Case "Ice Sigil"
                    lGemRootCode = 94
                Case "Fire Sigil"
                    lGemRootCode = 96
                Case "Air Sigil"
                    lGemRootCode = 98
                Case "Heat Sigil"
                    lGemRootCode = 100
                Case "Dust Sigil"
                    lGemRootCode = 102
                Case "Vapor Sigil"
                    lGemRootCode = 104
                Case "Water Sigil"
                    lGemRootCode = 106
                Case "Molten Magma War Sigil"
                    lGemRootCode = 108
                Case "Vacuous Fervor Sigil"
                    lGemRootCode = 110
                Case "Salt Crusted Fervor Sigil"
                    lGemRootCode = 112
                Case "Ashen Fervor Sigil"
                    lGemRootCode = 114
                Case "Steaming Fervor Sigil"
                    lGemRootCode = 116
                Case "Vacuous Sigil"
                    lGemRootCode = 118
                Case "Salt Crusted Sigil"
                    lGemRootCode = 120
                Case "Ashen Sigil"
                    lGemRootCode = 122
                Case "Brilliant Sigil"
                    lGemRootCode = 124
                Case "Finesse War Sigil"
                    lGemRootCode = 126
                Case "Finesse Fervor Sigil"
                    lGemRootCode = 128
                Case "Glacial War Sigil"
                    lGemRootCode = 130
                Case "Cinder War Sigil"
                    lGemRootCode = 132
                Case "Radiant Fervor Sigil"
                    lGemRootCode = 134
                Case "Magnetic Fervor Sigil"
                    lGemRootCode = 136
                Case "Clout Fervor Sigil"
                    lGemRootCode = 138
            End Select
        Case REALM_HIBERNIA
            Select Case sGemRoot
                Case "Fiery Essence"
                    lGemRootCode = 0
                Case "Earthen Essence"
                    lGemRootCode = 2
                Case "Vapor Essence"
                    lGemRootCode = 4
                Case "Airy Essence"
                    lGemRootCode = 6
                Case "Watery Essence"
                    lGemRootCode = 8
                Case "Heated Essence"
                    lGemRootCode = 10
                Case "Dusty Essence"
                    lGemRootCode = 12
                Case "Icy Essence"
                    lGemRootCode = 14
                Case "Earthen Shielding"
                    lGemRootCode = 16
                Case "Icy Shielding"
                    lGemRootCode = 18
                Case "Heated Shielding"
                    lGemRootCode = 20
                Case "Light Shielding"
                    lGemRootCode = 22
                Case "Airy Shielding"
                    lGemRootCode = 24
                Case "Vapor Shielding"
                    lGemRootCode = 26
                Case "Dusty Shielding"
                    lGemRootCode = 28
                Case "Fiery Shielding"
                    lGemRootCode = 30
                Case "Watery Shielding"
                    lGemRootCode = 32
                Case "Vapor Battle Jewel"
                    lGemRootCode = 34
                Case "Fiery Battle Jewel"
                    lGemRootCode = 36
                Case "Earthen Battle Jewel"
                    lGemRootCode = 38
                Case "Airy Battle Jewel"
                    lGemRootCode = 40
                Case "Dusty Battle Jewel"
                    lGemRootCode = 42
                Case "Heated Battle Jewel"
                    lGemRootCode = 44
                Case "Watery War Spell Stone"
                    lGemRootCode = 46
                Case "Fiery War Spell Stone"
                    lGemRootCode = 48
                Case "Dusty War Spell Stone"
                    lGemRootCode = 50
                Case "Heated War Spell Stone"
                    lGemRootCode = 52
                Case "Earthen War Spell Stone"
                    lGemRootCode = 54
                Case "Icy War Spell Stone"
                    lGemRootCode = 56
                Case "Airy War Spell Stone"
                    lGemRootCode = 58
                Case "Fiery Nature Spell Stone"
                    lGemRootCode = 60
                Case "Watery Nature Spell Stone"
                    lGemRootCode = 62
                Case "Earthen Nature Spell Stone"
                    lGemRootCode = 64
                Case "Airy Nature Spell Stone"
                    lGemRootCode = 66
                Case "Airy Arcane Spell Stone"
                    lGemRootCode = 68
                Case "Fiery Arcane Spell Stone"
                    lGemRootCode = 70
                Case "Watery Arcane Spell Stone"
                    lGemRootCode = 72
                Case "Vapor Arcane Spell Stone"
                    lGemRootCode = 74
                Case "Icy Arcane Spell Stone"
                    lGemRootCode = 76
                Case "Earthen Arcane Spell Stone"
                    lGemRootCode = 78
                Case "Blood Essence"
                    lGemRootCode = 80
                Case "Mystical Essence"
                    lGemRootCode = 82
                Case "Fire Spell Stone"
                    lGemRootCode = 84
                Case "Water Spell Stone"
                    lGemRootCode = 86
                Case "Vapor Spell Stone"
                    lGemRootCode = 88
                Case "Ice Spell Stone"
                    lGemRootCode = 90
                Case "Earth Spell Stone"
                    lGemRootCode = 92
                Case "Light War Spell Stone"
                    lGemRootCode = 94
                Case "Steaming Nature Spell Stone"
                    lGemRootCode = 96
                Case "Oozing Nature Spell Stone"
                    lGemRootCode = 98
                Case "Mineral Encrusted Nature Spell Stone"
                    lGemRootCode = 100
                Case "Steaming Spell Stone"
                    lGemRootCode = 102
                Case "Oozing Spell Stone"
                    lGemRootCode = 104
                Case "Mineral Encrusted Spell Stone"
                    lGemRootCode = 106
                Case "Spectral Spell Stone"
                    lGemRootCode = 108
                Case "Phantasmal Spell Stone"
                    lGemRootCode = 110
                Case "Ethereal Spell Stone"
                    lGemRootCode = 112
                Case "Spectral Arcane Spell Stone"
                    lGemRootCode = 114
                Case "Phantasmal Arcane Spell Stone"
                    lGemRootCode = 116
                Case "Ethereal Arcane Spell Stone"
                    lGemRootCode = 118
                Case "Shadowy Arcane Spell Stone"
                    lGemRootCode = 120
                Case "Embracing Arcane Spell Stone"
                    lGemRootCode = 122
                Case "Aberrant Arcane Spell Stone"
                    lGemRootCode = 124
                Case "Brilliant Spell Stone"
                    lGemRootCode = 126
                Case "Finesse War Spell Stone"
                    lGemRootCode = 128
                Case "Finesse Nature Spell Stone"
                    lGemRootCode = 130
                Case "Glacial War Spell Stone"
                    lGemRootCode = 132
                Case "Cinder War Spell Stone"
                    lGemRootCode = 134
                Case "Radiant Nature Spell Stone"
                    lGemRootCode = 136
                Case "Magnetic Nature Spell Stone"
                    lGemRootCode = 138
                Case "Clout Nature Spell Stone"
                    lGemRootCode = 140
            End Select
        Case REALM_MIDGARD
            Select Case sGemRoot
                Case "Fiery Essence"
                    lGemRootCode = 0
                Case "Earthen Essence"
                    lGemRootCode = 2
                Case "Vapor Essence"
                    lGemRootCode = 4
                Case "Airy Essence"
                    lGemRootCode = 6
                Case "Watery Essence"
                    lGemRootCode = 8
                Case "Heated Essence"
                    lGemRootCode = 10
                Case "Dusty Essence"
                    lGemRootCode = 12
                Case "Icy Essence"
                    lGemRootCode = 14
                Case "Earthen Shielding"
                    lGemRootCode = 16
                Case "Icy Shielding"
                    lGemRootCode = 18
                Case "Heated Shielding"
                    lGemRootCode = 20
                Case "Light Shielding"
                    lGemRootCode = 22
                Case "Airy Shielding"
                    lGemRootCode = 24
                Case "Vapor Shielding"
                    lGemRootCode = 26
                Case "Dusty Shielding"
                    lGemRootCode = 28
                Case "Fiery Shielding"
                    lGemRootCode = 30
                Case "Watery Shielding"
                    lGemRootCode = 32
                Case "Vapor Battle Jewel"
                    lGemRootCode = 34
                Case "Fiery Battle Jewel"
                    lGemRootCode = 36
                Case "Earthen Battle Jewel"
                    lGemRootCode = 38
                Case "Airy Battle Jewel"
                    lGemRootCode = 40
                Case "Dusty Battle Jewel"
                    lGemRootCode = 42
                Case "Heated Battle Jewel"
                    lGemRootCode = 44
                Case "Watery War Rune"
                    lGemRootCode = 46
                Case "Fiery War Rune"
                    lGemRootCode = 48
                Case "Earthen War Rune"
                    lGemRootCode = 50
                Case "Heated War Rune"
                    lGemRootCode = 52
                Case "Airy War Rune"
                    lGemRootCode = 54
                Case "Vapor War Rune"
                    lGemRootCode = 56
                Case "Icy War Rune"
                    lGemRootCode = 58
                Case "Earthen Primal Rune"
                    lGemRootCode = 60
                Case "Airy Primal Rune"
                    lGemRootCode = 62
                Case "Fiery Primal Rune"
                    lGemRootCode = 64
                Case "Icy Chaos Rune"
                    lGemRootCode = 66
                Case "Dusty Chaos Rune"
                    lGemRootCode = 68
                Case "Heated Chaos Rune"
                    lGemRootCode = 70
                Case "Vapor Chaos Rune"
                    lGemRootCode = 72
                Case "Watery Chaos Rune"
                    lGemRootCode = 74
                Case "Airy Chaos Rune"
                    lGemRootCode = 76
                Case "Fiery Chaos Rune"
                    lGemRootCode = 78
                Case "Blood Essence"
                    lGemRootCode = 82
                Case "Mystical Essence"
                    lGemRootCode = 84
                Case "Ice Rune"
                    lGemRootCode = 86
                Case "Dust Rune"
                    lGemRootCode = 88
                Case "Heat Rune"
                    lGemRootCode = 90
                Case "Vapor Rune"
                    lGemRootCode = 92
                Case "Lightning Charged War Rune"
                    lGemRootCode = 94
                Case "Ashen Primal Rune"
                    lGemRootCode = 96
                Case "Ashen Rune"
                    lGemRootCode = 98
                Case "Blighted Rune"
                    lGemRootCode = 100
                Case "Valiant Primal Rune"
                    lGemRootCode = 104
                Case "Blighted Primal Rune"
                    lGemRootCode = 106
                Case "Unholy Primal Rune"
                    lGemRootCode = 108
                Case "Brilliant Rune"
                    lGemRootCode = 110
                Case "Finesse War Rune"
                    lGemRootCode = 112
                Case "Finesse Primal Rune"
                    lGemRootCode = 114
                Case "Glacial War Rune"
                    lGemRootCode = 116
                Case "Cinder War Rune"
                    lGemRootCode = 118
                Case "Radiant Primal Rune"
                    lGemRootCode = 120
                Case "Magnetic Primal Rune"
                    lGemRootCode = 122
                Case "Clout Primal Rune"
                    lGemRootCode = 124
            End Select
    End Select
    
    lGemMasterCode = -1
    If lGemLevel <> -1 And lGemRootCode <> -1 Then
        lGemMasterCode = 1300000 + (lGemRootCode * 100) + lGemLevel
    End If
        
    GetGemCode = lGemMasterCode
    
End Function
