Attribute VB_Name = "RealmAbility"
Option Explicit

Private Realms(2) As String             'Realm Names
Private RealmRanks(119, 2) As Long      'Realm Point Scale
Private RealmTitles(15, 2, 1) As String 'Realm Rank Titles (rank#, realm, gender)

Private Type RA_TYPE

    Name As String
    Type As String
    Description As String
    
    CostLevel_1 As Long
    CostLevel_2 As Long
    CostLevel_3 As Long
    CostLevel_4 As Long
    CostLevel_5 As Long
    CostLevel_6 As Long
    CostLevel_7 As Long
    CostLevel_8 As Long
    CostLevel_9 As Long
    
    EffectLevel_1 As String
    EffectLevel_2 As String
    EffectLevel_3 As String
    EffectLevel_4 As String
    EffectLevel_5 As String
    EffectLevel_6 As String
    EffectLevel_7 As String
    EffectLevel_8 As String
    EffectLevel_9 As String
    
End Type

Public RealmAbilities(200) As RA_TYPE

Public Const RA_PASSIVE As Long = 1000
Public Const RA_ACTIVE As Long = 2000
Public Const RA_CLASS As Long = 3000

'passive ra's
Public Const RA_AUG_ACU As Long = 1000
Public Const RA_AUG_CON As Long = 1001
Public Const RA_AUG_DEX As Long = 1002
Public Const RA_AUG_QUI As Long = 1003
Public Const RA_AUG_STR As Long = 1004
Public Const RA_AVOIDANCE_OF_MAGIC As Long = 1005
Public Const RA_DETERMINATION As Long = 1006
Public Const RA_ETHEREAL_BOND As Long = 1007
Public Const RA_FALCONS_EYE As Long = 1008
Public Const RA_LIFTER As Long = 1009
Public Const RA_MASTERY_OF_BLOCKING As Long = 1010
Public Const RA_MASTERY_OF_FOCUS As Long = 1011
Public Const RA_MASTERY_OF_HEALING As Long = 1012
Public Const RA_MASTERY_OF_MAGERY As Long = 1013
Public Const RA_MASTERY_OF_PAIN As Long = 1014
Public Const RA_MASTERY_OF_PARRYING As Long = 1015
Public Const RA_MASTERY_OF_STEALTH As Long = 1016
Public Const RA_PHYSICAL_DEFENSE As Long = 1017
Public Const RA_TOUGHNESS As Long = 1018
Public Const RA_VEIL_RECOVERY As Long = 1019
Public Const RA_WILD_HEALING As Long = 1020
Public Const RA_WILD_MINION As Long = 1021
Public Const RA_WILD_POWER As Long = 1022

'active ra's
Public Const RA_ADRENALINE_RUSH As Long = 2023
Public Const RA_AMELIORATING_MELODIES As Long = 2024
Public Const RA_ANGER_OF_THE_GODS As Long = 2025
Public Const RA_BARRIER_OF_FORTITUDE As Long = 2026
Public Const RA_BEDAZZLING_AURA As Long = 2027
Public Const RA_CHARGE As Long = 2028
Public Const RA_CONCENTRATION As Long = 2029
Public Const RA_DASHING_DEFENSE As Long = 2030
Public Const RA_DECIMATION_TRAP As Long = 2031
Public Const RA_DIVINE_INTERVENTION As Long = 2032
Public Const RA_DUAL_THREAT As Long = 2033
Public Const RA_FIRST_AID As Long = 2034
Public Const RA_ICHOR_OF_THE_DEEP As Long = 2035
Public Const RA_IGNORE_PAIN As Long = 2036
Public Const RA_JUGGERNAUT As Long = 2037
Public Const RA_MASTERY_OF_CONCENTRATION As Long = 2038
Public Const RA_MYSTIC_CRYSTAL_LORE As Long = 2039
Public Const RA_NEGATIVE_MAELSTROM As Long = 2040
Public Const RA_PERFECT_RECOVERY As Long = 2041
Public Const RA_PURGE As Long = 2042
Public Const RA_RAGING_POWER As Long = 2043
Public Const RA_REFLEX_ATTACK As Long = 2044
Public Const RA_SECOND_WIND As Long = 2045
Public Const RA_SOLDIERS_BARRICADE As Long = 2046
Public Const RA_SPEED_OF_SOUND As Long = 2047
Public Const RA_STATIC_TEMPEST As Long = 2048
Public Const RA_STRIKE_PREDICTION As Long = 2049
Public Const RA_EMPTY_MIND As Long = 2050
Public Const RA_THORNWEED_FIELD As Long = 2051
Public Const RA_VANISH As Long = 2052
Public Const RA_VEHEMENT_RENEWAL As Long = 2053
Public Const RA_VIPER As Long = 2054
Public Const RA_VOLCANIC_PILLAR As Long = 2055
Public Const RA_WRATH_OF_CHAMPIONS As Long = 2056

'class unique ra's
Public Const RA_ALLURE_OF_DEATH As Long = 3057
Public Const RA_ARMS_LENGTH As Long = 3058
Public Const RA_BADGE_OF_VALOR As Long = 3059
Public Const RA_BLADE_BARRIER As Long = 3060
Public Const RA_BLINDING_DUST As Long = 3061
Public Const RA_BLOOD_DRINKING As Long = 3062
Public Const RA_BOILING_CAULDRON As Long = 3063
Public Const RA_CALL_OF_DARKNESS As Long = 3064
Public Const RA_CALMING_NOTES As Long = 3065
Public Const RA_CHAIN_LIGHTNING As Long = 3066
Public Const RA_COMBAT_AWARENESS As Long = 3067
Public Const RA_DESPERATE_BOWMAN As Long = 3068
Public Const RA_DREAMWEAVER As Long = 3069
Public Const RA_ENTWINING_SNAKES As Long = 3070
Public Const RA_VOICE_SKADI As Long = 3071
Public Const RA_FANATICISM As Long = 3072
Public Const RA_FEROCIOUS_WILL As Long = 3073
Public Const RA_FUELED_BY_RAGE As Long = 3074
Public Const RA_PROTECTION_OF_THE_UNDERHILL As Long = 3075
Public Const RA_FUNGAL_UNION As Long = 3076
Public Const RA_FURY_OF_NATURE As Long = 3077
Public Const RA_MARK_OF_PREY As Long = 3078
Public Const RA_MINION_RESCUE As Long = 3079
Public Const RA_NATURES_WOMB As Long = 3080
Public Const RA_OVERWHELM As Long = 3081
Public Const RA_RESOLUTE_MINION As Long = 3082
Public Const RA_RESTORATIVE_MIND As Long = 3083
Public Const RA_RETRIBUTION_OF_THE_FAITHFUL As Long = 3084
Public Const RA_RUNE_OF_UTTER_AGILITY As Long = 3085
Public Const RA_SELECTIVE_BLINDNESS As Long = 3086
Public Const RA_SELFLESS_DEVOTION As Long = 3087
Public Const RA_SHADOW_SHROUD As Long = 3088
Public Const RA_SHIELD_OF_IMMUNITY As Long = 3089
Public Const RA_SHIELD_TRIP As Long = 3090
Public Const RA_SOLDIERS_CITADEL As Long = 3091
Public Const RA_SONIC_BARRIER As Long = 3092
Public Const RA_SOUL_QUENCH As Long = 3093
Public Const RA_SPIRIT_MARTYR As Long = 3094
Public Const RA_SPUTINS_LEGACY As Long = 3095
Public Const RA_TESTUDO As Long = 3096
Public Const RA_VALE_DEFENSE As Long = 3097
Public Const RA_VALHALLAS_BLESSING As Long = 3098
Public Const RA_WALL_OF_FLAME As Long = 3099
Public Const RA_WHIRLING_STAFF As Long = 3100
Public Const RA_GIFT_OF_PERIZOR As Long = 3101

Public Sub InitRealmAbilityArrays(ByRef ProgressBar As Label, ByRef Status As Label)
'working as intended :: 9/10/06
    Status.Caption = "Initializing Titles..."
    DoEvents
    Call Init_Titles
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Status.Caption = "Initializing Realms..."
    DoEvents
    Call Init_Realms
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Status.Caption = "Initializing Ranks..."
    DoEvents
    Call Init_Ranks
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Status.Caption = "Initializing Realm Abilitites..."
    DoEvents
    Call Init_RealmAbilities
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE * 2
    
End Sub

'Public Function RealmPoints(RA As Long) As Long
'old formula for up to rank 10 - no longer works for 11+
'    RealmPoints = (((25 / 3) * RA ^ 3) - ((25 / 2) * RA ^ 2) + ((25 / 6) * RA))
'
'End Function

Public Function GetRealmSkillPoints(RR As Long, RL As Long) As Long
    
    Dim lResult As Long
    
    If RR <> 0 Then lResult = ((RR - 1) * 10) + RL
    
    GetRealmSkillPoints = lResult
        
End Function

Public Function GetRealmTitle(tToon As TOON_TYPE) As String
    
    Dim sStr As String
    
    If tToon.REALM_RANK > 0 Then sStr = RealmTitles(tToon.REALM_RANK - 1, tToon.REALM, tToon.GENDER)
    
    GetRealmTitle = sStr
    
End Function

Public Function GetRealmRank(POINTS As Long, ByRef tToon As TOON_TYPE) As Long
'working as intended :: 9/10/06
    Dim iCtr As Long
    Dim lResult As Long
                                 
    If (POINTS >= 0) Then
        For iCtr = 0 To 119
            If (iCtr = 119) Then
                Exit For
            Else
                If ((RealmRanks(iCtr, 0) < POINTS) And (RealmRanks(iCtr + 1, 0) > POINTS) Or (RealmRanks(iCtr, 0) = POINTS)) Then Exit For
            End If
        Next iCtr
        
        tToon.REALM_RANK = RealmRanks(iCtr, 1)
        tToon.REALM_LEVEL = RealmRanks(iCtr, 2)
        lResult = 0
    Else
        tToon.REALM_RANK = 0
        tToon.REALM_LEVEL = 0
        lResult = -1
    End If
    
    GetRealmRank = lResult
    
End Function

Public Function GetRealmPoints(RANK As Long, LEVEL As Long) As Long
'working as intended :: 9/10/06
    Dim r As Long
    Dim L As Long
    
    If RANK = 1 Then
        r = 0
    Else
        If (RANK > 1) And (LEVEL >= 0) Then
            r = (((RANK - 1) * 10) - 1) 'convert rank into the starting Index of the realm rank position
        End If
    End If
        
    If RANK <> 1 Then
        L = LEVEL
    Else
        L = LEVEL - 1
    End If
    
    If L < 0 Then L = 0
    
    If r + L > 119 Then
        r = 119
        L = 0
    Else
        r = r + L
    End If
    
    GetRealmPoints = RealmRanks(r, 0)
    
End Function

Public Function GetRealm(Id As Long) As String
'working as intended :: 9/10/06
    GetRealm = Realms(Id)
    
End Function

Private Sub Init_Realms()
'working as intended :: 9/10/06
    Realms(0) = "Albion"
    Realms(1) = "Hibernia"
    Realms(2) = "Midgard"
    
End Sub

Private Sub Init_Titles()
'working as intended :: 9/10/06

'rank, realm, gender
    RealmTitles(0, REALM_ALBION, GENDER_MALE) = "Guardian"         'Rank 1 Alb male
    RealmTitles(0, REALM_ALBION, GENDER_FEMALE) = "Guardian"         'Rank 1 Alb female
    RealmTitles(0, REALM_HIBERNIA, GENDER_MALE) = "Savant"           'Rank 1 Hib male
    RealmTitles(0, REALM_HIBERNIA, GENDER_FEMALE) = "Savant"           'Rank 1 Hib female
    RealmTitles(0, REALM_MIDGARD, GENDER_MALE) = "Skiltvakten"      'Rank 1 Mid male
    RealmTitles(0, REALM_MIDGARD, GENDER_FEMALE) = "Skiltvakten"      'Rank 1 Mid female
    
    RealmTitles(1, REALM_ALBION, GENDER_MALE) = "Warder"           'Rank 2 Alb male
    RealmTitles(1, REALM_ALBION, GENDER_FEMALE) = "Warder"           'Rank 2 Alb female
    RealmTitles(1, REALM_HIBERNIA, GENDER_MALE) = "Cosantoir"        'Rank 2 Hib male
    RealmTitles(1, REALM_HIBERNIA, GENDER_FEMALE) = "Cosantoir"        'Rank 2 Hib female
    RealmTitles(1, REALM_MIDGARD, GENDER_MALE) = "Isen Vakten"      'Rank 2 Mid male
    RealmTitles(1, REALM_MIDGARD, GENDER_FEMALE) = "Isen Vakten"      'Rank 2 Mid female
    
    RealmTitles(2, REALM_ALBION, GENDER_MALE) = "Myrmidon"         'Rank 3 Alb
    RealmTitles(2, REALM_ALBION, GENDER_FEMALE) = "Myrmidon"         'Rank 3 Alb
    RealmTitles(2, REALM_HIBERNIA, GENDER_MALE) = "Brehon"           'Rank 3 Hib
    RealmTitles(2, REALM_HIBERNIA, GENDER_FEMALE) = "Brehon"           'Rank 3 Hib
    RealmTitles(2, REALM_MIDGARD, GENDER_MALE) = "Flammen Vakten"   'Rank 3 Mid
    RealmTitles(2, REALM_MIDGARD, GENDER_FEMALE) = "Flammen Vakten"   'Rank 3 Mid
    
    RealmTitles(3, REALM_ALBION, GENDER_MALE) = "Gryphon Knight"   'Rank 4 Alb
    RealmTitles(3, REALM_ALBION, GENDER_FEMALE) = "Gryphon Knight"   'Rank 4 Alb
    RealmTitles(3, REALM_HIBERNIA, GENDER_MALE) = "Grove Protector"  'Rank 4 Hib
    RealmTitles(3, REALM_HIBERNIA, GENDER_FEMALE) = "Grove Protector"  'Rank 4 Hib
    RealmTitles(3, REALM_MIDGARD, GENDER_MALE) = "Elding Vakten"    'Rank 4 Mid
    RealmTitles(3, REALM_MIDGARD, GENDER_FEMALE) = "Elding Vakten"    'Rank 4 Mid

    RealmTitles(4, REALM_ALBION, GENDER_MALE) = "Eagle Knight"     'Rank 5 Alb
    RealmTitles(4, REALM_ALBION, GENDER_FEMALE) = "Eagle Knight"     'Rank 5 Alb
    RealmTitles(4, REALM_HIBERNIA, GENDER_MALE) = "Raven Ardent"     'Rank 5 Hib
    RealmTitles(4, REALM_HIBERNIA, GENDER_FEMALE) = "Raven Ardent"     'Rank 5 Hib
    RealmTitles(4, REALM_MIDGARD, GENDER_MALE) = "Stormur Vakten"   'Rank 5 Mid
    RealmTitles(4, REALM_MIDGARD, GENDER_FEMALE) = "Stormur Vakten"   'Rank 5 Mid

    RealmTitles(5, REALM_ALBION, GENDER_MALE) = "Phoenix Knight"   'Rank 6 Alb
    RealmTitles(5, REALM_ALBION, GENDER_FEMALE) = "Phoenix Knight"   'Rank 6 Alb
    RealmTitles(5, REALM_HIBERNIA, GENDER_MALE) = "Silver Hand"      'Rank 6 Hib
    RealmTitles(5, REALM_HIBERNIA, GENDER_FEMALE) = "Silver Hand"      'Rank 6 Hib
    RealmTitles(5, REALM_MIDGARD, GENDER_MALE) = "Isen Herra"       'Rank 6 Mid
    RealmTitles(5, REALM_MIDGARD, GENDER_FEMALE) = "Isen Herra"       'Rank 6 Mid

    RealmTitles(6, REALM_ALBION, GENDER_MALE) = "Alerion Knight"   'Rank 7 Alb
    RealmTitles(6, REALM_ALBION, GENDER_FEMALE) = "Alerion Knight"   'Rank 7 Alb
    RealmTitles(6, REALM_HIBERNIA, GENDER_MALE) = "Thunderer"        'Rank 7 Hib
    RealmTitles(6, REALM_HIBERNIA, GENDER_FEMALE) = "Thunderer"        'Rank 7 Hib
    RealmTitles(6, REALM_MIDGARD, GENDER_MALE) = "Flammen Herra"    'Rank 7 Mid
    RealmTitles(6, REALM_MIDGARD, GENDER_FEMALE) = "Flammen Herra"    'Rank 7 Mid

    RealmTitles(7, REALM_ALBION, GENDER_MALE) = "Unicorn Knight"   'Rank 8 Alb
    RealmTitles(7, REALM_ALBION, GENDER_FEMALE) = "Unicorn Knight"   'Rank 8 Alb
    RealmTitles(7, REALM_HIBERNIA, GENDER_MALE) = "Gilded Spear"     'Rank 8 Hib
    RealmTitles(7, REALM_HIBERNIA, GENDER_FEMALE) = "Gilded Spear"     'Rank 8 Hib
    RealmTitles(7, REALM_MIDGARD, GENDER_MALE) = "Elding Herra"     'Rank 8 Mid
    RealmTitles(7, REALM_MIDGARD, GENDER_FEMALE) = "Elding Herra"     'Rank 8 Mid

    RealmTitles(8, REALM_ALBION, GENDER_MALE) = "Lion Knight"      'Rank 9 Alb
    RealmTitles(8, REALM_ALBION, GENDER_FEMALE) = "Lion Knight"      'Rank 9 Alb
    RealmTitles(8, REALM_HIBERNIA, GENDER_MALE) = "Tiarna"           'Rank 9 Hib
    RealmTitles(8, REALM_HIBERNIA, GENDER_FEMALE) = "Tiarna"           'Rank 9 Hib
    RealmTitles(8, REALM_MIDGARD, GENDER_MALE) = "Stormur Vakten"   'Rank 9 Mid
    RealmTitles(8, REALM_MIDGARD, GENDER_FEMALE) = "Stormur Vakten"   'Rank 9 Mid

    RealmTitles(9, REALM_ALBION, GENDER_MALE) = "Dragon Knight"    'Rank 10 Alb
    RealmTitles(9, REALM_ALBION, GENDER_FEMALE) = "Dragon Knight"    'Rank 10 Alb
    RealmTitles(9, REALM_HIBERNIA, GENDER_MALE) = "Emerald Rider"    'Rank 10 Hib
    RealmTitles(9, REALM_HIBERNIA, GENDER_FEMALE) = "Emerald Rider"    'Rank 10 Hib
    RealmTitles(9, REALM_MIDGARD, GENDER_MALE) = "Einherjar"        'Rank 10 Mid
    RealmTitles(9, REALM_MIDGARD, GENDER_FEMALE) = "Einherjar"        'Rank 10 Mid
    
    RealmTitles(10, REALM_ALBION, GENDER_MALE) = "Lord"            'Rank 11 male Alb
    RealmTitles(10, REALM_ALBION, GENDER_FEMALE) = "Lady"            'Rank 11 female Alb
    RealmTitles(10, REALM_HIBERNIA, GENDER_MALE) = "Barun"           'Rank 11 male Hib
    RealmTitles(10, REALM_HIBERNIA, GENDER_FEMALE) = "Banbharun"       'Rank 11 female Hib
    RealmTitles(10, REALM_MIDGARD, GENDER_MALE) = "Herra"           'Rank 11 male Mid
    RealmTitles(10, REALM_MIDGARD, GENDER_FEMALE) = "Fru"             'Rank 11 female Mid
    
    RealmTitles(11, REALM_ALBION, GENDER_MALE) = "Baronet"         'Rank 12 male alb
    RealmTitles(11, REALM_ALBION, GENDER_FEMALE) = "Baronetess"      'Rank 12 female Alb
    RealmTitles(11, REALM_HIBERNIA, GENDER_MALE) = "Ard Tiarna"      'Rank 12 male Hib
    RealmTitles(11, REALM_HIBERNIA, GENDER_FEMALE) = "Ard Bantiarna"   'Rank 12 female Hib
    RealmTitles(11, REALM_MIDGARD, GENDER_MALE) = "Hersir"          'Rank 12 male Mid
    RealmTitles(11, REALM_MIDGARD, GENDER_FEMALE) = "Baronsfru"       'Rank 12 female Mid

    RealmTitles(12, REALM_ALBION, GENDER_MALE) = "Baron"           'Rank 13 male Alb
    RealmTitles(12, REALM_ALBION, GENDER_FEMALE) = "Baroness"        'Rank 13 female Alb
    RealmTitles(12, REALM_HIBERNIA, GENDER_MALE) = "Ciann"         'Rank 13 male Hib
    RealmTitles(12, REALM_HIBERNIA, GENDER_FEMALE) = "Cath"      'Rank 13 female Hib
    RealmTitles(12, REALM_MIDGARD, GENDER_MALE) = "Vicomte"           'Rank 13 male Mid
    RealmTitles(12, REALM_MIDGARD, GENDER_FEMALE) = "Vicomtessa"            'Rank 13 female Mid
    
End Sub

Private Sub Init_Ranks()
'working as intended :: 9/10/06
    '--realm rank 1
    RealmRanks(0, 0) = 0                'Points
    RealmRanks(0, 1) = 1                'Rank
    RealmRanks(0, 2) = 1                'Level
    
    RealmRanks(1, 0) = 25
    RealmRanks(1, 1) = 1
    RealmRanks(1, 2) = 2
    
    RealmRanks(2, 0) = 125
    RealmRanks(2, 1) = 1
    RealmRanks(2, 2) = 3
    
    RealmRanks(3, 0) = 350
    RealmRanks(3, 1) = 1
    RealmRanks(3, 2) = 4
    
    RealmRanks(4, 0) = 750
    RealmRanks(4, 1) = 1
    RealmRanks(4, 2) = 5
    
    RealmRanks(5, 0) = 1375
    RealmRanks(5, 1) = 1
    RealmRanks(5, 2) = 6
    
    RealmRanks(6, 0) = 2275
    RealmRanks(6, 1) = 1
    RealmRanks(6, 2) = 7
    
    RealmRanks(7, 0) = 3500
    RealmRanks(7, 1) = 1
    RealmRanks(7, 2) = 8
    
    RealmRanks(8, 0) = 5100
    RealmRanks(8, 1) = 1
    RealmRanks(8, 2) = 9
    '--------------------
    
    '--realm rank 2
    RealmRanks(9, 0) = 7125
    RealmRanks(9, 1) = 2
    RealmRanks(9, 2) = 0
    
    RealmRanks(10, 0) = 9625
    RealmRanks(10, 1) = 2
    RealmRanks(10, 2) = 1
    
    RealmRanks(11, 0) = 12650
    RealmRanks(11, 1) = 2
    RealmRanks(11, 2) = 2
    
    RealmRanks(12, 0) = 16250
    RealmRanks(12, 1) = 2
    RealmRanks(12, 2) = 3
    
    RealmRanks(13, 0) = 20475
    RealmRanks(13, 1) = 2
    RealmRanks(13, 2) = 4
    
    RealmRanks(14, 0) = 25375
    RealmRanks(14, 1) = 2
    RealmRanks(14, 2) = 5
    
    RealmRanks(15, 0) = 31000
    RealmRanks(15, 1) = 2
    RealmRanks(15, 2) = 6
    
    RealmRanks(16, 0) = 37400
    RealmRanks(16, 1) = 2
    RealmRanks(16, 2) = 7
    
    RealmRanks(17, 0) = 44625
    RealmRanks(17, 1) = 2
    RealmRanks(17, 2) = 8
    
    RealmRanks(18, 0) = 52725
    RealmRanks(18, 1) = 2
    RealmRanks(18, 2) = 9
    '--------------------
    
    '--realm rank 3
    RealmRanks(19, 0) = 61750
    RealmRanks(19, 1) = 3
    RealmRanks(19, 2) = 0
    
    RealmRanks(20, 0) = 71750
    RealmRanks(20, 1) = 3
    RealmRanks(20, 2) = 1
    
    RealmRanks(21, 0) = 82775
    RealmRanks(21, 1) = 3
    RealmRanks(21, 2) = 2
    
    RealmRanks(22, 0) = 94875
    RealmRanks(22, 1) = 3
    RealmRanks(22, 2) = 3
    
    RealmRanks(23, 0) = 108100
    RealmRanks(23, 1) = 3
    RealmRanks(23, 2) = 4
    
    RealmRanks(24, 0) = 122500
    RealmRanks(24, 1) = 3
    RealmRanks(24, 2) = 5
    
    RealmRanks(25, 0) = 138125
    RealmRanks(25, 1) = 3
    RealmRanks(25, 2) = 6
    
    RealmRanks(26, 0) = 155025
    RealmRanks(26, 1) = 3
    RealmRanks(26, 2) = 7
    
    RealmRanks(27, 0) = 173250
    RealmRanks(27, 1) = 3
    RealmRanks(27, 2) = 8
    
    RealmRanks(28, 0) = 192850
    RealmRanks(28, 1) = 3
    RealmRanks(28, 2) = 9
    '--------------------
    
    '--realm rank 4
    RealmRanks(29, 0) = 213875
    RealmRanks(29, 1) = 4
    RealmRanks(29, 2) = 0
    
    RealmRanks(30, 0) = 236375
    RealmRanks(30, 1) = 4
    RealmRanks(30, 2) = 1
    
    RealmRanks(31, 0) = 260400
    RealmRanks(31, 1) = 4
    RealmRanks(31, 2) = 2
    
    RealmRanks(32, 0) = 286000
    RealmRanks(32, 1) = 4
    RealmRanks(32, 2) = 3
    
    RealmRanks(33, 0) = 313225
    RealmRanks(33, 1) = 4
    RealmRanks(33, 2) = 4
    
    RealmRanks(34, 0) = 342125
    RealmRanks(34, 1) = 4
    RealmRanks(34, 2) = 5
    
    RealmRanks(35, 0) = 372750
    RealmRanks(35, 1) = 4
    RealmRanks(35, 2) = 6
    
    RealmRanks(36, 0) = 405150
    RealmRanks(36, 1) = 4
    RealmRanks(36, 2) = 7
    
    RealmRanks(37, 0) = 439375
    RealmRanks(37, 1) = 4
    RealmRanks(37, 2) = 8
    
    RealmRanks(38, 0) = 475475
    RealmRanks(38, 1) = 4
    RealmRanks(38, 2) = 9
    '--------------------
    
    '--realm rank 5
    RealmRanks(39, 0) = 513500
    RealmRanks(39, 1) = 5
    RealmRanks(39, 2) = 0
    
    RealmRanks(40, 0) = 553500
    RealmRanks(40, 1) = 5
    RealmRanks(40, 2) = 1
    
    RealmRanks(41, 0) = 595525
    RealmRanks(41, 1) = 5
    RealmRanks(41, 2) = 2
    
    RealmRanks(42, 0) = 639625
    RealmRanks(42, 1) = 5
    RealmRanks(42, 2) = 3
    
    RealmRanks(43, 0) = 685850
    RealmRanks(43, 1) = 5
    RealmRanks(43, 2) = 4
    
    RealmRanks(44, 0) = 734250
    RealmRanks(44, 1) = 5
    RealmRanks(44, 2) = 5
    
    RealmRanks(45, 0) = 784875
    RealmRanks(45, 1) = 5
    RealmRanks(45, 2) = 6
    
    RealmRanks(46, 0) = 837775
    RealmRanks(46, 1) = 5
    RealmRanks(46, 2) = 7
    
    RealmRanks(47, 0) = 893000
    RealmRanks(47, 1) = 5
    RealmRanks(47, 2) = 8
    
    RealmRanks(48, 0) = 950600
    RealmRanks(48, 1) = 5
    RealmRanks(48, 2) = 9
    '--------------------
    
    '--realm rank 6
    RealmRanks(49, 0) = 1010625
    RealmRanks(49, 1) = 6
    RealmRanks(49, 2) = 0
    
    RealmRanks(50, 0) = 1073125
    RealmRanks(50, 1) = 6
    RealmRanks(50, 2) = 1
    
    RealmRanks(51, 0) = 1138150
    RealmRanks(51, 1) = 6
    RealmRanks(51, 2) = 2
    
    RealmRanks(52, 0) = 1205750
    RealmRanks(52, 1) = 6
    RealmRanks(52, 2) = 3
    
    RealmRanks(53, 0) = 1275975
    RealmRanks(53, 1) = 6
    RealmRanks(53, 2) = 4
    
    RealmRanks(54, 0) = 1348875
    RealmRanks(54, 1) = 6
    RealmRanks(54, 2) = 5
    
    RealmRanks(55, 0) = 1424500
    RealmRanks(55, 1) = 6
    RealmRanks(55, 2) = 6
    
    RealmRanks(56, 0) = 1502900
    RealmRanks(56, 1) = 6
    RealmRanks(56, 2) = 7
    
    RealmRanks(57, 0) = 1584125
    RealmRanks(57, 1) = 6
    RealmRanks(57, 2) = 8
    
    RealmRanks(58, 0) = 1668225
    RealmRanks(58, 1) = 6
    RealmRanks(58, 2) = 9
    '--------------------
    
    '--realm rank 7
    RealmRanks(59, 0) = 1755250
    RealmRanks(59, 1) = 7
    RealmRanks(59, 2) = 0
    
    RealmRanks(60, 0) = 1845250
    RealmRanks(60, 1) = 7
    RealmRanks(60, 2) = 1
    
    RealmRanks(61, 0) = 1938275
    RealmRanks(61, 1) = 7
    RealmRanks(61, 2) = 2
    
    RealmRanks(62, 0) = 2034375
    RealmRanks(62, 1) = 7
    RealmRanks(62, 2) = 3
    
    RealmRanks(63, 0) = 2133600
    RealmRanks(63, 1) = 7
    RealmRanks(63, 2) = 4
    
    RealmRanks(64, 0) = 2236000
    RealmRanks(64, 1) = 7
    RealmRanks(64, 2) = 5
    
    RealmRanks(65, 0) = 2341625
    RealmRanks(65, 1) = 7
    RealmRanks(65, 2) = 6
    
    RealmRanks(66, 0) = 2450525
    RealmRanks(66, 1) = 7
    RealmRanks(66, 2) = 7
    
    RealmRanks(67, 0) = 2562750
    RealmRanks(67, 1) = 7
    RealmRanks(67, 2) = 8
    
    RealmRanks(68, 0) = 2678350
    RealmRanks(68, 1) = 7
    RealmRanks(68, 2) = 9
    '--------------------
    
    '--realm rank 8
    RealmRanks(69, 0) = 2797375
    RealmRanks(69, 1) = 8
    RealmRanks(69, 2) = 0
    
    RealmRanks(70, 0) = 2919875
    RealmRanks(70, 1) = 8
    RealmRanks(70, 2) = 1
    
    RealmRanks(71, 0) = 3045900
    RealmRanks(71, 1) = 8
    RealmRanks(71, 2) = 2
    
    RealmRanks(72, 0) = 3175500
    RealmRanks(72, 1) = 8
    RealmRanks(72, 2) = 3
    
    RealmRanks(73, 0) = 3308725
    RealmRanks(73, 1) = 8
    RealmRanks(73, 2) = 4
    
    RealmRanks(74, 0) = 3445625
    RealmRanks(74, 1) = 8
    RealmRanks(74, 2) = 5
    
    RealmRanks(75, 0) = 3586250
    RealmRanks(75, 1) = 8
    RealmRanks(75, 2) = 6
    
    RealmRanks(76, 0) = 3730650
    RealmRanks(76, 1) = 8
    RealmRanks(76, 2) = 7
    
    RealmRanks(77, 0) = 3878875
    RealmRanks(77, 1) = 8
    RealmRanks(77, 2) = 8
    
    RealmRanks(78, 0) = 4030975
    RealmRanks(78, 1) = 8
    RealmRanks(78, 2) = 9
    '--------------------
    
    '--realm rank 9
    RealmRanks(79, 0) = 4187000
    RealmRanks(79, 1) = 9
    RealmRanks(79, 2) = 0
    
    RealmRanks(80, 0) = 4347000
    RealmRanks(80, 1) = 9
    RealmRanks(80, 2) = 1
    
    RealmRanks(81, 0) = 4511025
    RealmRanks(81, 1) = 9
    RealmRanks(81, 2) = 2
    
    RealmRanks(82, 0) = 4679125
    RealmRanks(82, 1) = 9
    RealmRanks(82, 2) = 3
    
    RealmRanks(83, 0) = 4851350
    RealmRanks(83, 1) = 9
    RealmRanks(83, 2) = 4
    
    RealmRanks(84, 0) = 5027750
    RealmRanks(84, 1) = 9
    RealmRanks(84, 2) = 5
    
    RealmRanks(85, 0) = 5208375
    RealmRanks(85, 1) = 9
    RealmRanks(85, 2) = 6
    
    RealmRanks(86, 0) = 5393275
    RealmRanks(86, 1) = 9
    RealmRanks(86, 2) = 7
    
    RealmRanks(87, 0) = 5582500
    RealmRanks(87, 1) = 9
    RealmRanks(87, 2) = 8
    
    RealmRanks(88, 0) = 5776100
    RealmRanks(88, 1) = 9
    RealmRanks(88, 2) = 9
    '--------------------
    
    '--realm rank 10
    RealmRanks(89, 0) = 5974125
    RealmRanks(89, 1) = 10
    RealmRanks(89, 2) = 0
    
    RealmRanks(90, 0) = 6176625
    RealmRanks(90, 1) = 10
    RealmRanks(90, 2) = 1
    
    RealmRanks(91, 0) = 6383650
    RealmRanks(91, 1) = 10
    RealmRanks(91, 2) = 2
    
    RealmRanks(92, 0) = 6595250
    RealmRanks(92, 1) = 10
    RealmRanks(92, 2) = 3
    
    RealmRanks(93, 0) = 6811475
    RealmRanks(93, 1) = 10
    RealmRanks(93, 2) = 4
    
    RealmRanks(94, 0) = 7032375
    RealmRanks(94, 1) = 10
    RealmRanks(94, 2) = 5
    
    RealmRanks(95, 0) = 7258000
    RealmRanks(95, 1) = 10
    RealmRanks(95, 2) = 6
    
    RealmRanks(96, 0) = 7488400
    RealmRanks(96, 1) = 10
    RealmRanks(96, 2) = 7
    
    RealmRanks(97, 0) = 7723625
    RealmRanks(97, 1) = 10
    RealmRanks(97, 2) = 8
    
    RealmRanks(98, 0) = 7963725
    RealmRanks(98, 1) = 10
    RealmRanks(98, 2) = 9
    '--------------------
    
    '--realm rank 11
    RealmRanks(99, 0) = 8208750
    RealmRanks(99, 1) = 11
    RealmRanks(99, 2) = 0
    
    RealmRanks(100, 0) = 9111713
    RealmRanks(100, 1) = 11
    RealmRanks(100, 2) = 1
    
    RealmRanks(101, 0) = 10114001
    RealmRanks(101, 1) = 11
    RealmRanks(101, 2) = 2
    
    RealmRanks(102, 0) = 11226541
    RealmRanks(102, 1) = 11
    RealmRanks(102, 2) = 3
    
    RealmRanks(103, 0) = 12461460
    RealmRanks(103, 1) = 11
    RealmRanks(103, 2) = 4
    
    RealmRanks(104, 0) = 13832221
    RealmRanks(104, 1) = 11
    RealmRanks(104, 2) = 5
    
    RealmRanks(105, 0) = 15353765
    RealmRanks(105, 1) = 11
    RealmRanks(105, 2) = 6
    
    RealmRanks(106, 0) = 17042680
    RealmRanks(106, 1) = 11
    RealmRanks(106, 2) = 7
    
    RealmRanks(107, 0) = 18917374
    RealmRanks(107, 1) = 11
    RealmRanks(107, 2) = 8
    
    RealmRanks(108, 0) = 20998286
    RealmRanks(108, 1) = 11
    RealmRanks(108, 2) = 9
    '--------------------
    
    '--realm rank 12
    RealmRanks(109, 0) = 23308097
    RealmRanks(109, 1) = 12
    RealmRanks(109, 2) = 0
    
    RealmRanks(110, 0) = 25871988
    RealmRanks(110, 1) = 12
    RealmRanks(110, 2) = 1
    
    RealmRanks(111, 0) = 28717906
    RealmRanks(111, 1) = 12
    RealmRanks(111, 2) = 2
    
    RealmRanks(112, 0) = 31876876
    RealmRanks(112, 1) = 12
    RealmRanks(112, 2) = 3
    
    RealmRanks(113, 0) = 35383333
    RealmRanks(113, 1) = 12
    RealmRanks(113, 2) = 4
    
    RealmRanks(114, 0) = 39275499
    RealmRanks(114, 1) = 12
    RealmRanks(114, 2) = 5
    
    RealmRanks(115, 0) = 43595804
    RealmRanks(115, 1) = 12
    RealmRanks(115, 2) = 6
    
    RealmRanks(116, 0) = 48391343
    RealmRanks(116, 1) = 12
    RealmRanks(116, 2) = 7
    
    RealmRanks(117, 0) = 53714390
    RealmRanks(117, 1) = 12
    RealmRanks(117, 2) = 8
    
    RealmRanks(118, 0) = 59622973
    RealmRanks(118, 1) = 12
    RealmRanks(118, 2) = 9
    '---------------------
    
    '--realm rank 13
    RealmRanks(119, 0) = 66181501
    RealmRanks(119, 1) = 13
    RealmRanks(119, 2) = 0
    '---------------------
End Sub

Public Sub Init_RealmAbilities()
    
    'passive RAs
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).Name = "Augmented Acuity"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).Description = "Increases primary casting stat by the listed amount per level."
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_1 = "4 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_2 = "8 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_3 = "12 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_4 = "17 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_5 = "22 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_6 = "28 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_7 = "34 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_8 = "41 Acu"
    RealmAbilities(RA_AUG_ACU - RA_PASSIVE).EffectLevel_9 = "48 Acu"
    
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).Name = "Augmented Constitution"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).Description = "Increases primary casting stat by the listed amount per level."
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_1 = "4 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_2 = "8 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_3 = "12 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_4 = "17 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_5 = "22 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_6 = "28 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_7 = "34 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_8 = "41 Con"
    RealmAbilities(RA_AUG_CON - RA_PASSIVE).EffectLevel_9 = "48 Con"
    
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).Name = "Augmented Dexterity"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).Description = "Increases primary casting stat by the listed amount per level."
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_1 = "4 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_2 = "8 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_3 = "12 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_4 = "17 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_5 = "22 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_6 = "28 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_7 = "34 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_8 = "41 Dex"
    RealmAbilities(RA_AUG_DEX - RA_PASSIVE).EffectLevel_9 = "48 Dex"
    
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).Name = "Augmented Quickness"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).Description = "Increases primary casting stat by the listed amount per level."
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_1 = "4 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_2 = "8 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_3 = "12 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_4 = "17 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_5 = "22 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_6 = "28 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_7 = "34 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_8 = "41 Qui"
    RealmAbilities(RA_AUG_QUI - RA_PASSIVE).EffectLevel_9 = "48 Qui"
    
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).Name = "Augmented Strength"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).Description = "Increases primary casting stat by the listed amount per level."
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_1 = "4 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_2 = "8 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_3 = "12 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_4 = "17 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_5 = "22 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_6 = "28 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_7 = "34 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_8 = "41 Str"
    RealmAbilities(RA_AUG_STR - RA_PASSIVE).EffectLevel_9 = "48 Str"
    
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).Name = "Avoidance of Magic"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).Description = "Reduces all magic damage taken by the listed percentage. (This only works on damage. Does not work on disease, dots, or debuffs and does not affect the duration of crowd control spells)"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_1 = "2%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_2 = "3%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_3 = "5%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_4 = "7%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_5 = "10%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_6 = "12%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_7 = "15%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_8 = "17%"
    RealmAbilities(RA_AVOIDANCE_OF_MAGIC - RA_PASSIVE).EffectLevel_9 = "20%"
    
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).Name = "Determination"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).Description = "Reduces the duration of all crowd control spells by the listed percentage."
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_1 = "4%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_2 = "8%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_3 = "12%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_4 = "17%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_5 = "23%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_6 = "30%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_7 = "38%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_8 = "46%"
    RealmAbilities(RA_DETERMINATION - RA_PASSIVE).EffectLevel_9 = "55%"
    
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).Name = "Ethereal Bond"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).Description = "Increases power points by the listed amount."
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_1 = "15 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_2 = "25 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_3 = "40 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_4 = "55 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_5 = "75 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_6 = "100 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_7 = "130 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_8 = "165 Power"
    RealmAbilities(RA_ETHEREAL_BOND - RA_PASSIVE).EffectLevel_9 = "200 Power"
    
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).Name = "Falcon's Eye"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).Description = "Increases the chance of dealing a critical hit with archery by the listed percentage amount."
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_1 = "3%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_2 = "6%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_3 = "9%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_4 = "13%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_5 = "17%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_6 = "22%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_7 = "27%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_8 = "33%"
    RealmAbilities(RA_FALCONS_EYE - RA_PASSIVE).EffectLevel_9 = "39%"
    
    RealmAbilities(RA_LIFTER - RA_PASSIVE).Name = "Lifter"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).Description = "Increases the max encumbrance of the character and the speed at which the character can move rams they control by the listed percentage."
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_LIFTER - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_1 = "10%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_2 = "20%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_3 = "30%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_4 = "40%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_5 = "50%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_6 = "60%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_7 = "70%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_8 = "80%"
    RealmAbilities(RA_LIFTER - RA_PASSIVE).EffectLevel_9 = "90%"
        
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).Name = "Mastery of Blocking"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).Description = "Increases chance to block by the listed percentage."
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_1 = "2%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_2 = "4%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_3 = "6%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_4 = "9%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_5 = "12%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_6 = "15%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_7 = "18%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_8 = "21%"
    RealmAbilities(RA_MASTERY_OF_BLOCKING - RA_PASSIVE).EffectLevel_9 = "25%"
    
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).Name = "Mastery of Focus"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).Description = "Increases the level of all spells cast by the listed amount for out-right resistance purposes. (caps at level 50)"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_1 = "3 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_2 = "6 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_3 = "9 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_4 = "13 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_5 = "17 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_6 = "22 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_7 = "27 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_8 = "33 levels"
    RealmAbilities(RA_MASTERY_OF_FOCUS - RA_PASSIVE).EffectLevel_9 = "39 levels"
    
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).Name = "Mastery of Healing"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).Description = "Increases the effectiveness of healing spells by the listed percentage."
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_1 = "2%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_2 = "4%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_3 = "6%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_4 = "9%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_5 = "12%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_6 = "16%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_7 = "20%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_8 = "25%"
    RealmAbilities(RA_MASTERY_OF_HEALING - RA_PASSIVE).EffectLevel_9 = "30%"
    
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).Name = "Mastery of Magery"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).Description = "Additional effectiveness of magical damage by listed percentage."
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_1 = "2%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_2 = "3%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_3 = "4%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_4 = "6%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_5 = "8%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_6 = "10%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_7 = "12%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_8 = "14%"
    RealmAbilities(RA_MASTERY_OF_MAGERY - RA_PASSIVE).EffectLevel_9 = "16%"
    
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).Name = "Mastery of Pain"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).Description = "Increases chance to deal a critical hit in melee per listed percentage. (Passes on to Necro Pets)"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_1 = "3%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_2 = "6%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_3 = "9%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_4 = "13%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_5 = "17%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_6 = "22%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_7 = "27%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_8 = "33%"
    RealmAbilities(RA_MASTERY_OF_PAIN - RA_PASSIVE).EffectLevel_9 = "39%"
    
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).Name = "Mastery of Parrying"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).Description = "Increases chance to parry by the listed percentage."
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_1 = "2%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_2 = "4%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_3 = "6%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_4 = "9%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_5 = "12%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_6 = "15%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_7 = "18%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_8 = "21%"
    RealmAbilities(RA_MASTERY_OF_PARRYING - RA_PASSIVE).EffectLevel_9 = "25%"
    
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).Name = "Mastery of Stealth"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).Description = "Modifies detection and movement while stealthed. Camouflage negates the Mastery of Stealth bonus, allowing an archer to only be seen at the normal range. Has no effect on assassins detecting assassins."
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_1 = "10%/75"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_2 = "15%/125"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_3 = "20%/175"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_4 = "25%/235"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_5 = "30%/300"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_6 = "35%/375"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_7 = "40%/450"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_8 = "45%/535"
    RealmAbilities(RA_MASTERY_OF_STEALTH - RA_PASSIVE).EffectLevel_9 = "50%/625"
    
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).Name = "Physical Defense"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).Description = "Reduces all physical damage taken by the listed percentage."
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_1 = "2%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_2 = "4%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_3 = "6%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_4 = "9%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_5 = "12%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_6 = "16%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_7 = "20%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_8 = "25%"
    RealmAbilities(RA_PHYSICAL_DEFENSE - RA_PASSIVE).EffectLevel_9 = "30%"
    
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).Name = "Toughness"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).Description = "Increases hit points by the listed amount."
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_1 = "25 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_2 = "50 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_3 = "75 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_4 = "100 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_5 = "150 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_6 = "200 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_7 = "250 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_8 = "325 Hits"
    RealmAbilities(RA_TOUGHNESS - RA_PASSIVE).EffectLevel_9 = "400 Hits"
    
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).Name = "Veil Recovery"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).Description = "Reduces the duration of resurrection illnesses by the percentage listed."
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_1 = "10%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_2 = "15%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_3 = "20%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_4 = "30%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_5 = "40%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_6 = "50%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_7 = "60%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_8 = "70%"
    RealmAbilities(RA_VEIL_RECOVERY - RA_PASSIVE).EffectLevel_9 = "80%"
    
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).Name = "Wild Healing"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).Description = "Adds the listed percentage chance to critical heal on each target of a heal spell."
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_1 = "3%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_2 = "6%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_3 = "9%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_4 = "13%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_5 = "17%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_6 = "22%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_7 = "27%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_8 = "33%"
    RealmAbilities(RA_WILD_HEALING - RA_PASSIVE).EffectLevel_9 = "39%"
    
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).Name = "Wild Minion"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).Description = "Increases chance of pet dealing a critical hit in melee by the listed percentage."
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_1 = "3%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_2 = "6%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_3 = "9%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_4 = "13%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_5 = "17%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_6 = "22%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_7 = "27%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_8 = "33%"
    RealmAbilities(RA_WILD_MINION - RA_PASSIVE).EffectLevel_9 = "39%"
    
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).Name = "Wild Power"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).Type = "Passive"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).Description = "Increases chance to deal a critical hit with all spells that do damage, including DoTs, by listed percentage."
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_1 = 1
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_2 = 1
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_3 = 2
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_4 = 3
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_5 = 3
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_6 = 5
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_7 = 5
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_8 = 7
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).CostLevel_9 = 7
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_1 = "3%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_2 = "6%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_3 = "9%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_4 = "13%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_5 = "17%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_6 = "22%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_7 = "27%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_8 = "33%"
    RealmAbilities(RA_WILD_POWER - RA_PASSIVE).EffectLevel_9 = "39%"
    
    'active RAs
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).Name = "Adrenaline Rush"
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).Description = "Doubles the base melee damage for 20 sec. Reusable by time listed"
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).EffectLevel_1 = "20 min"
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).EffectLevel_2 = "15 min"
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).EffectLevel_3 = "10 min"
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).EffectLevel_4 = "7.5 min"
    RealmAbilities(RA_ADRENALINE_RUSH - RA_ACTIVE).EffectLevel_5 = "5 min"
        
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).Name = "Ameliorating Melodies"
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).Description = "Heals all members of the group (except the user) by the listed amount each tick for 30 sec.(10 total ticks)"
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).EffectLevel_1 = "100 Hits"
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).EffectLevel_2 = "175 Hits"
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).EffectLevel_3 = "250 Hits"
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).EffectLevel_4 = "325 Hits"
    RealmAbilities(RA_AMELIORATING_MELODIES - RA_ACTIVE).EffectLevel_5 = "400 Hits"
    
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).Name = "Anger of the Gods"
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).Description = "30 second group damage add that stacks with all other damage adds & ignores caps. DPS bonus as listed."
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).EffectLevel_1 = "10 DPS"
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).EffectLevel_2 = "15 DPS"
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).EffectLevel_3 = "20 DPS"
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).EffectLevel_4 = "25 DPS"
    RealmAbilities(RA_ANGER_OF_THE_GODS - RA_ACTIVE).EffectLevel_5 = "30 DPS"
    
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).Name = "Barrier of Fortitude"
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).Description = "Grants the group a melee absorption bonus based on the percentage listed. 30 second duration. (Does not stack with Soldier's Barricade or Bedazzling Aura.)"
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).EffectLevel_1 = "10%"
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).EffectLevel_2 = "15%"
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).EffectLevel_3 = "20%"
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).EffectLevel_4 = "30%"
    RealmAbilities(RA_BARRIER_OF_FORTITUDE - RA_ACTIVE).EffectLevel_5 = "40%"
    
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).Name = "Bedazzling Aura"
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).Description = "Grants the group increased resistance to magical damage for 29 sec by the percentage listed. (Does not stack with Soldier's Barricade or Barrier of Fortitude)"
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).EffectLevel_1 = "10%"
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).EffectLevel_2 = "15%"
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).EffectLevel_3 = "20%"
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).EffectLevel_4 = "30%"
    RealmAbilities(RA_BEDAZZLING_AURA - RA_ACTIVE).EffectLevel_5 = "40%"
    
    RealmAbilities(RA_CHARGE - RA_ACTIVE).Name = "Charge"
    RealmAbilities(RA_CHARGE - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_CHARGE - RA_ACTIVE).Description = "Unbreakable speed 3 for 15 sec. Immunity to crowd control spells. Target takes damage from snare/root spells that do damage. Reuse as listed."
    RealmAbilities(RA_CHARGE - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_CHARGE - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_CHARGE - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_CHARGE - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_CHARGE - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_CHARGE - RA_ACTIVE).EffectLevel_1 = "15 min"
    RealmAbilities(RA_CHARGE - RA_ACTIVE).EffectLevel_2 = "10 min"
    RealmAbilities(RA_CHARGE - RA_ACTIVE).EffectLevel_3 = "5 min"
    RealmAbilities(RA_CHARGE - RA_ACTIVE).EffectLevel_4 = "3 min"
    RealmAbilities(RA_CHARGE - RA_ACTIVE).EffectLevel_5 = "90 sec"
    
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).Name = "Concentration"
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).Description = "Resets the timer on quick-cast. Reuse as listed."
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).EffectLevel_1 = "15 min"
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).EffectLevel_2 = "9 min"
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).EffectLevel_3 = "3 min"
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).EffectLevel_4 = "90 sec"
    RealmAbilities(RA_CONCENTRATION - RA_ACTIVE).EffectLevel_5 = "30 sec"
    
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).Name = "Dashing Defense"
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).Description = "The tank can block and parry for all groupmates within a 1000 radius for the duration listed. Each attack only has the chance of being blocked/parried once, regardless of how many characters in a group have Dashing Defense active at one time."
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).EffectLevel_1 = "10 sec"
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).EffectLevel_2 = "20 sec"
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).EffectLevel_3 = "30 sec"
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).EffectLevel_4 = "45 sec"
    RealmAbilities(RA_DASHING_DEFENSE - RA_ACTIVE).EffectLevel_5 = "60 sec"
    
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).Name = "Decimation Trap"
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).Description = "AE damage trap with 350 radius. Damage as listed. The trap lasts ten minutes or until detonated (whichever comes first). Energy based. 2 second non-interruptible cast time"
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).EffectLevel_1 = "300 Dmg"
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).EffectLevel_2 = "450 Dmg"
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).EffectLevel_3 = "600 Dmg"
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).EffectLevel_4 = "750 Dmg"
    RealmAbilities(RA_DECIMATION_TRAP - RA_ACTIVE).EffectLevel_5 = "900 Dmg"
    
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).Name = "Divine Intervention"
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).Description = "Gives a pool of healing for the group. Combat damage is immediately healed from the pool. Pool size is based on the numbers listed. Does not heal the user. Buff Duration: 20 minutes"
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).EffectLevel_1 = "1000 Hits"
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).EffectLevel_2 = "1500 Hits"
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).EffectLevel_3 = "2000 Hits"
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).EffectLevel_4 = "2500 Hits"
    RealmAbilities(RA_DIVINE_INTERVENTION - RA_ACTIVE).EffectLevel_5 = "3000 Hits"
    
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).Name = "Dual Threat"
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).Description = "Grants a bonus chance to critical hit on both melee and magic based attacks. Percentage chance as listed. (While this stacks with Wild Power and Mastery of Pain, please note that there is a 50% hard cap on the chance to crit.)"
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).EffectLevel_1 = "5%"
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).EffectLevel_2 = "7%"
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).EffectLevel_3 = "10%"
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).EffectLevel_4 = "15%"
    RealmAbilities(RA_DUAL_THREAT - RA_ACTIVE).EffectLevel_5 = "20%"
    
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).Name = "First Aid"
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).Description = "Heals the user by the listed percentage."
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).EffectLevel_1 = "25%"
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).EffectLevel_2 = "40%"
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).EffectLevel_3 = "60%"
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).EffectLevel_4 = "80%"
    RealmAbilities(RA_FIRST_AID - RA_ACTIVE).EffectLevel_5 = "100%"
    
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).Name = "Ichor of the Deep"
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).Description = "Spirit-based root plus direct damage spell. 1875 range with 500 radius. Damage and duration as listed. Two second non-interruptible cast time."
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).EffectLevel_1 = "150/10 sec"
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).EffectLevel_2 = "275/15 sec"
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).EffectLevel_3 = "400/20 sec"
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).EffectLevel_4 = "500/25 sec"
    RealmAbilities(RA_ICHOR_OF_THE_DEEP - RA_ACTIVE).EffectLevel_5 = "600/30 sec"
        
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).Name = "Ignore Pain"
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).Description = "Heal that grants health equal to the percentage listed. Can be used when in combat."
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).EffectLevel_1 = "20%"
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).EffectLevel_2 = "35%"
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).EffectLevel_3 = "50%"
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).EffectLevel_4 = "65%"
    RealmAbilities(RA_IGNORE_PAIN - RA_ACTIVE).EffectLevel_5 = "80%"
    
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).Name = "Juggernaut"
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).Description = "Increases the effective level of the pet by the listed number of levels for 60 sec. (capped at level 70)"
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).EffectLevel_1 = "10 levels"
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).EffectLevel_2 = "15 levels"
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).EffectLevel_3 = "20 levels"
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).EffectLevel_4 = "25 levels"
    RealmAbilities(RA_JUGGERNAUT - RA_ACTIVE).EffectLevel_5 = "30 levels"
    
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).Name = "Mastery of Concentration"
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).Description = "Grants a 100% bonus to avoid being interrupted by any form of attack when casting a spell. The effect of the spell cast will be reduced to the percentages listed. (Necro version transfers to the pet)"
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).EffectLevel_1 = "25%"
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).EffectLevel_2 = "35%"
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).EffectLevel_3 = "50%"
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).EffectLevel_4 = "60%"
    RealmAbilities(RA_MASTERY_OF_CONCENTRATION - RA_ACTIVE).EffectLevel_5 = "75%"
    
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).Name = "Mystic Crystal Lore"
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).Description = "Grants a refresh of power based on the percentages listed. Cannot be used when in combat."
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).EffectLevel_1 = "25%"
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).EffectLevel_2 = "40%"
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).EffectLevel_3 = "60%"
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).EffectLevel_4 = "80%"
    RealmAbilities(RA_MYSTIC_CRYSTAL_LORE - RA_ACTIVE).EffectLevel_5 = "100%"
    
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).Name = "Negative Maelstrom"
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).Description = "A 2 second non-interruptible cast time cold based pulsing GTAE storm with 350 radius. Ground target range is 1500. The damage of the storm grows with each pulse."
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).EffectLevel_1 = "175 Dmg"
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).EffectLevel_2 = "260 Dmg"
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).EffectLevel_3 = "350 Dmg"
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).EffectLevel_4 = "425 Dmg"
    RealmAbilities(RA_NEGATIVE_MAELSTROM - RA_ACTIVE).EffectLevel_5 = "500 Dmg"
    
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).Name = "Perfect Recovery"
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).Description = "Instantly resurrects the target with no res effects with the listed amount of health and power."
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).EffectLevel_1 = "10%"
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).EffectLevel_2 = "25%"
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).EffectLevel_3 = "50%"
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).EffectLevel_4 = "75%"
    RealmAbilities(RA_PERFECT_RECOVERY - RA_ACTIVE).EffectLevel_5 = "100%"
    
    RealmAbilities(RA_PURGE - RA_ACTIVE).Name = "Purge"
    RealmAbilities(RA_PURGE - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_PURGE - RA_ACTIVE).Description = "Removes all negative effects but leaves any applicable immunity timers in place."
    RealmAbilities(RA_PURGE - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_PURGE - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_PURGE - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_PURGE - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_PURGE - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_PURGE - RA_ACTIVE).EffectLevel_1 = "5s / 15m"
    RealmAbilities(RA_PURGE - RA_ACTIVE).EffectLevel_2 = "2s / 15m"
    RealmAbilities(RA_PURGE - RA_ACTIVE).EffectLevel_3 = "0s / 15m"
    RealmAbilities(RA_PURGE - RA_ACTIVE).EffectLevel_4 = "0s / 10m"
    RealmAbilities(RA_PURGE - RA_ACTIVE).EffectLevel_5 = "0s / 5m"
    
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).Name = "Raging Power"
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).Description = "Grants a refresh of power based on the percentages listed. Can be used when in combat."
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).EffectLevel_1 = "20%"
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).EffectLevel_2 = "35%"
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).EffectLevel_3 = "50%"
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).EffectLevel_4 = "65%"
    RealmAbilities(RA_RAGING_POWER - RA_ACTIVE).EffectLevel_5 = "80%"
    
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).Name = "Reflex Attack"
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).Description = "Gives a chance to automatically counter-attack with an unstyled swing (or a swing from each hand in the case of duel wielding classes) anytime a hit is taken. Works against attacks from all 360 degrees with the chance based on the percentages listed."
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).EffectLevel_1 = "10%"
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).EffectLevel_2 = "20%"
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).EffectLevel_3 = "30%"
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).EffectLevel_4 = "40%"
    RealmAbilities(RA_REFLEX_ATTACK - RA_ACTIVE).EffectLevel_5 = "50%"
    
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).Name = "Second Wind"
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).Description = "Restores 100% of the user's endurance. Reuse as listed."
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).EffectLevel_1 = "15 min"
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).EffectLevel_2 = "10 min"
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).EffectLevel_3 = "5 min"
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).EffectLevel_4 = "3.5 min"
    RealmAbilities(RA_SECOND_WIND - RA_ACTIVE).EffectLevel_5 = "2 min"
    
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).Name = "Soldier's Barricade"
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).Description = "Grants the group an absorption bonus to all forms of damage based on the percentages listed. 30 Second duration. (Does not stack with Barrier of Fortitude or Bedazzling Aura)"
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).EffectLevel_1 = "5%"
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).EffectLevel_2 = "10%"
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).EffectLevel_3 = "15%"
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).EffectLevel_4 = "20%"
    RealmAbilities(RA_SOLDIERS_BARRICADE - RA_ACTIVE).EffectLevel_5 = "30%"
    
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).Name = "Speed of Sound"
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).Description = "Group ability that allows unstoppable speed 4 movement for the listed duration. Immunity to crowd control spells (if spell does damage, target takes damage). Breaks with any action taken except healing. Speedwarp negates."
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).EffectLevel_1 = "10 sec"
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).EffectLevel_2 = "20 sec"
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).EffectLevel_3 = "30 sec"
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).EffectLevel_4 = "45 sec"
    RealmAbilities(RA_SPEED_OF_SOUND - RA_ACTIVE).EffectLevel_5 = "60 sec"
    
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).Name = "Static Tempest"
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).Description = "Delivers a 360 radius targeted storm that procs a 3 second unresistible stun every 5 sec for the duration listed."
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).EffectLevel_1 = "15 sec"
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).EffectLevel_2 = "20 sec"
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).EffectLevel_3 = "25 sec"
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).EffectLevel_4 = "30 sec"
    RealmAbilities(RA_STATIC_TEMPEST - RA_ACTIVE).EffectLevel_5 = "35 sec"
    
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).Name = "Strike Prediction"
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).Description = "Grants all group members a chance to evade all melee and arrow attacks for 30 sec. This does not stack with any other chance to evade and will only benefit classes with no or very low chances of evading."
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).EffectLevel_1 = "5%"
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).EffectLevel_2 = "7%"
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).EffectLevel_3 = "10%"
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).EffectLevel_4 = "15%"
    RealmAbilities(RA_STRIKE_PREDICTION - RA_ACTIVE).EffectLevel_5 = "20%"
    
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).Name = "The Empty Mind"
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).Description = "Grants the user 45 sec of increased resistances to all magical damage by the percentage listed. This only works on damage and does not affect the duration of crowd control spell. (Necro version transfers to pet)"
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).EffectLevel_1 = "10%"
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).EffectLevel_2 = "15%"
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).EffectLevel_3 = "20%"
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).EffectLevel_4 = "25%"
    RealmAbilities(RA_EMPTY_MIND - RA_ACTIVE).EffectLevel_5 = "30%"
    
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).Name = "Thornweed Field"
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).Description = "Creates a field of thorns that damage and snare all enemies caught within. 500 radius. Essence Damage. Pulses every 3 sec. 2 second non-interruptible cast time. 1500 range"
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).EffectLevel_1 = "25/10s"
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).EffectLevel_2 = "50/15s"
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).EffectLevel_3 = "100/20s"
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).EffectLevel_4 = "175/25s"
    RealmAbilities(RA_THORNWEED_FIELD - RA_ACTIVE).EffectLevel_5 = "250/30s"
       
    RealmAbilities(RA_VANISH - RA_ACTIVE).Name = "Vanish"
    RealmAbilities(RA_VANISH - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_VANISH - RA_ACTIVE).Description = "Unbreakable stealth which purges DoTs and Bleeds, gives immunity to crowd control, and gives an increase in movement speed as listed. Lasts for 1 to 5 sec depending on level of Vanish.  Cannot attack for 30s, Silenced for 15s"
    RealmAbilities(RA_VANISH - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_VANISH - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_VANISH - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_VANISH - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_VANISH - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_VANISH - RA_ACTIVE).EffectLevel_1 = "S0/3 sec"
    RealmAbilities(RA_VANISH - RA_ACTIVE).EffectLevel_2 = "S1/3 sec"
    RealmAbilities(RA_VANISH - RA_ACTIVE).EffectLevel_3 = "S3/4 sec"
    RealmAbilities(RA_VANISH - RA_ACTIVE).EffectLevel_4 = "S4/5 sec"
    RealmAbilities(RA_VANISH - RA_ACTIVE).EffectLevel_5 = "S5/6 sec"
    
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).Name = "Vehement Renewal"
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).Description = "Instantly heals all group members (except the user) within 2000 range for the amount listed."
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).EffectLevel_1 = "375 Hits"
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).EffectLevel_2 = "525 Hits"
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).EffectLevel_3 = "750 Hits"
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).EffectLevel_4 = "1125 Hits"
    RealmAbilities(RA_VEHEMENT_RENEWAL - RA_ACTIVE).EffectLevel_5 = "1500 Hits"
    
    RealmAbilities(RA_VIPER - RA_ACTIVE).Name = "Viper"
    RealmAbilities(RA_VIPER - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_VIPER - RA_ACTIVE).Description = "Increases the damage of poisons by the listed amount."
    RealmAbilities(RA_VIPER - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_VIPER - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_VIPER - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_VIPER - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_VIPER - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_VIPER - RA_ACTIVE).EffectLevel_1 = "25%"
    RealmAbilities(RA_VIPER - RA_ACTIVE).EffectLevel_2 = "35%"
    RealmAbilities(RA_VIPER - RA_ACTIVE).EffectLevel_3 = "50%"
    RealmAbilities(RA_VIPER - RA_ACTIVE).EffectLevel_4 = "75%"
    RealmAbilities(RA_VIPER - RA_ACTIVE).EffectLevel_5 = "100%"
    
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).Name = "Volcanic Pillar"
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).Description = "Heat based AE damage spell with 500 radius. Damage as listed. 2 second non-interruptible cast time. 1500 range Target is enemy."
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).EffectLevel_1 = "200 Dmg"
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).EffectLevel_2 = "350 Dmg"
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).EffectLevel_3 = "500 Dmg"
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).EffectLevel_4 = "625 Dmg"
    RealmAbilities(RA_VOLCANIC_PILLAR - RA_ACTIVE).EffectLevel_5 = "750 Dmg"
    
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).Name = "Wrath of Champions"
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).Type = "Active"
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).Description = "Spirit Based instantly cast PBAE with 150 radius that does the listed damage."
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).CostLevel_1 = 5
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).CostLevel_2 = 5
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).CostLevel_3 = 5
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).CostLevel_4 = 7
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).CostLevel_5 = 8
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).EffectLevel_1 = "200 Dmg"
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).EffectLevel_2 = "350 Dmg"
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).EffectLevel_3 = "500 Dmg"
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).EffectLevel_4 = "625 Dmg"
    RealmAbilities(RA_WRATH_OF_CHAMPIONS - RA_ACTIVE).EffectLevel_5 = "750 Dmg"
    
    'class unique RA
    RealmAbilities(RA_ALLURE_OF_DEATH - RA_CLASS).Name = "Allure of Death"
    RealmAbilities(RA_ALLURE_OF_DEATH - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_ALLURE_OF_DEATH - RA_CLASS).Description = "Bonedancer appearance is changed to skeletal form for 60 sec to confuse enemy. When in skeletal form, the Bonedancer has a 75% chance of out-right resisting Nearsight and all Crowd Control spells."
    
    RealmAbilities(RA_ARMS_LENGTH - RA_CLASS).Name = "Arms Length"
    RealmAbilities(RA_ARMS_LENGTH - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_ARMS_LENGTH - RA_CLASS).Description = "10 second unbreakable burst of extreme speed."
    
    RealmAbilities(RA_BADGE_OF_VALOR - RA_CLASS).Name = "Badge of Valor"
    RealmAbilities(RA_BADGE_OF_VALOR - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_BADGE_OF_VALOR - RA_CLASS).Description = "Champion's melee damage for the next 20 sec will be INCREASED by the targets armor-based ABS instead of decreased."
    
    RealmAbilities(RA_BLADE_BARRIER - RA_CLASS).Name = "Blade Barrier"
    RealmAbilities(RA_BLADE_BARRIER - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_BLADE_BARRIER - RA_CLASS).Description = "For 30 sec the blademaster parries 90% of all melee attacks and receives 25% less damage from all attacks. The blademaster is unable to attack during this time. If the blademaster attempts a style while this effect is still up, it will cancel the effect."
    
    RealmAbilities(RA_BLINDING_DUST - RA_CLASS).Name = "Blinding Dust"
    RealmAbilities(RA_BLINDING_DUST - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_BLINDING_DUST - RA_CLASS).Description = "Insta-cast PBAE Attack that causes the enemy to have a 25% chance to fumble melee/bow attacks as well as 50% nearsight for the next 15 sec."
    
    RealmAbilities(RA_FUELED_BY_RAGE - RA_CLASS).Name = "Fueled By Rage"
    RealmAbilities(RA_FUELED_BY_RAGE - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_FUELED_BY_RAGE - RA_CLASS).Description = "This ability will reduce all damage that the Savage takes for the next 30 seconds by 20%. Also, 50% of the damage that this ability reduces will be returned as healing. Usable every 10 minutes."
    
    RealmAbilities(RA_BOILING_CAULDRON - RA_CLASS).Name = "Boiling Cauldron"
    RealmAbilities(RA_BOILING_CAULDRON - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_BOILING_CAULDRON - RA_CLASS).Description = "Summons a large cauldron that boils in place for 3.5 sec before spilling and doing 650 damage to all those nearby."
    
    RealmAbilities(RA_CALL_OF_DARKNESS - RA_CLASS).Name = "Call of Darkness"
    RealmAbilities(RA_CALL_OF_DARKNESS - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_CALL_OF_DARKNESS - RA_CLASS).Description = "When active, the necromancer can summon a pet with only a 3 second cast time. The effect remains active for 15 minutes, or until a pet is summoned. "
    
    RealmAbilities(RA_CALMING_NOTES - RA_CLASS).Name = "Calming Notes"
    RealmAbilities(RA_CALMING_NOTES - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_CALMING_NOTES - RA_CLASS).Description = "Insta-cast spell that mesmerizes all enemy pets, players, and guards within 600 radius for 20 sec."
    
    RealmAbilities(RA_CHAIN_LIGHTNING - RA_CLASS).Name = "Chain Lightning"
    RealmAbilities(RA_CHAIN_LIGHTNING - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_CHAIN_LIGHTNING - RA_CLASS).Description = "Lightning bolt that has a chance to hit up to 5 targets (player or pet). If only one target is available, it will hit once. Otherwise the spell has a chance to jump from target to target and back to the prior target. Damage is reduced by 25% with each jump."
    
    RealmAbilities(RA_COMBAT_AWARENESS - RA_CLASS).Name = "Combat Awareness"
    RealmAbilities(RA_COMBAT_AWARENESS - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_COMBAT_AWARENESS - RA_CLASS).Description = "The Hero gains a 360 degree 50% evade for 30 sec. During this time, the Hero will also be snared by 50%."
    
    RealmAbilities(RA_DESPERATE_BOWMAN - RA_CLASS).Name = "Desperate Bowman"
    RealmAbilities(RA_DESPERATE_BOWMAN - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_DESPERATE_BOWMAN - RA_CLASS).Description = "This ability is used with a bow. Does 300 damage and a 5 second (non resistible) stun. Bow and melee cannot be used for 15 sec afterwards. Also grants the Ranger 60% unbreakable speed increase for 15 sec."
    
    RealmAbilities(RA_DREAMWEAVER - RA_CLASS).Name = "Dreamweaver"
    RealmAbilities(RA_DREAMWEAVER - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_DREAMWEAVER - RA_CLASS).Description = "Creates an illusion that makes the Bard look like a different race and class for a 5 minutes. This ability will now also grant a reactive snare proc. This proc will have a 20% proc rate and will apply a 40% snare that will last 15 seconds. This snare will not leave immunity. The reactive proc will be removed if Dreamweaver is removed"
    
    RealmAbilities(RA_ENTWINING_SNAKES - RA_CLASS).Name = "Entwining Snakes"
    RealmAbilities(RA_ENTWINING_SNAKES - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_ENTWINING_SNAKES - RA_CLASS).Description = "Insta-cast spell that is PBAE 50% snare lasting 20 sec with a 350 unit radius. Snare breaks on attack. Additionally the Hunter receives a 60% unbreakable speed increase for 15 sec."
    
    RealmAbilities(RA_VOICE_SKADI - RA_CLASS).Name = "Voice of Skadi"
    RealmAbilities(RA_VOICE_SKADI - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_VOICE_SKADI - RA_CLASS).Description = "This ability will render the Slakd immune to crowd control effects for 15 seconds. This ability is usable every 5 minutes."
    
    RealmAbilities(RA_FANATICISM - RA_CLASS).Name = "Fanaticism"
    RealmAbilities(RA_FANATICISM - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_FANATICISM - RA_CLASS).Description = "All Heretic groupmates who are able to bind at a keep or tower lord receive a reduction in all spell damage taken for 45 sec."
    
    RealmAbilities(RA_FEROCIOUS_WILL - RA_CLASS).Name = "Ferocious Will"
    RealmAbilities(RA_FEROCIOUS_WILL - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_FEROCIOUS_WILL - RA_CLASS).Description = "Gives the Berserker an 25% Absorb buff, 25% bonus to all resistances and a 25% chance to shrug off crowd control spells for 30 seconds."
    
    RealmAbilities(RA_FUNGAL_UNION - RA_CLASS).Name = "Fungal Union"
    RealmAbilities(RA_FUNGAL_UNION - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_FUNGAL_UNION - RA_CLASS).Description = "Turns the animist into a mushroom for 60 sec. Does not break on attack. Grants the animist a 50% chance of not spending power for each spell cast during the duration."
    
    RealmAbilities(RA_FURY_OF_NATURE - RA_CLASS).Name = "Fury of Nature"
    RealmAbilities(RA_FURY_OF_NATURE - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_FURY_OF_NATURE - RA_CLASS).Description = "Double style damage for 30 sec. All damage done returns 100% to the group in spread heal form. (excluding the warden)"
    
    RealmAbilities(RA_MARK_OF_PREY - RA_CLASS).Name = "Mark of Prey"
    RealmAbilities(RA_MARK_OF_PREY - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_MARK_OF_PREY - RA_CLASS).Description = "Grants all members of the Vampiir's group a 30 second damage add that stacks with all other forms of damage add. All damage done via the damage add will be returned to the Vampiir as power."
    
    RealmAbilities(RA_MINION_RESCUE - RA_CLASS).Name = "Minion Rescue"
    RealmAbilities(RA_MINION_RESCUE - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_MINION_RESCUE - RA_CLASS).Description = "2s uninterruptible PBAE that summons 1 level 50 fire elemental for every enemy within 500 radius (max 8). Pets have 50 hit points, but proc a 3 sec stun (duration unaffected by resists). The pets have a max duration of 6 sec."
    
    RealmAbilities(RA_NATURES_WOMB - RA_CLASS).Name = "Nature's Womb"
    RealmAbilities(RA_NATURES_WOMB - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_NATURES_WOMB - RA_CLASS).Description = "Insta-cast spell that silences the druid for 5 sec and converts all damage taken into healing."
            
    RealmAbilities(RA_RESOLUTE_MINION - RA_CLASS).Name = "Resolute Minion"
    RealmAbilities(RA_RESOLUTE_MINION - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_RESOLUTE_MINION - RA_CLASS).Description = "Pet is immune to all forms of Crowd Control for 60 sec. Will not purge any CC that already exists on the pet. "
    
    RealmAbilities(RA_RESTORATIVE_MIND - RA_CLASS).Name = "Restorative Mind"
    RealmAbilities(RA_RESTORATIVE_MIND - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_RESTORATIVE_MIND - RA_CLASS).Description = "Group Frigg that heals health, power, and endurance over 30 sec for a total of 50%. (5% is granted every 3 sec regardless of combat state)"
    
    RealmAbilities(RA_RETRIBUTION_OF_THE_FAITHFUL - RA_CLASS).Name = "Retribution of the Faithful"
    RealmAbilities(RA_RETRIBUTION_OF_THE_FAITHFUL - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_RETRIBUTION_OF_THE_FAITHFUL - RA_CLASS).Description = "30 second buff that has a chance to proc a 3 second (duration undiminished by resists) stun on any melee attack on the cleric."
    
    RealmAbilities(RA_RUNE_OF_UTTER_AGILITY - RA_CLASS).Name = "Rune of Utter Agility"
    RealmAbilities(RA_RUNE_OF_UTTER_AGILITY - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_RUNE_OF_UTTER_AGILITY - RA_CLASS).Description = "Runemaster gets a 90% chance to evade all melee attacks (regardless of direction) for 15 sec."
    
    RealmAbilities(RA_PROTECTION_OF_THE_UNDERHILL - RA_CLASS).Name = "Protection of the Underhill"
    RealmAbilities(RA_PROTECTION_OF_THE_UNDERHILL - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_PROTECTION_OF_THE_UNDERHILL - RA_CLASS).Description = "2000 point, 50% ablative that absorbs both magic and melee damage. Effect last 3 minutes, usable every 10 minutes. Requires a living pet to activate."
    
    RealmAbilities(RA_SELECTIVE_BLINDNESS - RA_CLASS).Name = "Selective Blindness"
    RealmAbilities(RA_SELECTIVE_BLINDNESS - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SELECTIVE_BLINDNESS - RA_CLASS).Description = "When cast on an enemy player or pet, this debuff will prevent that player or pet from being able to attack the mentalist for 20 sec. (unaffected by resists). 1500 range, 150 radius, insta-cast. Effect drops on any target the Mentalist attacks."
    
    RealmAbilities(RA_SELFLESS_DEVOTION - RA_CLASS).Name = "Selfless Devotion"
    RealmAbilities(RA_SELFLESS_DEVOTION - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SELFLESS_DEVOTION - RA_CLASS).Description = "Triples the effect of the paladin healing chant for 1 minute on all groupmates excluding the Paladin himself."
    
    RealmAbilities(RA_SHIELD_OF_IMMUNITY - RA_CLASS).Name = "Shield of Immunity"
    RealmAbilities(RA_SHIELD_OF_IMMUNITY - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SHIELD_OF_IMMUNITY - RA_CLASS).Description = "Shield that absorbs 90% melee/archer damage for 20 sec."
    
    RealmAbilities(RA_SHIELD_TRIP - RA_CLASS).Name = "Shield Trip"
    RealmAbilities(RA_SHIELD_TRIP - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SHIELD_TRIP - RA_CLASS).Description = "Throws shield at target, causing 300 damage and rooting them in place for 10 sec (undiminished by resists). Scout cannot attack for 15 sec afterwards. Additionally gives the Scout a 60% unbreakable speed increase for 15 seconds."
    
    RealmAbilities(RA_SOLDIERS_CITADEL - RA_CLASS).Name = "Soldier's Citadel"
    RealmAbilities(RA_SOLDIERS_CITADEL - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SOLDIERS_CITADEL - RA_CLASS).Description = "This ability grants a 50% bonus to parry and block rates for the Armsman for 30 sec, but -10% bloc/parry rates for 15 sec after."
    
    RealmAbilities(RA_SONIC_BARRIER - RA_CLASS).Name = "Sonic Barrier"
    RealmAbilities(RA_SONIC_BARRIER - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SONIC_BARRIER - RA_CLASS).Description = "Increases the armor ABS of all groupmates for 45 sec by a multiple of their existing ABS."
    
    RealmAbilities(RA_SOUL_QUENCH - RA_CLASS).Name = "Soul Quench"
    RealmAbilities(RA_SOUL_QUENCH - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SOUL_QUENCH - RA_CLASS).Description = "Insta-PBAE attack that drains 250 points (modified up or down by the Reavers SR level) from all nearby enemies and returns 75% to the Reaver."
    
    RealmAbilities(RA_SPIRIT_MARTYR - RA_CLASS).Name = "Spirit Martyr"
    RealmAbilities(RA_SPIRIT_MARTYR - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SPIRIT_MARTYR - RA_CLASS).Description = "Heals all in group, amount healed is dependant on the health of the pet at the time of activation. 1200 total healing pool for a full health pet. Max of 600 hit points per any group member."
    
    RealmAbilities(RA_SPUTINS_LEGACY - RA_CLASS).Name = "Sputin's Legacy"
    RealmAbilities(RA_SPUTINS_LEGACY - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SPUTINS_LEGACY - RA_CLASS).Description = "Allows the healer to Cheat Death. If you take enough damage to die while the buff is active, it heals for a minimum of 30% health and you will not die. Also procs a 5 second spell and damage immunity."

    RealmAbilities(RA_TESTUDO - RA_CLASS).Name = "Testudo"
    RealmAbilities(RA_TESTUDO - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_TESTUDO - RA_CLASS).Description = "Warrior with shield equipped covers up and takes 90% less damage for all attacks for 45 sec. Speed is reduced (speed buffs have no effect). Using a style will break Testudo. Effective versus realm enemies only."
    
    RealmAbilities(RA_VALE_DEFENSE - RA_CLASS).Name = "Vale Defense"
    RealmAbilities(RA_VALE_DEFENSE - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_VALE_DEFENSE - RA_CLASS).Description = "Gives the group a 1000 point 50% ablative that lasts for 10 minutes or until depleted."
    
    RealmAbilities(RA_VALHALLAS_BLESSING - RA_CLASS).Name = "Valhalla's Blessing"
    RealmAbilities(RA_VALHALLAS_BLESSING - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_VALHALLAS_BLESSING - RA_CLASS).Description = "Each spell and style used by group members has a 75% chance of not costing power or endurance for 30 sec."
    
    RealmAbilities(RA_WALL_OF_FLAME - RA_CLASS).Name = "Wall of Flame"
    RealmAbilities(RA_WALL_OF_FLAME - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_WALL_OF_FLAME - RA_CLASS).Description = "Insta-cast spell that drops a ward that pulses a 150 radius PBAE fire based for 15 sec. Pulse is 400 points of damage every 3 sec. Wizard will receive a 200 value, 100% ablative to absorb melee and magic damage every 3 seconds for the duration of the wall."
    
    RealmAbilities(RA_WHIRLING_STAFF - RA_CLASS).Name = "Whirling Staff"
    RealmAbilities(RA_WHIRLING_STAFF - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_WHIRLING_STAFF - RA_CLASS).Description = "PBAE attack that does moderate damage and makes all melee targets in 350 radius unable to attack for 6 sec."
    
    RealmAbilities(RA_GIFT_OF_PERIZOR - RA_CLASS).Name = "The Gift of Perizor"
    RealmAbilities(RA_GIFT_OF_PERIZOR - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_GIFT_OF_PERIZOR - RA_CLASS).Description = "Reduces damage on all players in the user's group by 25% for 60 seconds. Damage reduced by this ability is returned to the user in power. Ability has a 10 min reuse timer"
    
    RealmAbilities(RA_OVERWHELM - RA_CLASS).Name = "Overwhelm"
    RealmAbilities(RA_OVERWHELM - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_OVERWHELM - RA_CLASS).Description = "Will give the Infiltrator a 15% increased chance to bypass their target’s block, parry, and evade defenses for 30 seconds, and can be used every 5 minutes."
    
    RealmAbilities(RA_SHADOW_SHROUD - RA_CLASS).Name = "Shadow Shroud"
    RealmAbilities(RA_SHADOW_SHROUD - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_SHADOW_SHROUD - RA_CLASS).Description = "Will reduce all incoming damage by 15% and increase the Nightshade’s chance to be missed by 15% for 30 seconds, and can be used every 5 minutes."
    
    RealmAbilities(RA_BLOOD_DRINKING - RA_CLASS).Name = "Blood Drinking"
    RealmAbilities(RA_BLOOD_DRINKING - RA_CLASS).Type = "Unique"
    RealmAbilities(RA_BLOOD_DRINKING - RA_CLASS).Description = "Will cause the Shadowblade to be healed for 15% of all damage he does for 30 seconds, and can be used every 5 minutes."
    
End Sub

Public Sub SetPassiveRealmAbilityInfo(RA_ID As Long, _
                               ByRef RANAME As Label, _
                               ByRef LEVEL1 As Label, _
                               ByRef LEVEL2 As Label, _
                               ByRef LEVEL3 As Label, _
                               ByRef LEVEL4 As Label, _
                               ByRef LEVEL5 As Label, _
                               ByRef LEVEL6 As Label, _
                               ByRef LEVEL7 As Label, _
                               ByRef LEVEL8 As Label, _
                               ByRef LEVEL9 As Label, _
                               ByRef PASSIVECOST As Label, _
                               ByRef PASSIVEBONUS As Label, _
                               ByRef INCREASE As Image, _
                               ByRef DECREASE As Image)

    'set name
    RANAME.Caption = RealmAbilities(RA_ID - RA_PASSIVE).Name
    'set description
    RANAME.Tag = RealmAbilities(RA_ID - RA_PASSIVE).Description
    
    'set level costs
    LEVEL1.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_1
    LEVEL2.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_2
    LEVEL3.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_3
    LEVEL4.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_4
    LEVEL5.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_5
    LEVEL6.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_6
    LEVEL7.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_7
    LEVEL8.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_8
    LEVEL9.Tag = RealmAbilities(RA_ID - RA_PASSIVE).CostLevel_9
    
    'set cost of level 1 on display
    PASSIVECOST.Caption = "Pts: " & LEVEL1.Tag
    'set the true array Index for reference later
    PASSIVEBONUS.Tag = (RA_ID - RA_PASSIVE)
    
    'set the current level to zero (no ra's trained)
    INCREASE.Tag = 0
    DECREASE.Tag = 0
        
End Sub

Public Sub SetActiveRealmAbilityInfo(RA_ID As Long, _
                                     ByRef RANAME As Label, _
                                     ByRef LEVEL1 As Label, _
                                     ByRef LEVEL2 As Label, _
                                     ByRef LEVEL3 As Label, _
                                     ByRef LEVEL4 As Label, _
                                     ByRef LEVEL5 As Label, _
                                     ByRef ACTIVECOST As Label, _
                                     ByRef ACTIVEBONUS As Label, _
                                     ByRef INCREASE As Image, _
                                     ByRef DECREASE As Image)

    'set name
    RANAME.Caption = RealmAbilities(RA_ID - RA_ACTIVE).Name
    'set description
    RANAME.Tag = RealmAbilities(RA_ID - RA_ACTIVE).Description
    
    'set level costs
    LEVEL1.Tag = RealmAbilities(RA_ID - RA_ACTIVE).CostLevel_1
    LEVEL2.Tag = RealmAbilities(RA_ID - RA_ACTIVE).CostLevel_2
    LEVEL3.Tag = RealmAbilities(RA_ID - RA_ACTIVE).CostLevel_3
    LEVEL4.Tag = RealmAbilities(RA_ID - RA_ACTIVE).CostLevel_4
    LEVEL5.Tag = RealmAbilities(RA_ID - RA_ACTIVE).CostLevel_5
    
    'set cost of level 1 on display
    ACTIVECOST.Caption = "Pts: " & LEVEL1.Tag
    'set the true array Index for reference later
    ACTIVEBONUS.Tag = (RA_ID - RA_ACTIVE)
    
    'set the current level to zero (no ra's trained)
    INCREASE.Tag = 0
    DECREASE.Tag = 0
    
End Sub

Public Sub SetUniqueRealmAbilityInfo(RA_ID As Long, ByRef RANAME As Label)

    RANAME.Caption = RealmAbilities(RA_ID - RA_CLASS).Name
    RANAME.Tag = RealmAbilities(RA_ID - RA_CLASS).Description
    
End Sub
                                     
Public Sub SetRealmAbilityToMatrix(RA_ID As Long, RA_LEVEL As Long, ByRef tToon As TOON_TYPE)

    Select Case RA_ID
        Case RA_AUG_ACU
            Select Case RA_LEVEL
                Case 0
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = 0
                Case 1
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_1)
                Case 2
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_2)
                Case 3
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_3)
                Case 4
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_4)
                Case 5
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_5)
                Case 6
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_6)
                Case 7
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_7)
                Case 8
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_8)
                Case 9
                    tToon.STAT_MATRIX(SM_ACU, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_9)
            End Select
        Case RA_AUG_CON
            Select Case RA_LEVEL
                Case 0
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = 0
                Case 1
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_1)
                Case 2
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_2)
                Case 3
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_3)
                Case 4
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_4)
                Case 5
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_5)
                Case 6
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_6)
                Case 7
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_7)
                Case 8
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_8)
                Case 9
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_9)
            End Select
        Case RA_AUG_DEX
            Select Case RA_LEVEL
                Case 0
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = 0
                Case 1
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_1)
                Case 2
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_2)
                Case 3
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_3)
                Case 4
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_4)
                Case 5
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_5)
                Case 6
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_6)
                Case 7
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_7)
                Case 8
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_8)
                Case 9
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_9)
            End Select
        Case RA_AUG_QUI
            Select Case RA_LEVEL
                Case 0
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = 0
                Case 1
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_1)
                Case 2
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_2)
                Case 3
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_3)
                Case 4
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_4)
                Case 5
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_5)
                Case 6
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_6)
                Case 7
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_7)
                Case 8
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_8)
                Case 9
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_9)
            End Select
        Case RA_AUG_STR
            Select Case RA_LEVEL
                Case 0
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = 0
                Case 1
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_1)
                Case 2
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_2)
                Case 3
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_3)
                Case 4
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_4)
                Case 5
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_5)
                Case 6
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_6)
                Case 7
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_7)
                Case 8
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_8)
                Case 9
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_9)
            End Select
        Case RA_TOUGHNESS
            Select Case RA_LEVEL
                Case 0
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = 0
                Case 1
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_1)
                Case 2
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_2)
                Case 3
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_3)
                Case 4
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_4)
                Case 5
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_5)
                Case 6
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_6)
                Case 7
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_7)
                Case 8
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_8)
                Case 9
                    tToon.STAT_MATRIX(SM_HIT, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_9)
            End Select
        Case RA_ETHEREAL_BOND
            Select Case RA_LEVEL
                Case 0
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = 0
                Case 1
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_1)
                Case 2
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_2)
                Case 3
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_3)
                Case 4
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_4)
                Case 5
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_5)
                Case 6
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_6)
                Case 7
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_7)
                Case 8
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_8)
                Case 9
                    tToon.STAT_MATRIX(SM_POW, SM_LOC_REALMABILITY) = Val(RealmAbilities(RA_ID - RA_PASSIVE).EffectLevel_9)
            End Select
    End Select

End Sub
