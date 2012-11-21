Attribute VB_Name = "ClassInit"
Option Explicit
'# 1 str, 2 con, 3 dex, 4 qui, 5 int, 6 emp, 7 pie, 8 cha

Public Sub SetBuffs(ByRef tToon As TOON_TYPE)

    If tToon.REALM = REALM_HIBERNIA And tToon.CLASS = TCH_VAMPIIR Then
        tToon.STAT_MATRIX(SM_STR, SM_LOC_BUFFS) = 0
        tToon.STAT_MATRIX(SM_CON, SM_LOC_BUFFS) = 0
        tToon.STAT_MATRIX(SM_DEX, SM_LOC_BUFFS) = 0
        tToon.STAT_MATRIX(SM_QUI, SM_LOC_BUFFS) = 0
        tToon.STAT_MATRIX(SM_INT, SM_LOC_BUFFS) = 0
        tToon.STAT_MATRIX(SM_EMP, SM_LOC_BUFFS) = 0
        tToon.STAT_MATRIX(SM_PIE, SM_LOC_BUFFS) = 0
        tToon.STAT_MATRIX(SM_CHA, SM_LOC_BUFFS) = 0
    Else
        tToon.STAT_MATRIX(SM_STR, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_STR).Text)
        tToon.STAT_MATRIX(SM_CON, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_CON).Text)
        tToon.STAT_MATRIX(SM_DEX, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_DEX).Text)
        tToon.STAT_MATRIX(SM_QUI, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_QUI).Text)
        tToon.STAT_MATRIX(SM_INT, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_INT).Text)
        tToon.STAT_MATRIX(SM_EMP, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_EMP).Text)
        tToon.STAT_MATRIX(SM_PIE, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_PIE).Text)
        tToon.STAT_MATRIX(SM_CHA, SM_LOC_BUFFS) = Val(Options.txt_BuffValue(WS_ATTR_CHA).Text)
    End If

End Sub

Public Sub Init_ARMSMAN(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCA_ARMSMAN
        
        .pStat = SM_STR 'str
        .sStat = SM_CON 'con
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 2
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_BATTLEMASTER
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(8) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(10) = RA_TOUGHNESS
        .REALM_ABILITIES(11) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(12) = RA_DASHING_DEFENSE
        .REALM_ABILITIES(13) = RA_FIRST_AID
        .REALM_ABILITIES(14) = RA_IGNORE_PAIN
        .REALM_ABILITIES(15) = RA_PURGE
        .REALM_ABILITIES(16) = RA_SECOND_WIND
        .REALM_ABILITIES(17) = RA_SOLDIERS_BARRICADE
        .REALM_ABILITIES(18) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(19) = RA_SOLDIERS_CITADEL
        
        .SPELL_COUNT = 0
        .STYLE_COUNT = 8
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAlb_Crush(0)
        Call InitStyleAlb_Slash(1)
        Call InitStyleAlb_Thrust_Tank(2)
        Call InitStyleAlb_Polearm(3)
        Call InitStyleAlb_TwoHanded(4)
        Call InitStyleAlb_Shield_Primary(5)
        Call InitStyleAll_Parry(6)
        Call InitStyleAlb_Crossbow(7)
        
    End With 'ttoon
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_CABALIST(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_CABALIST
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_MINION
        .REALM_ABILITIES(14) = RA_WILD_POWER
        
        .REALM_ABILITIES(15) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(16) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(17) = RA_CONCENTRATION
        .REALM_ABILITIES(18) = RA_FIRST_AID
        .REALM_ABILITIES(19) = RA_JUGGERNAUT
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(23) = RA_PURGE
        .REALM_ABILITIES(24) = RA_RAGING_POWER
        .REALM_ABILITIES(25) = RA_SECOND_WIND
        .REALM_ABILITIES(26) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(27) = RA_RESOLUTE_MINION
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellAlb_Matter(0)
        Call InitSpellAlb_Matter_Manipulation(1)
        Call InitSpellAlb_Body_Destruction(2)
        Call InitSpellAlb_Essence_Manipulation(3)
        Call InitSpellAlb_Spirit_Animation(4)
        Call InitSpellAlb_Vivification(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_CLERIC(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_CLERIC
        
        .pStat = SM_PIE 'pie
        .sStat = SM_CON 'con
        .tStat = SM_STR 'str
                    
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_PERFECTER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = vbNull
        .REALM_ABILITIES(9) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(10) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(11) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(13) = vbNull
        .REALM_ABILITIES(14) = RA_TOUGHNESS
        .REALM_ABILITIES(15) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(16) = RA_WILD_HEALING
        .REALM_ABILITIES(17) = RA_WILD_POWER
        
        .REALM_ABILITIES(18) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(19) = RA_BARRIER_OF_FORTITUDE
        .REALM_ABILITIES(20) = RA_DIVINE_INTERVENTION
        .REALM_ABILITIES(21) = RA_FIRST_AID
        .REALM_ABILITIES(22) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(23) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(24) = RA_PERFECT_RECOVERY
        .REALM_ABILITIES(25) = RA_PURGE
        .REALM_ABILITIES(26) = RA_RAGING_POWER
        .REALM_ABILITIES(27) = RA_SECOND_WIND
        .REALM_ABILITIES(28) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(29) = RA_RETRIBUTION_OF_THE_FAITHFUL
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellAlb_Rejuvenation(0)
        Call InitSpellAlb_Rejuvenation_Spec_Cleric(1)
        Call InitSpellAlb_Enhancement(2)
        Call InitSpellAlb_Enhancement_Spec_Cleric(3)
        Call InitSpellAlb_Smiting(4)
        Call InitSpellAlb_Smiting_Spec(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_FRIAR(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_FRIAR
        
        .pStat = SM_PIE 'pie
        .sStat = SM_CON 'con
        .tStat = SM_STR 'str
                        
        .MULTIPLIER = 1.8
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_PERFECTER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_HEALING
        
        .REALM_ABILITIES(15) = RA_FIRST_AID
        .REALM_ABILITIES(16) = RA_IGNORE_PAIN
        .REALM_ABILITIES(17) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(18) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(19) = RA_PURGE
        .REALM_ABILITIES(20) = RA_RAGING_POWER
        .REALM_ABILITIES(21) = RA_REFLEX_ATTACK
        .REALM_ABILITIES(22) = RA_SECOND_WIND
        .REALM_ABILITIES(23) = RA_STATIC_TEMPEST
        .REALM_ABILITIES(24) = RA_EMPTY_MIND
        .REALM_ABILITIES(25) = RA_VEHEMENT_RENEWAL
        
        .REALM_ABILITIES(26) = RA_WHIRLING_STAFF
        
        .SPELL_COUNT = 4
        .STYLE_COUNT = 2
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAlb_Staff_Friar(0)
        Call InitStyleAll_Parry(1)
        
        Call InitSpellAlb_Rejuvenation(0)
        Call InitSpellAlb_Rejuvenation_Spec_Friar(1)
        Call InitSpellAlb_Enhancement(2)
        Call InitSpellAlb_Enhancement_Spec_Friar(3)

    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_HERETIC(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_HERETIC
        
        .pStat = SM_PIE 'pie
        .sStat = SM_DEX 'dex
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BANELORD
        .ML_OPTION_2 = ML_PERFECTER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(9) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(10) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(11) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(13) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(14) = RA_TOUGHNESS
        .REALM_ABILITIES(15) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(16) = RA_WILD_HEALING
        .REALM_ABILITIES(17) = RA_WILD_POWER
        
        .REALM_ABILITIES(18) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(19) = RA_DIVINE_INTERVENTION
        .REALM_ABILITIES(20) = RA_FIRST_AID
        .REALM_ABILITIES(21) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(22) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(23) = RA_PERFECT_RECOVERY
        .REALM_ABILITIES(24) = RA_PURGE
        .REALM_ABILITIES(25) = RA_RAGING_POWER
        .REALM_ABILITIES(26) = RA_SECOND_WIND
        .REALM_ABILITIES(27) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(28) = RA_FANATICISM
        
        .SPELL_COUNT = 4
        .STYLE_COUNT = 3
        
        Call ClearStyles
        
        Call InitStyleAlb_Crush(0)
        Call InitStyleAlb_Flex_Heretic(1)
        Call InitStyleAll_Shield_Basic(2)
        
        Call InitSpellAlb_Rejuvenation(0)
        Call InitSpellAlb_Rejuvenation_Spec_Heretic(1)
        Call InitSpellAlb_Enhancement(2)
        Call InitSpellAlb_Enhancement_Spec_Heretic(3)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_INFILTRATOR(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_INFILTRATOR
        
        .pStat = SM_DEX 'dex
        .sStat = SM_QUI 'qui
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 2.5
                        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
        
        .ML_OPTION_1 = ML_SPYMASTER
        .ML_OPTION_2 = ML_BATTLEMASTER
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_LIFTER
        .REALM_ABILITIES(6) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(7) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(8) = RA_MASTERY_OF_STEALTH
        .REALM_ABILITIES(9) = RA_TOUGHNESS
        .REALM_ABILITIES(10) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(11) = RA_FIRST_AID
        .REALM_ABILITIES(12) = RA_PURGE
        .REALM_ABILITIES(13) = RA_SECOND_WIND
        .REALM_ABILITIES(14) = RA_EMPTY_MIND
        .REALM_ABILITIES(15) = RA_VANISH
        .REALM_ABILITIES(16) = RA_VIPER
        
        .REALM_ABILITIES(17) = RA_OVERWHELM
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 5
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAlb_Slash(0)
        Call InitStyleAlb_Thrust_Rogue(1)
        Call InitStyleAll_CriticalStrike(2)
        Call InitStyleAlb_Dualwield_Inf(3)
        Call InitStyleAll_Stealth(4)
        
        Call InitSpellAll_Mundane_Poisons(0)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_MERCENARY(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_MERCENARY
        
        .pStat = SM_STR 'str
        .sStat = SM_DEX 'dex
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 2
                    
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(8) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(10) = RA_TOUGHNESS
        .REALM_ABILITIES(11) = RA_VEIL_RECOVERY
       
        .REALM_ABILITIES(12) = RA_CHARGE
        .REALM_ABILITIES(13) = RA_FIRST_AID
        .REALM_ABILITIES(14) = RA_IGNORE_PAIN
        .REALM_ABILITIES(15) = RA_PURGE
        .REALM_ABILITIES(16) = RA_REFLEX_ATTACK
        .REALM_ABILITIES(17) = RA_SECOND_WIND
        .REALM_ABILITIES(18) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(19) = RA_BLINDING_DUST
        
        .SPELL_COUNT = 0
        .STYLE_COUNT = 6
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAlb_Crush(0)
        Call InitStyleAlb_Slash(1)
        Call InitStyleAlb_Thrust_Rogue(2)
        Call InitStyleAlb_Dualwield(3)
        Call InitStyleAll_Shield_Basic(4)
        Call InitStyleAll_Parry(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_MINSTREL(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_MINSTREL
        
        .pStat = SM_CHA 'cha
        .sStat = SM_DEX 'dex
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 1.5
                            
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_SOJOURNER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_MINION
        .REALM_ABILITIES(15) = RA_WILD_POWER
        
        .REALM_ABILITIES(16) = RA_AMELIORATING_MELODIES
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_IGNORE_PAIN
        .REALM_ABILITIES(19) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(20) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(21) = RA_PURGE
        .REALM_ABILITIES(22) = RA_RAGING_POWER
        .REALM_ABILITIES(23) = RA_SECOND_WIND
        .REALM_ABILITIES(24) = RA_SPEED_OF_SOUND
        .REALM_ABILITIES(25) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(26) = RA_CALMING_NOTES
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 3
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAlb_Slash(0)
        Call InitStyleAlb_Thrust_Rogue(1)
        Call InitStyleAll_Stealth(2)
        
        Call InitSpellAlb_Instruments(0)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_NECROMANCER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_NECROMANCER
    
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_STR
        .REALM_ABILITIES(3) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(4) = RA_LIFTER
        .REALM_ABILITIES(5) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(6) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(7) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(8) = RA_TOUGHNESS
        .REALM_ABILITIES(9) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(10) = RA_WILD_POWER
        
        .REALM_ABILITIES(11) = RA_CONCENTRATION
        .REALM_ABILITIES(12) = RA_FIRST_AID
        .REALM_ABILITIES(13) = RA_ICHOR_OF_THE_DEEP
        .REALM_ABILITIES(14) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(15) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(16) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(17) = RA_PURGE
        .REALM_ABILITIES(18) = RA_RAGING_POWER
        .REALM_ABILITIES(19) = RA_SECOND_WIND
        .REALM_ABILITIES(20) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(21) = RA_CALL_OF_DARKNESS
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellAlb_Deathsight(0)
        Call InitSpellAlb_Deathsight_Spec(1)
        Call InitSpellAlb_Painworking(2)
        Call InitSpellAlb_Painworking_Spec(3)
        Call InitSpellAlb_Death_Servant(4)
        Call InitSpellAlb_Death_Servant_Spec(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_PALADIN(ByRef tToon As TOON_TYPE)
    
    With tToon
    
        .CLASS = TCA_PALADIN
        
        .pStat = SM_CON 'con
        .sStat = SM_PIE 'pie
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 2#
                    
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 2
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_BATTLEMASTER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(13) = RA_TOUGHNESS
        .REALM_ABILITIES(14) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(15) = RA_WILD_POWER
        
        .REALM_ABILITIES(16) = RA_ANGER_OF_THE_GODS
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_IGNORE_PAIN
        .REALM_ABILITIES(19) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(20) = RA_PURGE
        .REALM_ABILITIES(21) = RA_RAGING_POWER
        .REALM_ABILITIES(22) = RA_SECOND_WIND
        .REALM_ABILITIES(23) = RA_EMPTY_MIND
        .REALM_ABILITIES(24) = RA_VEHEMENT_RENEWAL
        .REALM_ABILITIES(25) = RA_WRATH_OF_CHAMPIONS
        
        .REALM_ABILITIES(26) = RA_SELFLESS_DEVOTION
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 6
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAlb_Crush(0)
        Call InitStyleAlb_Slash(1)
        Call InitStyleAlb_Thrust_Tank(2)
        Call InitStyleAlb_TwoHanded(3)
        Call InitStyleAll_Shield_Basic(4)
        Call InitStyleAll_Parry(5)
        
        Call InitSpellAlb_Chants(0)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_REAVER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_REAVER
        
        .pStat = SM_STR 'str
        .sStat = SM_DEX 'dex
        .tStat = SM_PIE 'pie
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 2
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(11) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(13) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(14) = RA_TOUGHNESS
        .REALM_ABILITIES(15) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(16) = RA_WILD_POWER
        
        .REALM_ABILITIES(17) = RA_DUAL_THREAT
        .REALM_ABILITIES(18) = RA_FIRST_AID
        .REALM_ABILITIES(19) = RA_IGNORE_PAIN
        .REALM_ABILITIES(20) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(21) = RA_PURGE
        .REALM_ABILITIES(22) = RA_RAGING_POWER
        .REALM_ABILITIES(23) = RA_SECOND_WIND
        .REALM_ABILITIES(24) = RA_STRIKE_PREDICTION
        .REALM_ABILITIES(25) = RA_EMPTY_MIND
        .REALM_ABILITIES(26) = RA_THORNWEED_FIELD
        .REALM_ABILITIES(27) = RA_CHARGE
                
        .REALM_ABILITIES(28) = RA_SOUL_QUENCH
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 6
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAlb_Crush(0)
        Call InitStyleAlb_Slash(1)
        Call InitStyleAlb_Thrust_Rogue(2)
        Call InitStyleAlb_Flex_Reaver(3)
        Call InitStyleAll_Shield_Basic(4)
        Call InitStyleAll_Parry(5)
        
        Call InitSpellAlb_Soulrending(0)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_SCOUT(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCA_SCOUT
        
        .pStat = SM_DEX 'dex
        .sStat = SM_QUI 'qui
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_SOJOURNER
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_FALCONS_EYE
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_MASTERY_OF_STEALTH
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(13) = RA_FIRST_AID
        .REALM_ABILITIES(14) = RA_IGNORE_PAIN
        .REALM_ABILITIES(15) = RA_PURGE
        .REALM_ABILITIES(16) = RA_SECOND_WIND
        .REALM_ABILITIES(17) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(18) = RA_SHIELD_TRIP
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 4
        
        Call ClearSpells
        Call ClearStyles
               
        
        'Call InitStyleAlb_Longbow(0)
        Call InitStyleAlb_Slash(0)
        Call InitStyleAlb_Thrust_Rogue(1)
        Call InitStyleAll_Shield_Basic(2)
        Call InitStyleAll_Stealth(3)
        
        Call InitSpellAll_Archery(0)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_SORCERER(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCA_SORCERER
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_MINION
        .REALM_ABILITIES(14) = RA_WILD_POWER
        
        .REALM_ABILITIES(15) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(16) = RA_CONCENTRATION
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_ICHOR_OF_THE_DEEP
        .REALM_ABILITIES(19) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(20) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(21) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(22) = RA_PURGE
        .REALM_ABILITIES(23) = RA_RAGING_POWER
        .REALM_ABILITIES(24) = RA_SECOND_WIND
        .REALM_ABILITIES(25) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(26) = RA_SHIELD_OF_IMMUNITY
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellAlb_Matter(0)
        Call InitSpellAlb_Telekinesis(1)
        Call InitSpellAlb_Body_Destruction(2)
        Call InitSpellAlb_Disorientation(3)
        Call InitSpellAlb_Mind_Twisting(4)
        Call InitSpellAlb_Domination(5)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_THEURGIST(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCA_THEURGIST
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
    
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        
        .REALM_ABILITIES(14) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(15) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(16) = RA_CONCENTRATION
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(19) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(20) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(21) = RA_PURGE
        .REALM_ABILITIES(22) = RA_RAGING_POWER
        .REALM_ABILITIES(23) = RA_SECOND_WIND
        .REALM_ABILITIES(24) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(25) = RA_MINION_RESCUE
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellAlb_Path_Of_Earth(0)
        Call InitSpellAlb_Abrasion(1)
        Call InitSpellAlb_Path_Of_Ice(2)
        Call InitSpellAlb_Refrigeration(3)
        Call InitSpellAlb_Path_Of_Air(4)
        Call InitSpellAlb_Vapormancy(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_WIZARD(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCA_WIZARD
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        
        .REALM_ABILITIES(14) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(15) = RA_CONCENTRATION
        .REALM_ABILITIES(16) = RA_DECIMATION_TRAP
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(19) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(20) = RA_PURGE
        .REALM_ABILITIES(21) = RA_RAGING_POWER
        .REALM_ABILITIES(22) = RA_SECOND_WIND
        .REALM_ABILITIES(23) = RA_EMPTY_MIND
        .REALM_ABILITIES(24) = RA_VOLCANIC_PILLAR
        
        .REALM_ABILITIES(25) = RA_WALL_OF_FLAME
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellAlb_Path_Of_Earth(0)
        Call InitSpellAlb_Calefaction(1)
        Call InitSpellAlb_Path_Of_Ice(2)
        Call InitSpellAlb_Liquifaction(3)
        Call InitSpellAlb_Path_Of_Fire(4)
        Call InitSpellAlb_Pyromancy(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_ANIMIST(ByRef tToon As TOON_TYPE)
    
    With tToon
    
        .CLASS = TCH_ANIMIST
        
        .pStat = SM_INT 'int
        .sStat = SM_CON 'con
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        
        .REALM_ABILITIES(14) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(15) = RA_CONCENTRATION
        .REALM_ABILITIES(16) = RA_FIRST_AID
        .REALM_ABILITIES(17) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(18) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(19) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(20) = RA_PURGE
        .REALM_ABILITIES(21) = RA_RAGING_POWER
        .REALM_ABILITIES(22) = RA_SECOND_WIND
        .REALM_ABILITIES(23) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(24) = RA_FUNGAL_UNION
                
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellHib_Arboreal_Path(0)
        Call InitSpellHib_Arboreal_Mastery_Animist(1)
        Call InitSpellHib_Creeping_Path(2)
        Call InitSpellHib_Creeping_Mastery(3)
        Call InitSpellHib_Verdant_Path(4)
        Call InitSpellHib_Verdant_Mastery(5)
        
    End With 'ttoon
    
    Call SetBuffs(tToon)
        
End Sub

Public Sub Init_BAINSHEE(ByRef tToon As TOON_TYPE)
    
    With tToon
    
        .CLASS = TCH_BAINSHEE
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(10) = RA_TOUGHNESS
        .REALM_ABILITIES(11) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(12) = RA_WILD_POWER
        
        .REALM_ABILITIES(13) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(14) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(15) = RA_CONCENTRATION
        .REALM_ABILITIES(16) = RA_FIRST_AID
        .REALM_ABILITIES(17) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(18) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(19) = RA_PURGE
        .REALM_ABILITIES(20) = RA_RAGING_POWER
        .REALM_ABILITIES(21) = RA_SECOND_WIND
        .REALM_ABILITIES(22) = RA_STRIKE_PREDICTION
        .REALM_ABILITIES(23) = RA_EMPTY_MIND
        .REALM_ABILITIES(24) = RA_VOLCANIC_PILLAR
        
        .REALM_ABILITIES(25) = RA_SONIC_BARRIER
        
        .SPELL_COUNT = 4
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellHib_Spectral_Force(0)
        Call InitSpellHib_Spectral_Guard(1)
        Call InitSpellHib_Phantasmal_Wail(2)
        Call InitSpellHib_Ethereal_Shriek(3)
        
    End With 'ttoon
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_BARD(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCH_BARD
        
        .pStat = SM_CHA 'cha
        .sStat = SM_EMP 'emp
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
        
        .ML_OPTION_1 = ML_SOJOURNER
        .ML_OPTION_2 = ML_PERFECTER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_HEALING
        .REALM_ABILITIES(15) = RA_WILD_POWER
        
        .REALM_ABILITIES(16) = RA_AMELIORATING_MELODIES
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_IGNORE_PAIN
        .REALM_ABILITIES(19) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(20) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(21) = RA_PURGE
        .REALM_ABILITIES(22) = RA_RAGING_POWER
        .REALM_ABILITIES(23) = RA_SECOND_WIND
        .REALM_ABILITIES(24) = RA_SPEED_OF_SOUND
        .REALM_ABILITIES(25) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(26) = RA_DREAMWEAVER
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 2
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Blades(0)
        Call InitStyleHib_Blunt(1)
        
        Call InitSpellHib_Regrowth(0)
        Call InitSpellHib_Regrowth_Spec_Bard(1)
        Call InitSpellHib_Nurture(2)
        Call InitSpellHib_Nurture_Spec_Bard(3)
        Call InitSpellHib_Music(4)
        Call InitSpellHib_Music_Spec(5)
        
    End With 'ttoon
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_BLADEMASTER(ByRef tToon As TOON_TYPE)

    With TOON
    
        .CLASS = TCH_BLADEMASTER
        
        .pStat = SM_STR 'str
        .sStat = SM_DEX 'dex
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(8) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(10) = RA_TOUGHNESS
        .REALM_ABILITIES(11) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(12) = RA_WILD_POWER
       
        .REALM_ABILITIES(13) = RA_CHARGE
        .REALM_ABILITIES(14) = RA_FIRST_AID
        .REALM_ABILITIES(15) = RA_IGNORE_PAIN
        .REALM_ABILITIES(16) = RA_PURGE
        .REALM_ABILITIES(17) = RA_REFLEX_ATTACK
        .REALM_ABILITIES(18) = RA_SECOND_WIND
        .REALM_ABILITIES(19) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(20) = RA_BLADE_BARRIER
        
        .SPELL_COUNT = 0
        .STYLE_COUNT = 6
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Blades(0)
        Call InitStyleHib_Blunt(1)
        Call InitStyleHib_Pierce(2)
        Call InitStyleHib_CelticDual(3)
        Call InitStyleAll_Shield_Basic(4)
        Call InitStyleAll_Parry(5)
                
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_CHAMPION(ByRef tToon As TOON_TYPE)

    With tToon
        .CLASS = TCH_CHAMPION
        
        .pStat = SM_STR 'str
        .sStat = SM_INT 'int
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(13) = RA_TOUGHNESS
        .REALM_ABILITIES(14) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(15) = RA_WILD_POWER
        
        .REALM_ABILITIES(16) = RA_DUAL_THREAT
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_IGNORE_PAIN
        .REALM_ABILITIES(19) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(20) = RA_PURGE
        .REALM_ABILITIES(21) = RA_RAGING_POWER
        .REALM_ABILITIES(22) = RA_SECOND_WIND
        .REALM_ABILITIES(23) = RA_STATIC_TEMPEST
        .REALM_ABILITIES(24) = RA_EMPTY_MIND
        .REALM_ABILITIES(25) = RA_WRATH_OF_CHAMPIONS
        
        .REALM_ABILITIES(26) = RA_BADGE_OF_VALOR
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 6
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Blades(0)
        Call InitStyleHib_Blunt(1)
        Call InitStyleHib_Pierce(2)
        Call InitStyleHib_LargeWeaponry(3)
        Call InitStyleAll_Shield_Basic(4)
        Call InitStyleAll_Parry(5)
        
        Call InitSpellHib_Valor(0)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_DRUID(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCH_DRUID
        
        .pStat = SM_EMP 'emp
        .sStat = SM_CON 'con
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_PERFECTER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_HEALING
        .REALM_ABILITIES(15) = RA_WILD_MINION
        .REALM_ABILITIES(16) = RA_WILD_POWER
        
        .REALM_ABILITIES(17) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(18) = RA_BARRIER_OF_FORTITUDE
        .REALM_ABILITIES(19) = RA_DIVINE_INTERVENTION
        .REALM_ABILITIES(20) = RA_FIRST_AID
        .REALM_ABILITIES(21) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(22) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(23) = RA_PERFECT_RECOVERY
        .REALM_ABILITIES(24) = RA_PURGE
        .REALM_ABILITIES(25) = RA_RAGING_POWER
        .REALM_ABILITIES(26) = RA_SECOND_WIND
        .REALM_ABILITIES(27) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(28) = RA_NATURES_WOMB
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellHib_Regrowth(0)
        Call InitSpellHib_Regrowth_Spec_Druid(1)
        Call InitSpellHib_Nurture(2)
        Call InitSpellHib_Nurture_Spec_Druid(3)
        Call InitSpellHib_Nature(4)
        Call InitSpellHib_Nature_Spec(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_ELDRITCH(ByRef tToon As TOON_TYPE)
    
    With tToon
    
        .CLASS = TCH_ELDRITCH
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        .REALM_ABILITIES(14) = RA_LONGwind
        
        
        
        .REALM_ABILITIES(14) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(15) = RA_CONCENTRATION
        .REALM_ABILITIES(16) = RA_DECIMATION_TRAP
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(19) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(20) = RA_PURGE
        .REALM_ABILITIES(21) = RA_RAGING_POWER
        .REALM_ABILITIES(22) = RA_SECOND_WIND
        .REALM_ABILITIES(23) = RA_EMPTY_MIND
        .REALM_ABILITIES(24) = RA_VOLCANIC_PILLAR
        
        .REALM_ABILITIES(25) = RA_ARMS_LENGTH
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellHib_Way_Of_The_Sun(0)
        Call InitSpellHib_Shadow_Control(1)
        Call InitSpellHib_Way_Of_The_Moon(2)
        Call InitSpellHib_Vacuumancy(3)
        Call InitSpellHib_Way_Of_The_Eclipse(4)
        Call InitSpellHib_Void_Mastery(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_ENCHANTER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCH_ENCHANTER
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_MINION
        .REALM_ABILITIES(14) = RA_WILD_POWER
        
        .REALM_ABILITIES(15) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(16) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(17) = RA_CONCENTRATION
        .REALM_ABILITIES(18) = RA_FIRST_AID
        .REALM_ABILITIES(19) = RA_JUGGERNAUT
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(23) = RA_PURGE
        .REALM_ABILITIES(24) = RA_RAGING_POWER
        .REALM_ABILITIES(25) = RA_SECOND_WIND
        .REALM_ABILITIES(26) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(27) = RA_PROTECTION_OF_THE_UNDERHILL
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellHib_Way_Of_The_Sun(0)
        Call InitSpellHib_Bedazzling(1)
        Call InitSpellHib_Way_Of_The_Moon(2)
        Call InitSpellHib_Empowering(3)
        Call InitSpellHib_Enchantment(4)
        Call InitSpellHib_Enchantment_Mastery(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_HERO(ByRef tToon As TOON_TYPE)

    With tToon
        .CLASS = TCH_HERO
        
        .pStat = SM_STR 'str
        .sStat = SM_CON 'con
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_WARLORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(8) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(10) = RA_TOUGHNESS
        .REALM_ABILITIES(11) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(12) = RA_DASHING_DEFENSE
        .REALM_ABILITIES(13) = RA_FIRST_AID
        .REALM_ABILITIES(14) = RA_IGNORE_PAIN
        .REALM_ABILITIES(15) = RA_PURGE
        .REALM_ABILITIES(16) = RA_SECOND_WIND
        .REALM_ABILITIES(17) = RA_SOLDIERS_BARRICADE
        .REALM_ABILITIES(18) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(19) = RA_COMBAT_AWARENESS
        
        .SPELL_COUNT = 0
        .STYLE_COUNT = 7
    
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Blades(0)
        Call InitStyleHib_Blunt(1)
        Call InitStyleHib_Pierce(2)
        Call InitStyleHib_CelticSpear(3)
        Call InitStyleHib_LargeWeaponry(4)
        Call InitStyleHib_Shield_Primary(5)
        Call InitStyleAll_Parry(6)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_MENTALIST(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCH_MENTALIST
        
        .pStat = SM_INT 'int
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_STORMLORD
        .ML_OPTION_2 = ML_WARLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(9) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(10) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(11) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_HEALING
        .REALM_ABILITIES(15) = RA_WILD_MINION
        .REALM_ABILITIES(16) = RA_WILD_POWER
        
        .REALM_ABILITIES(17) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(18) = RA_CONCENTRATION
        .REALM_ABILITIES(19) = RA_FIRST_AID
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(23) = RA_PURGE
        .REALM_ABILITIES(24) = RA_RAGING_POWER
        .REALM_ABILITIES(25) = RA_SECOND_WIND
        .REALM_ABILITIES(26) = RA_STATIC_TEMPEST
        .REALM_ABILITIES(27) = RA_STRIKE_PREDICTION
        .REALM_ABILITIES(28) = RA_EMPTY_MIND
    
        .REALM_ABILITIES(29) = RA_SELECTIVE_BLINDNESS
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellHib_Way_Of_The_Sun(0)
        Call InitSpellHib_Illusions(1)
        Call InitSpellHib_Way_Of_The_Moon(2)
        Call InitSpellHib_Holism(3)
        Call InitSpellHib_Mentalism(4)
        Call InitSpellHib_Mind_Mastery(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_NIGHTSHADE(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCH_NIGHTSHADE
        
        .pStat = SM_DEX 'dex
        .sStat = SM_QUI 'qui
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 2.2
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
        
        .ML_OPTION_1 = ML_SPYMASTER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_MASTERY_OF_STEALTH
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(13) = RA_FIRST_AID
        .REALM_ABILITIES(14) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(15) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(16) = RA_PURGE
        .REALM_ABILITIES(17) = RA_RAGING_POWER
        .REALM_ABILITIES(18) = RA_SECOND_WIND
        .REALM_ABILITIES(19) = RA_EMPTY_MIND
        .REALM_ABILITIES(20) = RA_VANISH
        .REALM_ABILITIES(21) = RA_VIPER
        
        .REALM_ABILITIES(22) = RA_SHADOW_SHROUD
        
        .SPELL_COUNT = 2
        .STYLE_COUNT = 5
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Blades(0)
        Call InitStyleHib_Pierce(1)
        Call InitStyleHib_CelticDual(2)
        Call InitStyleAll_CriticalStrike(3)
        Call InitStyleAll_Stealth(4)
        
        Call InitSpellAll_Mundane_Poisons(0)
        Call InitSpellHib_Nightshade_Magic(1)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_RANGER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCH_RANGER
        
        .pStat = SM_DEX 'dex
        .sStat = SM_QUI 'qui
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
    
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_SOJOURNER
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_FALCONS_EYE
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_MASTERY_OF_STEALTH
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
                
        .REALM_ABILITIES(14) = RA_FIRST_AID
        .REALM_ABILITIES(15) = RA_IGNORE_PAIN
        .REALM_ABILITIES(16) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(17) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(18) = RA_PURGE
        .REALM_ABILITIES(19) = RA_RAGING_POWER
        .REALM_ABILITIES(20) = RA_SECOND_WIND
        .REALM_ABILITIES(21) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(22) = RA_DESPERATE_BOWMAN
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 4
        
        Call ClearSpells
        Call ClearStyles
        
               
        'Call InitStyleHib_Recurve(0)
        Call InitStyleHib_Blades(0)
        Call InitStyleHib_Pierce(1)
        Call InitStyleHib_CelticDual(2)
        Call InitStyleAll_Stealth(3)
        
        Call InitSpellAll_Archery(0)
        'Call InitSpellHib_Pathfinding(0)
        
    End With
    
    Call SetBuffs(tToon)
        
End Sub

Public Sub Init_VALEWALKER(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCH_VALEWALKER
        
        .pStat = SM_STR 'str
        .sStat = SM_INT 'int
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(13) = RA_TOUGHNESS
        .REALM_ABILITIES(14) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(15) = RA_WILD_POWER
        
        .REALM_ABILITIES(16) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_ICHOR_OF_THE_DEEP
        .REALM_ABILITIES(19) = RA_IGNORE_PAIN
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_PURGE
        .REALM_ABILITIES(23) = RA_RAGING_POWER
        .REALM_ABILITIES(24) = RA_SECOND_WIND
        .REALM_ABILITIES(25) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(26) = RA_VALE_DEFENSE
        
        .SPELL_COUNT = 2
        .STYLE_COUNT = 2
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Scythe(0)
        Call InitStyleAll_Parry(1)
        
        Call InitSpellHib_Arboreal_Path(0)
        Call InitSpellHib_Arboreal_Mastery_Valewalker(1)
                        
    End With
    
    Call SetBuffs(tToon)
        
End Sub

Public Sub Init_VAMPIIR(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCH_VAMPIIR
        
        .pStat = SM_CON 'con
        .sStat = SM_STR 'str
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BANELORD
        .ML_OPTION_2 = ML_WARLORD

        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_LIFTER
        .REALM_ABILITIES(6) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(7) = RA_TOUGHNESS
        .REALM_ABILITIES(8) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(9) = RA_CHARGE
        .REALM_ABILITIES(10) = RA_FIRST_AID
        .REALM_ABILITIES(11) = RA_IGNORE_PAIN
        .REALM_ABILITIES(12) = RA_PURGE
        .REALM_ABILITIES(13) = RA_SECOND_WIND
        .REALM_ABILITIES(14) = RA_STRIKE_PREDICTION
        .REALM_ABILITIES(15) = RA_EMPTY_MIND
        .REALM_ABILITIES(16) = RA_WRATH_OF_CHAMPIONS
        
        .REALM_ABILITIES(17) = RA_MARK_OF_PREY
        
        .SPELL_COUNT = 3
        .STYLE_COUNT = 1
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Pierce_Vampiir(0)
        
        Call InitSpellHib_Shadow_Mastery(0)
        Call InitSpellHib_Vampiiric_Embrace(1)
        Call InitSpellHib_Dementia(2)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_WARDEN(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCH_WARDEN
        
        .pStat = SM_EMP 'emp
        .sStat = SM_STR 'str
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 1.8
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_PERFECTER
    
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(11) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(13) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(14) = RA_TOUGHNESS
        .REALM_ABILITIES(15) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(16) = RA_WILD_HEALING
        .REALM_ABILITIES(17) = RA_WILD_POWER
        
        .REALM_ABILITIES(18) = RA_ANGER_OF_THE_GODS
        .REALM_ABILITIES(19) = RA_FIRST_AID
        .REALM_ABILITIES(20) = RA_IGNORE_PAIN
        .REALM_ABILITIES(21) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(22) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(23) = RA_PURGE
        .REALM_ABILITIES(24) = RA_RAGING_POWER
        .REALM_ABILITIES(25) = RA_SECOND_WIND
        .REALM_ABILITIES(26) = RA_EMPTY_MIND
        .REALM_ABILITIES(27) = RA_THORNWEED_FIELD
        .REALM_ABILITIES(28) = RA_VEHEMENT_RENEWAL
        
        .REALM_ABILITIES(29) = RA_FURY_OF_NATURE
        
        .SPELL_COUNT = 4
        .STYLE_COUNT = 4
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleHib_Blades(0)
        Call InitStyleHib_Blunt(1)
        Call InitStyleAll_Shield_Basic(2)
        Call InitStyleAll_Parry(3)
        
        Call InitSpellHib_Regrowth(0)
        Call InitSpellHib_Regrowth_Spec_Warden(1)
        Call InitSpellHib_Nurture(2)
        Call InitSpellHib_Nurture_Spec_Warden(3)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_BERSERKER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCM_BERSERKER
        
        .pStat = SM_STR 'str
        .sStat = SM_DEX 'dex
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(8) = RA_TOUGHNESS
        .REALM_ABILITIES(9) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(10) = RA_CHARGE
        .REALM_ABILITIES(11) = RA_FIRST_AID
        .REALM_ABILITIES(12) = RA_IGNORE_PAIN
        .REALM_ABILITIES(13) = RA_PURGE
        .REALM_ABILITIES(14) = RA_REFLEX_ATTACK
        .REALM_ABILITIES(15) = RA_SECOND_WIND
        .REALM_ABILITIES(16) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(17) = RA_FEROCIOUS_WILL
        
        .SPELL_COUNT = 0
        .STYLE_COUNT = 5
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Axe(0)
        Call InitStyleMid_Hammer(1)
        Call InitStyleMid_Sword(2)
        Call InitStyleMid_LeftAxe(3)
        Call InitStyleAll_Parry(4)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_BONEDANCER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCM_BONEDANCER
        
        .pStat = SM_PIE 'pie
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_MINION
        .REALM_ABILITIES(14) = RA_WILD_POWER
        
        .REALM_ABILITIES(15) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(16) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(17) = RA_CONCENTRATION
        .REALM_ABILITIES(18) = RA_FIRST_AID
        .REALM_ABILITIES(19) = RA_JUGGERNAUT
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(23) = RA_PURGE
        .REALM_ABILITIES(24) = RA_RAGING_POWER
        .REALM_ABILITIES(25) = RA_SECOND_WIND
        .REALM_ABILITIES(26) = RA_EMPTY_MIND
        .REALM_ABILITIES(27) = RA_THORNWEED_FIELD
        
        .REALM_ABILITIES(28) = RA_ALLURE_OF_DEATH
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellMid_Darkness(0)
        Call InitSpellMid_Bone_Mystics(1)
        Call InitSpellMid_Suppression(2)
        Call InitSpellMid_Bone_Guardians(3)
        Call InitSpellMid_Bone_Army(4)
        Call InitSpellMid_Bone_Warriors(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_HEALER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCM_HEALER
        
        .pStat = SM_PIE 'pie
        .sStat = SM_CON 'con
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_SOJOURNER
        .ML_OPTION_2 = ML_PERFECTER
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_HEALING
        .REALM_ABILITIES(15) = RA_WILD_POWER
        
        .REALM_ABILITIES(16) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(17) = RA_BARRIER_OF_FORTITUDE
        .REALM_ABILITIES(18) = RA_DIVINE_INTERVENTION
        .REALM_ABILITIES(19) = RA_FIRST_AID
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_PERFECT_RECOVERY
        .REALM_ABILITIES(23) = RA_PURGE
        .REALM_ABILITIES(24) = RA_RAGING_POWER
        .REALM_ABILITIES(25) = RA_SECOND_WIND
        .REALM_ABILITIES(26) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(27) = RA_SPUTINS_LEGACY
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellMid_Mending(0)
        Call InitSpellMid_Mending_Spec_Healer(1)
        Call InitSpellMid_Augmentation(2)
        Call InitSpellMid_Augmentation_Spec_Healer(3)
        Call InitSpellMid_Pacification(4)
        Call InitSpellMid_Pacification_Spec(5)
        
    End With
        
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_HUNTER(ByRef tToon As TOON_TYPE)

    With TOON
    
        .CLASS = TCM_HUNTER
        
        .pStat = SM_DEX 'dex
        .sStat = SM_QUI 'qui
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
        
        .ML_OPTION_1 = ML_SOJOURNER
        .ML_OPTION_2 = ML_BATTLEMASTER
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_FALCONS_EYE
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_MASTERY_OF_STEALTH
        .REALM_ABILITIES(13) = RA_TOUGHNESS
        .REALM_ABILITIES(14) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(15) = RA_WILD_MINION
        .REALM_ABILITIES(16) = RA_WILD_POWER
        
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_IGNORE_PAIN
        .REALM_ABILITIES(19) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(20) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(21) = RA_PURGE
        .REALM_ABILITIES(22) = RA_RAGING_POWER
        .REALM_ABILITIES(23) = RA_SECOND_WIND
        .REALM_ABILITIES(24) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(25) = RA_ENTWINING_SNAKES
        
        .SPELL_COUNT = 2
        .STYLE_COUNT = 3
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Sword(0)
        Call InitStyleMid_Spear(1)
        Call InitStyleAll_Stealth(2)
        
        Call InitSpellAll_Archery(0)
        Call InitSpellMid_Beastcraft(1)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_RUNEMASTER(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCM_RUNEMASTER
        
        .pStat = SM_PIE 'pie
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        
        .REALM_ABILITIES(14) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(15) = RA_CONCENTRATION
        .REALM_ABILITIES(16) = RA_DECIMATION_TRAP
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(19) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(20) = RA_PURGE
        .REALM_ABILITIES(21) = RA_RAGING_POWER
        .REALM_ABILITIES(22) = RA_SECOND_WIND
        .REALM_ABILITIES(23) = RA_EMPTY_MIND
        .REALM_ABILITIES(24) = RA_VOLCANIC_PILLAR
        
        .REALM_ABILITIES(25) = RA_RUNE_OF_UTTER_AGILITY
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellMid_Darkness(0)
        Call InitSpellMid_Runes_Of_Darkness(1)
        Call InitSpellMid_Suppression(2)
        Call InitSpellMid_Runes_Of_Suppression(3)
        Call InitSpellMid_Runecarving(4)
        Call InitSpellMid_Runes_Of_Destruction(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_SAVAGE(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCM_SAVAGE
        
        .pStat = SM_DEX 'dex
        .sStat = SM_QUI 'qui
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_BATTLEMASTER
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        
        .REALM_ABILITIES(14) = RA_FIRST_AID
        .REALM_ABILITIES(15) = RA_IGNORE_PAIN
        .REALM_ABILITIES(16) = RA_PURGE
        .REALM_ABILITIES(17) = RA_SECOND_WIND
        .REALM_ABILITIES(18) = RA_STRIKE_PREDICTION
        .REALM_ABILITIES(19) = RA_EMPTY_MIND
        .REALM_ABILITIES(20) = RA_CHARGE
        
        .REALM_ABILITIES(21) = RA_FUELED_BY_RAGE
                
        .SPELL_COUNT = 1
        .STYLE_COUNT = 5
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Axe(0)
        Call InitStyleMid_Hammer(1)
        Call InitStyleMid_Sword(2)
        Call InitStyleMid_HandToHand(3)
        Call InitStyleAll_Parry(4)
        
        Call InitSpellMid_Savagery(0)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_SHADOWBLADE(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCM_SHADOWBLADE
        
        .pStat = SM_DEX 'dex
        .sStat = SM_QUI 'qui
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 2.2
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 1
        
        .ML_OPTION_1 = ML_SPYMASTER
        .ML_OPTION_2 = ML_BATTLEMASTER

        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_LIFTER
        .REALM_ABILITIES(6) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(7) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(8) = RA_MASTERY_OF_STEALTH
        .REALM_ABILITIES(9) = RA_TOUGHNESS
        .REALM_ABILITIES(10) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(11) = RA_FIRST_AID
        .REALM_ABILITIES(12) = RA_PURGE
        .REALM_ABILITIES(13) = RA_SECOND_WIND
        .REALM_ABILITIES(14) = RA_EMPTY_MIND
        .REALM_ABILITIES(15) = RA_VANISH
        .REALM_ABILITIES(16) = RA_VIPER
        
        .REALM_ABILITIES(17) = RA_BLOOD_DRINKING
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 5
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Axe(0)
        Call InitStyleMid_Sword(1)
        Call InitStyleMid_LeftAxe_SB(2)
        Call InitStyleAll_CriticalStrike(3)
        Call InitStyleAll_Stealth(4)
        
        Call InitSpellAll_Mundane_Poisons(0)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_SHAMAN(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCM_SHAMAN
        
        .pStat = SM_PIE 'pie
        .sStat = SM_CON 'con
        .tStat = SM_STR 'str
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_PERFECTER

        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_HEALING
        .REALM_ABILITIES(15) = RA_WILD_POWER
        
        .REALM_ABILITIES(16) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_ICHOR_OF_THE_DEEP
        .REALM_ABILITIES(19) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(20) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(21) = RA_PURGE
        .REALM_ABILITIES(22) = RA_RAGING_POWER
        .REALM_ABILITIES(23) = RA_SECOND_WIND
        .REALM_ABILITIES(24) = RA_EMPTY_MIND
        .REALM_ABILITIES(25) = RA_VEHEMENT_RENEWAL
        
        .REALM_ABILITIES(26) = RA_RESTORATIVE_MIND
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellMid_Mending(0)
        Call InitSpellMid_Mending_Spec_Shaman(1)
        Call InitSpellMid_Augmentation(2)
        Call InitSpellMid_Augmentation_Spec_Shaman(3)
        Call InitSpellMid_Cave_Magic(4)
        Call InitSpellMid_Cave_Magic_Spec(5)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_SKALD(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCM_SKALD
        
        .pStat = SM_CHA 'cha
        .sStat = SM_STR 'str
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_SOJOURNER
    
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(10) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(11) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(12) = RA_TOUGHNESS
        .REALM_ABILITIES(13) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(14) = RA_WILD_POWER
        
        .REALM_ABILITIES(15) = RA_AMELIORATING_MELODIES
        .REALM_ABILITIES(16) = RA_ANGER_OF_THE_GODS
        .REALM_ABILITIES(17) = RA_FIRST_AID
        .REALM_ABILITIES(18) = RA_IGNORE_PAIN
        .REALM_ABILITIES(19) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(20) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(21) = RA_PURGE
        .REALM_ABILITIES(22) = RA_RAGING_POWER
        .REALM_ABILITIES(23) = RA_SECOND_WIND
        .REALM_ABILITIES(24) = RA_SPEED_OF_SOUND
        .REALM_ABILITIES(25) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(26) = RA_VOICE_SKADI
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 4
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Axe(0)
        Call InitStyleMid_Hammer(1)
        Call InitStyleMid_Sword(2)
        Call InitStyleAll_Parry(3)
        
        Call InitSpellMid_Battlesongs(0)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_SPIRITMASTER(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCM_SPIRITMASTER
        
        .pStat = SM_PIE 'pie
        .sStat = SM_DEX 'dex
        .tStat = SM_QUI 'qui
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_CONVOKER
        .ML_OPTION_2 = ML_STORMLORD
         
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(10) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_MINION
        .REALM_ABILITIES(14) = RA_WILD_POWER
        
        .REALM_ABILITIES(15) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(16) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(17) = RA_CONCENTRATION
        .REALM_ABILITIES(18) = RA_FIRST_AID
        .REALM_ABILITIES(19) = RA_JUGGERNAUT
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_NEGATIVE_MAELSTROM
        .REALM_ABILITIES(23) = RA_PURGE
        .REALM_ABILITIES(24) = RA_RAGING_POWER
        .REALM_ABILITIES(25) = RA_SECOND_WIND
        .REALM_ABILITIES(26) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(27) = RA_SPIRIT_MARTYR
        
        .SPELL_COUNT = 6
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellMid_Darkness(0)
        Call InitSpellMid_Spirit_Dimming(1)
        Call InitSpellMid_Suppression(2)
        Call InitSpellMid_Spirit_Suppression(3)
        Call InitSpellMid_Summoning(4)
        Call InitSpellMid_Spirit_Enhancement(5)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_THANE(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCM_THANE
        
        .pStat = SM_STR 'str
        .sStat = SM_PIE 'pie
        .tStat = SM_CON 'con
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BATTLEMASTER
        .ML_OPTION_2 = ML_STORMLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(9) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(11) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(12) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(13) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(14) = RA_TOUGHNESS
        .REALM_ABILITIES(15) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(16) = RA_WILD_POWER
        
        .REALM_ABILITIES(17) = RA_DUAL_THREAT
        .REALM_ABILITIES(18) = RA_FIRST_AID
        .REALM_ABILITIES(19) = RA_IGNORE_PAIN
        .REALM_ABILITIES(20) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(21) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(22) = RA_PURGE
        .REALM_ABILITIES(23) = RA_RAGING_POWER
        .REALM_ABILITIES(24) = RA_SECOND_WIND
        .REALM_ABILITIES(25) = RA_STATIC_TEMPEST
        .REALM_ABILITIES(26) = RA_EMPTY_MIND
        .REALM_ABILITIES(27) = RA_WRATH_OF_CHAMPIONS
        
        .REALM_ABILITIES(28) = RA_CHAIN_LIGHTNING
        
        .SPELL_COUNT = 1
        .STYLE_COUNT = 5
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Axe_Thane(0)
        Call InitStyleMid_Hammer_Thane(1)
        Call InitStyleMid_Sword_Thane(2)
        Call InitStyleAll_Shield_Basic(3)
        Call InitStyleAll_Parry(4)
        
        Call InitSpellMid_Stormcalling(0)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_VALKYRIE(ByRef tToon As TOON_TYPE)

    With tToon
    
        .CLASS = TCM_VALKYRIE
        
        .pStat = SM_CON 'con
        .sStat = SM_STR 'str
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_STORMLORD
        .ML_OPTION_2 = ML_WARLORD
        
        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(6) = RA_DETERMINATION
        .REALM_ABILITIES(7) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(8) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(10) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(11) = RA_MASTERY_OF_HEALING
        .REALM_ABILITIES(12) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(13) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(14) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(15) = RA_TOUGHNESS
        .REALM_ABILITIES(16) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(17) = RA_WILD_HEALING
        .REALM_ABILITIES(18) = RA_WILD_POWER
        
        .REALM_ABILITIES(19) = RA_CHARGE
        .REALM_ABILITIES(20) = RA_DUAL_THREAT
        .REALM_ABILITIES(21) = RA_FIRST_AID
        .REALM_ABILITIES(22) = RA_ICHOR_OF_THE_DEEP
        .REALM_ABILITIES(23) = RA_IGNORE_PAIN
        .REALM_ABILITIES(24) = RA_MASTERY_OF_CONCENTRATION
        .REALM_ABILITIES(25) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(26) = RA_PURGE
        .REALM_ABILITIES(27) = RA_RAGING_POWER
        .REALM_ABILITIES(28) = RA_SECOND_WIND
        .REALM_ABILITIES(29) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(30) = RA_VALHALLAS_BLESSING
        
        .SPELL_COUNT = 3
        .STYLE_COUNT = 4
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Shield_Valkyrie(0)
        Call InitStyleMid_Spear_Valkyrie(1)
        Call InitStyleMid_Sword_Valkyrie(2)
        Call InitStyleAll_Parry(3)
        
        Call InitSpellMid_Odin(0)
        Call InitSpellMid_Mending(1)
        Call InitSpellMid_Mending_Spec_Valkyrie(2)
        
    End With
    
    Call SetBuffs(tToon)

End Sub

Public Sub Init_WARLOCK(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCM_WARLOCK
        
        .pStat = SM_PIE 'pie
        .sStat = SM_CON 'con
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 1#
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_BANELORD
        .ML_OPTION_2 = ML_CONVOKER

        .REALM_ABILITIES(0) = RA_AUG_ACU
        .REALM_ABILITIES(1) = RA_AUG_CON
        .REALM_ABILITIES(2) = RA_AUG_DEX
        .REALM_ABILITIES(3) = RA_AUG_QUI
        .REALM_ABILITIES(4) = RA_AUG_STR
        .REALM_ABILITIES(5) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(8) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(9) = RA_PHYSICAL_DEFENSE
        .REALM_ABILITIES(10) = RA_TOUGHNESS
        .REALM_ABILITIES(11) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(12) = RA_WILD_POWER
        
        .REALM_ABILITIES(13) = RA_ADRENALINE_RUSH
        .REALM_ABILITIES(14) = RA_BEDAZZLING_AURA
        .REALM_ABILITIES(15) = RA_DECIMATION_TRAP
        .REALM_ABILITIES(16) = RA_FIRST_AID
        .REALM_ABILITIES(17) = RA_MYSTIC_CRYSTAL_LORE
        .REALM_ABILITIES(18) = RA_PURGE
        .REALM_ABILITIES(19) = RA_RAGING_POWER
        .REALM_ABILITIES(20) = RA_SECOND_WIND
        .REALM_ABILITIES(21) = RA_EMPTY_MIND
        .REALM_ABILITIES(22) = RA_VOLCANIC_PILLAR
        
        .REALM_ABILITIES(23) = RA_BOILING_CAULDRON
        
        .SPELL_COUNT = 4
        .STYLE_COUNT = 0
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitSpellMid_Cursing(0)
        Call InitSpellMid_Cursing_Spec(1)
        Call InitSpellMid_Hexing(2)
        Call InitSpellMid_Witchcraft(3)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_WARRIOR(ByRef tToon As TOON_TYPE)

    With tToon
        
        .CLASS = TCM_WARRIOR
        
        .pStat = SM_STR 'str
        .sStat = SM_CON 'con
        .tStat = SM_DEX 'dex
        
        .MULTIPLIER = 2#
        
        .AUTO_TRAIN = True
        .AUTO_TRAIN_LINES = 3
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_BATTLEMASTER
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_LIFTER
        .REALM_ABILITIES(7) = RA_MASTERY_OF_BLOCKING
        .REALM_ABILITIES(8) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(9) = RA_MASTERY_OF_PARRYING
        .REALM_ABILITIES(10) = RA_TOUGHNESS
        .REALM_ABILITIES(11) = RA_VEIL_RECOVERY
        
        .REALM_ABILITIES(12) = RA_DASHING_DEFENSE
        .REALM_ABILITIES(13) = RA_FIRST_AID
        .REALM_ABILITIES(14) = RA_IGNORE_PAIN
        .REALM_ABILITIES(15) = RA_PURGE
        .REALM_ABILITIES(16) = RA_SECOND_WIND
        .REALM_ABILITIES(17) = RA_SOLDIERS_BARRICADE
        .REALM_ABILITIES(18) = RA_EMPTY_MIND
        
        .REALM_ABILITIES(19) = RA_TESTUDO
        
        .SPELL_COUNT = 0
        .STYLE_COUNT = 5
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleMid_Axe(0)
        Call InitStyleMid_Hammer(1)
        Call InitStyleMid_Sword(2)
        Call InitStyleAlb_Shield_Primary(3) 'arms and warrior use same shield styles
        Call InitStyleAll_Parry(4)
        
    End With
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_MAULER_ALB(ByRef tToon As TOON_TYPE)

    With tToon
        .CLASS = TCA_MAULER
        
        .pStat = SM_STR
        .sStat = SM_CON
        .tStat = SM_QUI
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(10) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        .REALM_ABILITIES(14) = RA_DUAL_THREAT
        .REALM_ABILITIES(15) = RA_FIRST_AID
        .REALM_ABILITIES(16) = RA_IGNORE_PAIN
        .REALM_ABILITIES(17) = RA_PURGE
        .REALM_ABILITIES(18) = RA_REFLEX_ATTACK
        .REALM_ABILITIES(19) = RA_SECOND_WIND
        .REALM_ABILITIES(20) = RA_EMPTY_MIND
        .REALM_ABILITIES(21) = RA_THORNWEED_FIELD
        
        .REALM_ABILITIES(22) = RA_GIFT_OF_PERIZOR
        
        .SPELL_COUNT = 3
        .STYLE_COUNT = 2
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAll_Fist_Mauler(0)
        Call InitStyleAll_Staff_Mauler(1)
        
        Call InitSpellAll_Mauler_Aura_Manipulation(0)
        Call InitSpellAll_Mauler_Magnetism(1)
        Call InitSpellAll_Mauler_Powerstrikes(2)
        
    End With 'ttoon
        
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_MAULER_HIB(ByRef tToon As TOON_TYPE)

    With tToon
        .CLASS = TCA_MAULER
        
        .pStat = SM_STR
        .sStat = SM_CON
        .tStat = SM_QUI
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(10) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        .REALM_ABILITIES(14) = RA_DUAL_THREAT
        .REALM_ABILITIES(15) = RA_FIRST_AID
        .REALM_ABILITIES(16) = RA_IGNORE_PAIN
        .REALM_ABILITIES(17) = RA_PURGE
        .REALM_ABILITIES(18) = RA_REFLEX_ATTACK
        .REALM_ABILITIES(19) = RA_SECOND_WIND
        .REALM_ABILITIES(20) = RA_EMPTY_MIND
        .REALM_ABILITIES(21) = RA_THORNWEED_FIELD
        
        .REALM_ABILITIES(22) = RA_GIFT_OF_PERIZOR
        
        .SPELL_COUNT = 3
        .STYLE_COUNT = 2
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAll_Fist_Mauler(0)
        Call InitStyleAll_Staff_Mauler(1)
        
        Call InitSpellAll_Mauler_Aura_Manipulation(0)
        Call InitSpellAll_Mauler_Magnetism(1)
        Call InitSpellAll_Mauler_Powerstrikes(2)
        
    End With 'ttoon
    
    Call SetBuffs(tToon)
    
End Sub

Public Sub Init_MAULER_MID(ByRef tToon As TOON_TYPE)

    With tToon
        .CLASS = TCM_MAULER
        
        .pStat = SM_STR
        .sStat = SM_CON
        .tStat = SM_QUI
        
        .MULTIPLIER = 1.5
        
        .AUTO_TRAIN = False
        .AUTO_TRAIN_LINES = 0
        
        .ML_OPTION_1 = ML_WARLORD
        .ML_OPTION_2 = ML_BANELORD
        
        .REALM_ABILITIES(0) = RA_AUG_CON
        .REALM_ABILITIES(1) = RA_AUG_DEX
        .REALM_ABILITIES(2) = RA_AUG_QUI
        .REALM_ABILITIES(3) = RA_AUG_STR
        .REALM_ABILITIES(4) = RA_AVOIDANCE_OF_MAGIC
        .REALM_ABILITIES(5) = RA_DETERMINATION
        .REALM_ABILITIES(6) = RA_ETHEREAL_BOND
        .REALM_ABILITIES(7) = RA_LIFTER
        .REALM_ABILITIES(8) = RA_MASTERY_OF_FOCUS
        .REALM_ABILITIES(9) = RA_MASTERY_OF_MAGERY
        .REALM_ABILITIES(10) = RA_MASTERY_OF_PAIN
        .REALM_ABILITIES(11) = RA_TOUGHNESS
        .REALM_ABILITIES(12) = RA_VEIL_RECOVERY
        .REALM_ABILITIES(13) = RA_WILD_POWER
        .REALM_ABILITIES(14) = RA_DUAL_THREAT
        .REALM_ABILITIES(15) = RA_FIRST_AID
        .REALM_ABILITIES(16) = RA_IGNORE_PAIN
        .REALM_ABILITIES(17) = RA_PURGE
        .REALM_ABILITIES(18) = RA_REFLEX_ATTACK
        .REALM_ABILITIES(19) = RA_SECOND_WIND
        .REALM_ABILITIES(20) = RA_EMPTY_MIND
        .REALM_ABILITIES(21) = RA_THORNWEED_FIELD
        
        .REALM_ABILITIES(22) = RA_GIFT_OF_PERIZOR
        
        .SPELL_COUNT = 3
        .STYLE_COUNT = 2
        
        Call ClearSpells
        Call ClearStyles
        
        Call InitStyleAll_Fist_Mauler(0)
        Call InitStyleAll_Staff_Mauler(1)
        
        Call InitSpellAll_Mauler_Aura_Manipulation(0)
        Call InitSpellAll_Mauler_Magnetism(1)
        Call InitSpellAll_Mauler_Powerstrikes(2)
        
    End With 'ttoon
    
    Call SetBuffs(tToon)
    
End Sub
