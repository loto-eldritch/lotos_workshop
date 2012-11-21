Attribute VB_Name = "CharCon"
Option Explicit
'--------------SPELL STUFF----------------------------------------------------------------------------------------------------------------------------
Public Type SPELL_TYPE
    SPELL_TYPE_NAME As String           'example:lifedrain
    SPELL_TYPE_DESC As String           'example: kills your enemies like crispy
    SPELL_DETAILS(20, 7) As String      'spell matrix (oh joy)
                                        'spell_details(m,n)
                                        'm = levels of spell
                                        'n = spell stats (level,name,cost,effect,cast time, recast time,
                                        '                    duration, damage type, ...)
End Type

Public Type SPELL_SCHOOL
    AUTO_TRAIN As Boolean
    AUTO_TRAIN_CLASS(10) As Long
    SPELL_PATH_NAME As String           'example: arboreal path
    SPELL_PATH_CATEGORY As Long         'example: SPELL_PATH_SPEC, SPELL_PATH_BASELINE
    SPELL_CLASS(20) As SPELL_TYPE       'example: shield, lifedrain, damage shield, etc...
End Type

'--------------STYLE STUFF----------------------------------------------------------------------------------------------------------------------------

Public Type STYLE_TYPE
    STYLE_DETAILS(15) As String         'style array (oh joy)
                                        'n = style stats (level,name,cost,effect..etc
End Type

Public Type STYLE_SCHOOL
    AUTO_TRAIN As Boolean
    AUTO_TRAIN_CLASS(10) As Long
    STYLE_PATH_NAME As String           'example: piercing
    STYLE_NO_STYLES As Boolean          'for the case of Parry and Stealth (not quite magic, not quite melee)
    STYLE(22) As STYLE_TYPE             'example: diamond-back
End Type
'-----------------------------------------------------------------------------------------------------------------------------------------------------

Public Type CLASS_ARMOR
    ALLOWED_CLASS(10) As Long
    ARMOR_TYPE(3) As String
End Type

Public ARMORS(2) As CLASS_ARMOR

Public SPELLS(10) As SPELL_SCHOOL
Public STYLES(10) As STYLE_SCHOOL

Public Type TOON_TYPE

    REALM As Long           '0 albion 1 hibernia 2 midgard
    RACE As Long
    GENDER As Long          '0 male, 1 female
    CLASS As Long
    
    CATEGORY As Long        'one of 3 values toon_melee toon_magic or toon_hybrid
    
    ML_OPTION_1 As Long
    ML_OPTION_2 As Long
    
    Name As String * 50
    
    LEVEL As Single             'toon level
    MULTIPLIER As Single        'level multiplier
    
    AUTO_TRAIN As Boolean       'character autotrains true/false
    AUTO_TRAIN_LINES As Long    'how many lines the character auto trains 1-3
    
    CREATION_POINTS As Long     'creation points (initial 30)
    
    
    '# STAT_MATRIX (I,J)
    '# I=Stat | J=Location
    '# I VALUE Index
    '# ---------------------------------------------------------------------------
    '# attributes:
    '# 1 str, 2 con, 3 dex, 4 qui, 5 int, 6 emp, 7 pie, 8 cha, 9 acuity
    '# 10 power, 11 hits
    '# ---------------------------------------------------------------------------
    '# resists:
    '# 12 crush, 13 slash, 14 thrust, 15 body, 16 cold, 17 energy
    '# 18 heat, 19 matter, 20 spirit
    '# ---------------------------------------------------------------------------
    '# melee skills: hib:
    '# 21 blades, 22 blunt, 23 pierce, 24 large-weapon, 25 celtic-spear
    '# 26 scythe, 27 celtic-dual, 28 recurve bow
    '# ---------------------------------------------------------------------------
    '# magic skills: hib:
    '# 29 arboreal, 30 creeping, 31 verdant, 32 light, 33 mana, 34 void
    '# 35 mentalism, 36 enchantments, 37 dementia, 38 vampiiric embrace
    '# 39 shadow mastery, 40 phantasmal wail, 41 spectral guard
    '# 42 ethereal shriek, 43 nurture, 44 nature, 45 regrowth, 46 music
    '# 47 valor
    '# ---------------------------------------------------------------------------
    '# melee skills: alb:
    '# 48 crush, 49 slash, 50 thrust, 51 dual-wield, 52 crossbow
    '# 53 polearm, 54 two-handed, 55 staff, 56 flexible, 57 longbow
    '# ---------------------------------------------------------------------------
    '# magic skills: alb:
    '# 58 cold, 59 earth, 60 fire, 61 wind, 62 matter, 63 mind, 64 body
    '# 65 spirit, 66 soulrending, 67 death-servant, 68 deathsight
    '# 69 painworking, 70 instruments, 71 enhancement, 72 rejuvenation
    '# 73 smite, 74 chants
    '# ---------------------------------------------------------------------------
    '# melee skills: mid:
    '# 75 hammer, 76 axe, 77 sword, 78 spear, 79 hand-to-hand
    '# 80 composite-bow, 81 thrown, 82 left-axe
    '# ---------------------------------------------------------------------------
    '# magic skills: mid:
    '# 83 bone-army, 84 darkness, 85 suppression, 86 mending
    '# 87 augmentation, 88 beastcraft, 89 runecarving, 90 cave-magic
    '# 91 battlesongs, 92 summoning, 93 stormcalling, 94 odin's-will
    '# 95 cursing, 96 hexing, 97 pacification, 98 witchcraft
    '# ---------------------------------------------------------------------------
    '# common skills: hib/alb/mid
    '# 99 stealth, 100 critical-strike, 101 envenom, 102 shield, 103 parry
    '# ---------------------------------------------------------------------------
    '# cap increases:
    '# 104 str-cap, 105 con-cap, 106 dex-cap, 107 qui-cap, 108 int-cap
    '# 109 emp-cap, 110 pie-cap, 111 cha-cap, 112 hits-cap, 113 pow-cap
    '# 114 acuity-cap
    '# ---------------------------------------------------------------------------
    '# toa bonuses
    '# 115 +af, 116 %power pool, 117 archery damage, 118 archery range
    '# 119 archery speed, 120 fatigue, 121 healing effectiveness
    '# 122 melee combat speed, 123 melee damage, 124 melee style damage
    '# 125 spell damage, 126 spell duration, 127 spell pierce, 128 spell haste
    '# 129 spell range, 130 stat buff effectiveness, 131 stat debuff effectiveness
    '# 132 toa-unique
    '# ---------------------------------------------------------------------------
    '# pve bonuses
    '# 133 arrow recovery, 134 bladeturn reinforcement, 135 block
    '# 136 concentration, 137 damage reduction
    '# 138 death experience loss reduction, 139 defensive, 140 evade
    '# 141 negative effect duration, 142 parry, 143 piece ablative
    '# 144 reactionary style damage, 145 spell power cost reduction
    '# 146 style cost reduction, 147 to-hit, 148 pve-unique
    '# ---------------------------------------------------------------------------
    '# focus hib
    '# 149 all focus, 150 arboreal focus, 151 creeping focus, 152 verdant focus
    '# 153 ethereal shriek focus, 154 phantasmal wail focus, 155 spectral guard focus
    '# 156 light focus, 157 mana focus, 158 void focus
    '# 159 enchantment focus, 160 mentalism focus
    '# focus alb
    '# 161 body magic focus, 162 matter magic focus, 163 spirit magic focus
    '# 164 deathsight focus, 165 death servant focus, 166 painworking focus
    '# 167 mind magic focus, 168 cold magic focus, 169 earth magic focus
    '# 170 wind magic focus, 171 fire magic focus
    '# focus mid
    '# 172 bone army focus, 173 darkness focus, 174 suppression focus
    '# 175 runecarving focus, 176 summoning focus, 177 cursing focus
    '# 178 hexing focus, 179 witchcraft focus
    '# ---------------------------------------------------------------------------
    '# 180 all melee skills, 181 all magic skills, 182 all dual wield skills
    '# 183 all archery skills
    '# ---------------------------------------------------------------------------
    '# J VALUE Index
    '# ---------------------------------------------------------------------------
    '# core toon:
    '# 0 reserved for skill totals, 1 base, 2 creation, 3 level
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
    '# buffs and realm ability locations:
    '# 22 total buffs, 23 realm ability, 24 vampiiir reserved, 25 mythical item
    '# ---------------------------------------------------------------------------
    '# 5th slot bonuses:
    '# 30 head, 31 chest, 32 arms, 33 gloves, 34 pants, 35 boots
    
    
      
    STAT_MATRIX(300, 40) As Long
    
    pStat As Long               'primary stat
    sStat As Long               'secondary stat
    tStat As Long               'tertiary stat
    
    STR As Long
    CON As Long
    DEX As Long
    QUI As Long
    INT As Long
    PIE As Long
    EMP As Long
    CHA As Long
    
    BODY As Long
    COLD As Long
    HEAT As Long
    ENERGY As Long
    MATTER As Long
    SPIRIT As Long
    CRUSH As Long
    SLASH As Long
    THRUST As Long
    ESSENCE As Long
    
    HITS As Long
    POWER As Long
    
    REALM_RANK As Long              'realm rank
    REALM_LEVEL As Long             'realm level
    RANK_MOD As Long                'rank skill modifier
    
    REALM_TITLE As String           'realm rank title
    REALM_POINTS As Long            'rps
    REALM_ABILITY_POINTS As Long    'rsps
    REALM_ABILITIES(40) As Long     'realm ability Index
    
    STYLE_COUNT As Long
    SPELL_COUNT As Long
    
    EVADE As Long
    PARRY As Long
    CDHIT As Long
    CAST_SPD As Long
End Type

Public SpecialStats(1, 50) As Long

Public StatMatrixRowName(40) As String
Public StatMatrixColumnName(300) As String

Private RealmRaces(2, 6) As String      'Realm/Race
Private RealmClass(2, 15) As String     'Realm/Class

Public Const TOON_MELEE As Long = 1
Public Const TOON_MAGIC As Long = 2
Public Const TOON_HYBRID As Long = TOON_MELEE Or TOON_MAGIC

Public Const SPELL_PATH_SPECLINE As Long = 1
Public Const SPELL_PATH_BASELINE As Long = 2

Public TOON As TOON_TYPE

Public Const LOTM_MAX_STAT_CAP_BONUS As Long = 8
Public Const LOTM_MAX_RESIST_CAP_BONUS As Long = 8
Public Const LOTM_MAX_ESSENCE_RESIST_BONUS As Long = 26

Public Const SPELL_LEVEL As Long = 0
Public Const SPELL_NAME As Long = 1
Public Const SPELL_TARGET As Long = 2
Public Const SPELL_CAST_INFO As Long = 3
Public Const SPELL_RANGE_INFO As Long = 4
Public Const SPELL_EFFECT As Long = 5
Public Const SPELL_COST As Long = 6

Public Const STYLE_LEVEL As Long = 0
Public Const STYLE_NAME As Long = 1
Public Const STYLE_PREREQ As Long = 2
Public Const STYLE_DEFENSE As Long = 3
Public Const STYLE_ATTACK As Long = 4
Public Const STYLE_FATIGUE As Long = 5
Public Const STYLE_DAMAGE As Long = 6
Public Const STYLE_EFFECT As Long = 7
Public Const STYLE_DRAWTIME As Long = 8

Public Const GENDER_MALE As Long = 0
Public Const GENDER_FEMALE As Long = 1

Public Const WS_NO_VALUE As Long = 9999

Public Const ML_BANELORD As Long = 0
Public Const ML_BATTLEMASTER As Long = 1
Public Const ML_CONVOKER As Long = 2
Public Const ML_PERFECTER As Long = 3
Public Const ML_SOJOURNER As Long = 4
Public Const ML_SPYMASTER As Long = 5
Public Const ML_STORMLORD As Long = 6
Public Const ML_WARLORD As Long = 7

Public Const RESET_TOON_FULL As Long = 0
Public Const RESET_TOON_REALM As Long = 1
Public Const RESET_TOON_RACE As Long = 2
Public Const RESET_TOON_CLASS As Long = 3

Public Const REALM_ALBION As Long = 0
Public Const REALM_HIBERNIA As Long = 1
Public Const REALM_MIDGARD As Long = 2

Public Const SM_STR As Long = 1
Public Const SM_CON As Long = 2
Public Const SM_DEX As Long = 3
Public Const SM_QUI As Long = 4
Public Const SM_INT As Long = 5
Public Const SM_EMP As Long = 6
Public Const SM_PIE As Long = 7
Public Const SM_CHA As Long = 8
Public Const SM_ACU As Long = 9
Public Const SM_POW As Long = 10
Public Const SM_HIT As Long = 11

Public Const SM_CRUSH_RESIST As Long = 12
Public Const SM_SLASH_RESIST As Long = 13
Public Const SM_THRUST_RESIST As Long = 14
Public Const SM_BODY_RESIST As Long = 15
Public Const SM_COLD_RESIST As Long = 16
Public Const SM_ENERGY_RESIST As Long = 17
Public Const SM_HEAT_RESIST As Long = 18
Public Const SM_MATTER_RESIST As Long = 19
Public Const SM_SPIRIT_RESIST As Long = 20

Public Const SM_BLADES As Long = 21
Public Const SM_BLUNT As Long = 22
Public Const SM_PIERCE As Long = 23
Public Const SM_LARGEWEAP As Long = 24
Public Const SM_CELTICSPEAR As Long = 25
Public Const SM_SCYTHE As Long = 26
Public Const SM_CELTICDUAL As Long = 27
Public Const SM_RECURVE As Long = 28

Public Const SM_ARBOREAL As Long = 29
Public Const SM_CREEPING As Long = 30
Public Const SM_VERDANT As Long = 31
Public Const SM_LIGHT As Long = 32
Public Const SM_MANA As Long = 33
Public Const SM_VOID As Long = 34
Public Const SM_MENTALISM As Long = 35
Public Const SM_ENCHANTMENTS As Long = 36
Public Const SM_DEMENTIA As Long = 37
Public Const SM_VAMPEMBRACE As Long = 38
Public Const SM_SHADOWMASTERY As Long = 39
Public Const SM_PHANTASMALWAIL As Long = 40
Public Const SM_SPECTRALGUARD As Long = 41
Public Const SM_ETHEREALSHRIEK As Long = 42
Public Const SM_NURTURE As Long = 43
Public Const SM_NATURE As Long = 44
Public Const SM_REGROWTH As Long = 45
Public Const SM_MUSIC As Long = 46
Public Const SM_VALOR As Long = 47

Public Const SM_CRUSH As Long = 48
Public Const SM_SLASH As Long = 49
Public Const SM_THRUST As Long = 50
Public Const SM_DUALWIELD As Long = 51
Public Const SM_CROSSBOW As Long = 52
Public Const SM_POLEARM As Long = 53
Public Const SM_TWOHANDED As Long = 54
Public Const SM_STAFF As Long = 55
Public Const SM_FLEXIBLE As Long = 56
Public Const SM_LONGBOW As Long = 57

Public Const SM_COLD As Long = 58
Public Const SM_EARTH As Long = 59
Public Const SM_FIRE As Long = 60
Public Const SM_WIND As Long = 61
Public Const SM_MATTER As Long = 62
Public Const SM_MIND As Long = 63
Public Const SM_BODY As Long = 64
Public Const SM_SPIRIT As Long = 65
Public Const SM_SOULRENDING As Long = 66
Public Const SM_DEATHSERVANT As Long = 67
Public Const SM_DEATHSIGHT As Long = 68
Public Const SM_PAINWORKING As Long = 69
Public Const SM_INSTRUMENTS As Long = 70
Public Const SM_ENHANCEMENT As Long = 71
Public Const SM_REJUVENATION As Long = 72
Public Const SM_SMITE As Long = 73
Public Const SM_CHANTS As Long = 74

Public Const SM_HAMMER As Long = 75
Public Const SM_AXE As Long = 76
Public Const SM_SWORD As Long = 77
Public Const SM_SPEAR As Long = 78
Public Const SM_HANDTOHAND As Long = 79
Public Const SM_COMPOSITEBOW As Long = 80
Public Const SM_THROWN As Long = 81
Public Const SM_LEFTAXE As Long = 82

Public Const SM_BONEARMY As Long = 83
Public Const SM_DARKNESS As Long = 84
Public Const SM_SUPPRESSION As Long = 85
Public Const SM_MENDING As Long = 86
Public Const SM_AUGMENTATION As Long = 87
Public Const SM_BEASTCRAFT As Long = 88
Public Const SM_RUNECARVING As Long = 89
Public Const SM_CAVEMAGIC As Long = 90
Public Const SM_BATTLESONGS As Long = 91
Public Const SM_SUMMONING As Long = 92
Public Const SM_STORMCALLING As Long = 93
Public Const SM_ODINSWILL As Long = 94
Public Const SM_CURSING As Long = 95
Public Const SM_HEXING As Long = 96
Public Const SM_PACIFICATION As Long = 97
Public Const SM_WITCHCRAFT As Long = 98

Public Const SM_STEALTH As Long = 99
Public Const SM_CRITICALSTRIKE As Long = 100
Public Const SM_ENVENOM As Long = 101
Public Const SM_SHIELD As Long = 102
Public Const SM_PARRY As Long = 103

Public Const SM_STR_CAP As Long = 104
Public Const SM_CON_CAP As Long = 105
Public Const SM_DEX_CAP As Long = 106
Public Const SM_QUI_CAP As Long = 107
Public Const SM_INT_CAP As Long = 108
Public Const SM_EMP_CAP As Long = 109
Public Const SM_PIE_CAP As Long = 110
Public Const SM_CHA_CAP As Long = 111
Public Const SM_HIT_CAP As Long = 112
Public Const SM_POW_CAP As Long = 113
Public Const SM_ACU_CAP As Long = 114

Public Const SM_AFBONUS As Long = 115
Public Const SM_PERCPOWER As Long = 116
Public Const SM_ARCHERYDAMAGE As Long = 117
Public Const SM_ARCHERYRANGE As Long = 118
Public Const SM_ARCHERYSPEED As Long = 119
Public Const SM_FATIGUE As Long = 120
Public Const SM_HEALINGBONUS As Long = 121
Public Const SM_MELEESPEED As Long = 122
Public Const SM_MELEEDAMAGE As Long = 123
Public Const SM_MELEESTYLE As Long = 124
Public Const SM_SPELLDAMAGE As Long = 125
Public Const SM_SPELLDURATION As Long = 126
Public Const SM_SPELLPIERCE As Long = 127
Public Const SM_SPELLSPEED As Long = 128
Public Const SM_SPELLRANGE As Long = 129
Public Const SM_BUFFBONUS As Long = 130
Public Const SM_DEBUFFBONUS As Long = 131
Public Const SM_TOAUNIQUE As Long = 132

Public Const SM_ARROWRECOVERY As Long = 133
Public Const SM_BLADETURN As Long = 134
Public Const SM_BLOCKBONUS As Long = 135
Public Const SM_CONCENTRATION As Long = 136
Public Const SM_DAMAGEREDUCTION As Long = 137
Public Const SM_EXPLOSSREDUCTION As Long = 138
Public Const SM_DEFENSIVEBONUS As Long = 139
Public Const SM_EVADEBONUS As Long = 140

Public Const SM_NEGEFFECTDURATION As Long = 141
Public Const SM_PARRYBONUS As Long = 142
Public Const SM_PIECEABLATIVE As Long = 143
Public Const SM_REACTIONARYBONUS As Long = 144
Public Const SM_SPELLCOSTBONUS As Long = 145
Public Const SM_STYLECOSTBONUS As Long = 146
Public Const SM_TOHITBONUS As Long = 147
Public Const SM_PVEUNIQUE As Long = 148

Public Const SM_ALLFOCUS As Long = 149
Public Const SM_ARBOREAL_FOCUS As Long = 150
Public Const SM_CREEPING_FOCUS As Long = 151
Public Const SM_VERDANT_FOCUS As Long = 152
Public Const SM_ETHEREAL_FOCUS As Long = 153
Public Const SM_PHANTASMAL_FOCUS As Long = 154
Public Const SM_SPECTRAL_FOCUS As Long = 155
Public Const SM_LIGHT_FOCUS As Long = 156
Public Const SM_MANA_FOCUS As Long = 157
Public Const SM_VOID_FOCUS As Long = 158
Public Const SM_ENCHANTMENT_FOCUS As Long = 159
Public Const SM_MENTALISM_FOCUS As Long = 160

Public Const SM_BODY_FOCUS As Long = 161
Public Const SM_MATTER_FOCUS As Long = 162
Public Const SM_SPIRIT_FOCUS As Long = 163
Public Const SM_DEATHSIGHT_FOCUS As Long = 164
Public Const SM_DEATHSERVANT_FOCUS As Long = 165
Public Const SM_PAINWORKING_FOCUS As Long = 166
Public Const SM_MIND_FOCUS As Long = 167
Public Const SM_COLD_FOCUS As Long = 168
Public Const SM_EARTH_FOCUS As Long = 169
Public Const SM_WIND_FOCUS As Long = 170
Public Const SM_FIRE_FOCUS As Long = 171

Public Const SM_BONE_FOCUS As Long = 172
Public Const SM_DARKNESS_FOCUS As Long = 173
Public Const SM_SUPPRESSION_FOCUS As Long = 174
Public Const SM_RUNECARVING_FOCUS As Long = 175
Public Const SM_SUMMONING_FOCUS As Long = 176
Public Const SM_CURSING_FOCUS As Long = 177
Public Const SM_HEXING_FOCUS As Long = 178
Public Const SM_WITCHCRAFT_FOCUS As Long = 179

Public Const SM_ALLMELEE As Long = 180
Public Const SM_ALLMAGIC As Long = 181
Public Const SM_ALLDUALWIELD As Long = 182
Public Const SM_ALLARCHERY As Long = 183

Public Const SM_LOTM_STR_CAP As Long = 184
Public Const SM_LOTM_CON_CAP As Long = 185
Public Const SM_LOTM_DEX_CAP As Long = 186
Public Const SM_LOTM_QUI_CAP As Long = 187
Public Const SM_LOTM_ACU_CAP As Long = 188
Public Const SM_LOTM_STR_CAP_PLUS As Long = 189
Public Const SM_LOTM_CON_CAP_PLUS As Long = 190
Public Const SM_LOTM_DEX_CAP_PLUS As Long = 191
Public Const SM_LOTM_QUI_CAP_PLUS As Long = 192
Public Const SM_LOTM_ACU_CAP_PLUS As Long = 193

Public Const SM_LOTM_CRUSH_RESIST_CAP As Long = 194
Public Const SM_LOTM_SLASH_RESIST_CAP As Long = 195
Public Const SM_LOTM_THRUST_RESIST_CAP As Long = 196
Public Const SM_LOTM_BODY_RESIST_CAP As Long = 197
Public Const SM_LOTM_COLD_RESIST_CAP As Long = 198
Public Const SM_LOTM_ENERGY_RESIST_CAP As Long = 199
Public Const SM_LOTM_HEAT_RESIST_CAP As Long = 200
Public Const SM_LOTM_MATTER_RESIST_CAP As Long = 201
Public Const SM_LOTM_SPIRIT_RESIST_CAP As Long = 202

Public Const SM_LOTM_CRUSH_RESIST_CAP_PLUS As Long = 203
Public Const SM_LOTM_SLASH_RESIST_CAP_PLUS As Long = 204
Public Const SM_LOTM_THRUST_RESIST_CAP_PLUS As Long = 205
Public Const SM_LOTM_BODY_RESIST_CAP_PLUS As Long = 206
Public Const SM_LOTM_COLD_RESIST_CAP_PLUS As Long = 207
Public Const SM_LOTM_ENERGY_RESIST_CAP_PLUS As Long = 208
Public Const SM_LOTM_HEAT_RESIST_CAP_PLUS As Long = 209
Public Const SM_LOTM_MATTER_RESIST_CAP_PLUS As Long = 210
Public Const SM_LOTM_SPIRIT_RESIST_CAP_PLUS As Long = 211

Public Const SM_LOTM_ENC As Long = 212
Public Const SM_LOTM_ESSENCE_RESIST As Long = 213
Public Const SM_LOTM_COIN_DROP_INCREASE As Long = 214
Public Const SM_LOTM_REZ_SICK_REDUCTION As Long = 215
Public Const SM_LOTM_SAFE_FALL As Long = 216
Public Const SM_LOTM_ARROW_RECOVERY As Long = SM_ARROWRECOVERY
Public Const SM_LOTM_REALMPOINT_INCREASE As Long = 217
Public Const SM_LOTM_PARRY As Long = 218
Public Const SM_LOTM_BLOCK As Long = 219
Public Const SM_LOTM_EVADE As Long = 220
Public Const SM_LOTM_SIEGE_DAMAGE_REDUCTION As Long = 221
Public Const SM_LOTM_SPELL_LEVEL_INCREASE As Long = 222
Public Const SM_LOTM_CROWDCONTROL_REDUCTION As Long = 223
Public Const SM_LOTM_DAMAGE_INCREASE As Long = 224
Public Const SM_LOTM_PHYSICAL_DAMAGE_DECREASE As Long = 225
Public Const SM_LOTM_HEALTH_REGEN As Long = 226
Public Const SM_LOTM_POWER_REGEN As Long = 227
Public Const SM_LOTM_ENDURANCE_REGEN As Long = 228
Public Const SM_LOTM_SIEGE_SPEED_INCREASE As Long = 229
Public Const SM_LOTM_WATERBREATHING As Long = 230
    
Public Const SM_LOTM_AURAMANIP As Long = 231
Public Const SM_LOTM_FISTWRAP As Long = 232
Public Const SM_LOTM_MAGNETISM As Long = 233
Public Const SM_LOTM_MAULERSTAFF As Long = 234
Public Const SM_LOTM_POWERSTRIKES As Long = 235

Public Const SM_LOC_TOTAL As Long = 0
Public Const SM_LOC_BASE As Long = 1
Public Const SM_LOC_CREATION As Long = 2
Public Const SM_LOC_LEVEL As Long = 3
Public Const SM_LOC_HEAD As Long = 4
Public Const SM_LOC_CHEST As Long = 5
Public Const SM_LOC_ARMS As Long = 6
Public Const SM_LOC_HANDS As Long = 7
Public Const SM_LOC_LEGS As Long = 8
Public Const SM_LOC_FEET As Long = 9
Public Const SM_LOC_CLOAK As Long = 10
Public Const SM_LOC_NECK As Long = 11
Public Const SM_LOC_GEM As Long = 12
Public Const SM_LOC_BELT As Long = 13
Public Const SM_LOC_RRING As Long = 14
Public Const SM_LOC_LRING As Long = 15
Public Const SM_LOC_RBRACER As Long = 16
Public Const SM_LOC_LBRACER As Long = 17
Public Const SM_LOC_RHAND As Long = 18
Public Const SM_LOC_LHAND As Long = 19
Public Const SM_LOC_2HAND As Long = 20
Public Const SM_LOC_RANGED As Long = 21
' -- 5th slot sc bonus locations
Public Const SM_LOC_HEAD_5 As Long = 22
Public Const SM_LOC_CHEST_5 As Long = 23
Public Const SM_LOC_ARMS_5 As Long = 24
Public Const SM_LOC_HANDS_5 As Long = 25
Public Const SM_LOC_LEGS_5 As Long = 26
Public Const SM_LOC_FEET_5 As Long = 27
Public Const SM_LOC_RHAND_5 As Long = 28
Public Const SM_LOC_LHAND_5 As Long = 29
Public Const SM_LOC_2HAND_5 As Long = 30
Public Const SM_LOC_RANGED_5 As Long = 31
Public Const SM_LOC_MYTHICAL As Long = 32

Public Const SM_LOC_BUFFS As Long = 33
Public Const SM_LOC_REALMABILITY As Long = 34
Public Const SM_LOC_VAMPIIR_BUFF As Long = 35

Public Const TCA_ARMSMAN As Long = 0
Public Const TCA_CABALIST As Long = 1
Public Const TCA_CLERIC As Long = 2
Public Const TCA_FRIAR As Long = 3
Public Const TCA_HERETIC As Long = 4
Public Const TCA_INFILTRATOR As Long = 5
Public Const TCA_MERCENARY As Long = 6
Public Const TCA_MINSTREL As Long = 7
Public Const TCA_NECROMANCER As Long = 8
Public Const TCA_PALADIN As Long = 9
Public Const TCA_REAVER As Long = 10
Public Const TCA_SCOUT As Long = 11
Public Const TCA_SORCERER As Long = 12
Public Const TCA_THEURGIST As Long = 13
Public Const TCA_WIZARD As Long = 14
Public Const TCA_MAULER As Long = 15

Public Const TCH_ANIMIST As Long = 0
Public Const TCH_BAINSHEE As Long = 1
Public Const TCH_BARD As Long = 2
Public Const TCH_BLADEMASTER As Long = 3
Public Const TCH_CHAMPION As Long = 4
Public Const TCH_DRUID As Long = 5
Public Const TCH_ELDRITCH As Long = 6
Public Const TCH_ENCHANTER As Long = 7
Public Const TCH_HERO As Long = 8
Public Const TCH_MENTALIST As Long = 9
Public Const TCH_NIGHTSHADE As Long = 10
Public Const TCH_RANGER As Long = 11
Public Const TCH_VALEWALKER As Long = 12
Public Const TCH_VAMPIIR As Long = 13
Public Const TCH_WARDEN As Long = 14
Public Const TCH_MAULER As Long = 15

Public Const TCM_BERSERKER As Long = 0
Public Const TCM_BONEDANCER As Long = 1
Public Const TCM_HEALER As Long = 2
Public Const TCM_HUNTER  As Long = 3
Public Const TCM_RUNEMASTER As Long = 4
Public Const TCM_SAVAGE As Long = 5
Public Const TCM_SHADOWBLADE As Long = 6
Public Const TCM_SHAMAN As Long = 7
Public Const TCM_SKALD As Long = 8
Public Const TCM_SPIRITMASTER As Long = 9
Public Const TCM_THANE As Long = 10
Public Const TCM_VALKYRIE As Long = 11
Public Const TCM_WARLOCK As Long = 12
Public Const TCM_WARRIOR As Long = 13
Public Const TCM_MAULER As Long = 14
    
Public Const TRA_AVALONIAN As Long = 0
Public Const TRA_BRITON As Long = 1
Public Const TRA_HALFOGRE As Long = 2
Public Const TRA_HIGHLANDER As Long = 3
Public Const TRA_INCONNU As Long = 4
Public Const TRA_SARACEN As Long = 5
Public Const TRA_MINOTAUR As Long = 6

Public Const TRH_CELT As Long = 0
Public Const TRH_ELF As Long = 1
Public Const TRH_FIRBOLG As Long = 2
Public Const TRH_LURIKEEN As Long = 3
Public Const TRH_SHAR As Long = 4
Public Const TRH_SYLVAN As Long = 5
Public Const TRH_MINOTAUR As Long = 6
    
Public Const TRM_DWARF As Long = 0
Public Const TRM_FROSTALF As Long = 1
Public Const TRM_KOBOLD As Long = 2
Public Const TRM_NORSEMAN As Long = 3
Public Const TRM_TROLL As Long = 4
Public Const TRM_VALKYN As Long = 5
Public Const TRM_MINOTAUR As Long = 6

Public Sub ClearStyles()

    Dim iCtr As Long
    Dim jCtr As Long
    Dim kCtr As Long
    Dim lCtr As Long
    
    For iCtr = 0 To 10
    
        STYLES(iCtr).STYLE_PATH_NAME = vbNullString
        STYLES(iCtr).AUTO_TRAIN = False
        STYLES(iCtr).STYLE_NO_STYLES = False
        
        For lCtr = 0 To 10
            STYLES(iCtr).AUTO_TRAIN_CLASS(lCtr) = -1
        Next lCtr
        
        For jCtr = 0 To 22
            For kCtr = 0 To 15
                STYLES(iCtr).STYLE(jCtr).STYLE_DETAILS(kCtr) = vbNullString
            Next kCtr
        Next jCtr
    Next iCtr
    
End Sub

Public Sub ClearSpells()

    Dim iCtr As Long
    Dim jCtr As Long
    Dim kCtr As Long
    Dim lCtr As Long
    Dim mCtr As Long
    
    For iCtr = 0 To 10
        SPELLS(iCtr).SPELL_PATH_NAME = vbNullString
        SPELLS(iCtr).SPELL_PATH_CATEGORY = vbNull
        SPELLS(iCtr).AUTO_TRAIN = False
        
        For mCtr = 0 To 10
            SPELLS(iCtr).AUTO_TRAIN_CLASS(mCtr) = -1
        Next mCtr
        
        For jCtr = 0 To 20
            SPELLS(iCtr).SPELL_CLASS(jCtr).SPELL_TYPE_NAME = vbNullString
            For kCtr = 0 To 20
                For lCtr = 0 To 7
                    SPELLS(iCtr).SPELL_CLASS(jCtr).SPELL_DETAILS(kCtr, lCtr) = vbNullString
                Next lCtr
            Next kCtr
        Next jCtr
    Next iCtr
    
End Sub

Public Sub ResetCharacter(ByRef tToon As TOON_TYPE)

    Dim i As Long
    Dim j As Long
    
    With tToon
        
        For i = 0 To 24
            If i <> SM_LOC_BUFFS Then
                For j = 0 To 200
                    .STAT_MATRIX(j, i) = 0
                Next j
            End If
        Next i
        
        .CREATION_POINTS = 30
        .LEVEL = 50
        
        Call ResetRealmAbilityArray(TOON)
                
        .ML_OPTION_1 = WS_NO_VALUE
        .ML_OPTION_2 = WS_NO_VALUE
       
    End With 'ttoon
    
    
End Sub

Public Function GetClassComboID(REALM As Long, CLASS As Long, Combo As ComboBox) As Long

    Dim Ctr As Long
    
    For Ctr = 0 To Combo.ListCount - 1
        If Combo.list(Ctr) = RealmClass(REALM, CLASS) Then Exit For
    Next Ctr
    
    GetClassComboID = Ctr
    
End Function

Private Sub Init_StatMatrixRowName()

    StatMatrixRowName(SM_LOC_TOTAL) = "Total Skill"
    StatMatrixRowName(SM_LOC_BASE) = "Racial Base"
    StatMatrixRowName(SM_LOC_CREATION) = "Added at Creation"
    StatMatrixRowName(SM_LOC_LEVEL) = "Added from Level"
    StatMatrixRowName(SM_LOC_HEAD) = "Head"
    StatMatrixRowName(SM_LOC_CHEST) = "Chest"
    StatMatrixRowName(SM_LOC_ARMS) = "Arms"
    StatMatrixRowName(SM_LOC_HANDS) = "Hands"
    StatMatrixRowName(SM_LOC_LEGS) = "Legs"
    StatMatrixRowName(SM_LOC_FEET) = "Feet"
    StatMatrixRowName(SM_LOC_CLOAK) = "Cloak"
    StatMatrixRowName(SM_LOC_NECK) = "Necklace"
    StatMatrixRowName(SM_LOC_GEM) = "Jewel"
    StatMatrixRowName(SM_LOC_BELT) = "Belt"
    StatMatrixRowName(SM_LOC_RRING) = "Right Ring"
    StatMatrixRowName(SM_LOC_LRING) = "Left Ring"
    StatMatrixRowName(SM_LOC_RBRACER) = "Right Bracer"
    StatMatrixRowName(SM_LOC_LBRACER) = "Left Bracer"
    StatMatrixRowName(SM_LOC_RHAND) = "Main Hand"
    StatMatrixRowName(SM_LOC_LHAND) = "Off Hand"
    StatMatrixRowName(SM_LOC_2HAND) = "Two Handed"
    StatMatrixRowName(SM_LOC_RANGED) = "Ranged"
    StatMatrixRowName(SM_LOC_HEAD_5) = "Head 5th"
    StatMatrixRowName(SM_LOC_CHEST_5) = "Chest 5th"
    StatMatrixRowName(SM_LOC_ARMS_5) = "Arms 5th"
    StatMatrixRowName(SM_LOC_HANDS_5) = "Hands 5th"
    StatMatrixRowName(SM_LOC_LEGS_5) = "Legs 5th"
    StatMatrixRowName(SM_LOC_FEET_5) = "Feet 5th"
    StatMatrixRowName(SM_LOC_RHAND_5) = "Main Hand 5th"
    StatMatrixRowName(SM_LOC_LHAND_5) = "Off Hand 5th"
    StatMatrixRowName(SM_LOC_2HAND_5) = "Two Handed 5th"
    StatMatrixRowName(SM_LOC_RANGED_5) = "Ranged 5th"
    StatMatrixRowName(SM_LOC_MYTHICAL) = "Mythical"
    StatMatrixRowName(SM_LOC_BUFFS) = "Buffs"
    StatMatrixRowName(SM_LOC_REALMABILITY) = "Realm Ability"
    StatMatrixRowName(SM_LOC_VAMPIIR_BUFF) = "Vamp Stats"
    
End Sub

Public Sub CalculatePSTValues(ByRef tToon As TOON_TYPE)

    Dim LVL As Long
    Dim Lo As Single
        
    LVL = Trunc(tToon.LEVEL)
    
    If (LVL > 5) And (LVL < 51) Then
        tToon.STAT_MATRIX(tToon.pStat, SM_LOC_LEVEL) = Trunc(LVL - 5)
        
        Lo = ((LVL - 5) / 2) - Trunc((LVL - 5) / 2)
        
        If Lo <> 0 Then
            tToon.STAT_MATRIX(tToon.sStat, SM_LOC_LEVEL) = Trunc((LVL - 5) / 2) + 1
        Else
            tToon.STAT_MATRIX(tToon.sStat, SM_LOC_LEVEL) = Trunc((LVL - 5) / 2)
        End If
        
        Lo = ((LVL - 5) / 3) - Trunc((LVL - 5) / 3)
        
        If Lo <> 0 Then
            tToon.STAT_MATRIX(tToon.tStat, SM_LOC_LEVEL) = Trunc((LVL - 5) / 3) + 1
        Else
            tToon.STAT_MATRIX(tToon.tStat, SM_LOC_LEVEL) = Trunc((LVL - 5) / 3)
        End If
    Else
        tToon.STAT_MATRIX(tToon.pStat, SM_LOC_LEVEL) = 0
        tToon.STAT_MATRIX(tToon.sStat, SM_LOC_LEVEL) = 0
        tToon.STAT_MATRIX(tToon.tStat, SM_LOC_LEVEL) = 0
    End If
    
    If (tToon.REALM = REALM_HIBERNIA) And (tToon.CLASS = TCH_VAMPIIR) Then
        tToon.STAT_MATRIX(SM_STR, SM_LOC_VAMPIIR_BUFF) = SpecialStats(0, CLng(tToon.LEVEL))
        tToon.STAT_MATRIX(SM_CON, SM_LOC_VAMPIIR_BUFF) = SpecialStats(0, CLng(tToon.LEVEL))
        tToon.STAT_MATRIX(SM_DEX, SM_LOC_VAMPIIR_BUFF) = SpecialStats(0, CLng(tToon.LEVEL))
        tToon.STAT_MATRIX(SM_QUI, SM_LOC_VAMPIIR_BUFF) = SpecialStats(1, CLng(tToon.LEVEL))
    Else
        tToon.STAT_MATRIX(SM_STR, SM_LOC_VAMPIIR_BUFF) = 0
        tToon.STAT_MATRIX(SM_CON, SM_LOC_VAMPIIR_BUFF) = 0
        tToon.STAT_MATRIX(SM_DEX, SM_LOC_VAMPIIR_BUFF) = 0
        tToon.STAT_MATRIX(SM_QUI, SM_LOC_VAMPIIR_BUFF) = 0
    End If
    
End Sub

Private Sub Init_StatMatrixColumnName()

    StatMatrixColumnName(0) = vbNullString
    StatMatrixColumnName(SM_STR) = "Strength"
    StatMatrixColumnName(SM_CON) = "Constitution"
    StatMatrixColumnName(SM_DEX) = "Dexterity"
    StatMatrixColumnName(SM_QUI) = "Quickness"
    StatMatrixColumnName(SM_INT) = "Intelligence"
    StatMatrixColumnName(SM_EMP) = "Empathy"
    StatMatrixColumnName(SM_PIE) = "Piety"
    StatMatrixColumnName(SM_CHA) = "Charisma"
    StatMatrixColumnName(SM_ACU) = "Acuity"
    StatMatrixColumnName(SM_POW) = "Power"
    StatMatrixColumnName(SM_HIT) = "Hits"
    StatMatrixColumnName(SM_CRUSH_RESIST) = "Crush"
    StatMatrixColumnName(SM_SLASH_RESIST) = "Slash"
    StatMatrixColumnName(SM_THRUST_RESIST) = "Thrust"
    StatMatrixColumnName(SM_BODY_RESIST) = "Body"
    StatMatrixColumnName(SM_COLD_RESIST) = "Cold"
    StatMatrixColumnName(SM_ENERGY_RESIST) = "Energy"
    StatMatrixColumnName(SM_HEAT_RESIST) = "Heat"
    StatMatrixColumnName(SM_MATTER_RESIST) = "Matter"
    StatMatrixColumnName(SM_SPIRIT_RESIST) = "Spirit"
    StatMatrixColumnName(SM_BLADES) = "Blades"
    StatMatrixColumnName(SM_BLUNT) = "Blunt"
    StatMatrixColumnName(SM_PIERCE) = "Pierce"
    StatMatrixColumnName(SM_LARGEWEAP) = "Large Weaponry"
    StatMatrixColumnName(SM_CELTICSPEAR) = "Celtic Spear"
    StatMatrixColumnName(SM_SCYTHE) = "Scythe"
    StatMatrixColumnName(SM_CELTICDUAL) = "Celtic Dual"
    StatMatrixColumnName(SM_RECURVE) = "Archery"
    StatMatrixColumnName(SM_ARBOREAL) = "Arboreal Path"
    StatMatrixColumnName(SM_CREEPING) = "Creeping Path"
    StatMatrixColumnName(SM_VERDANT) = "Verdant Path"
    StatMatrixColumnName(SM_LIGHT) = "Light"
    StatMatrixColumnName(SM_MANA) = "Mana"
    StatMatrixColumnName(SM_VOID) = "Void"
    StatMatrixColumnName(SM_MENTALISM) = "Mentalism"
    StatMatrixColumnName(SM_ENCHANTMENTS) = "Enchantments"
    StatMatrixColumnName(SM_DEMENTIA) = "Dementia"
    StatMatrixColumnName(SM_VAMPEMBRACE) = "Vampiiric Embrace"
    StatMatrixColumnName(SM_SHADOWMASTERY) = "Shadow Mastery"
    StatMatrixColumnName(SM_PHANTASMALWAIL) = "Phantasmal Wail"
    StatMatrixColumnName(SM_SPECTRALGUARD) = "Spectral Guard"
    StatMatrixColumnName(SM_ETHEREALSHRIEK) = "Ethereal Shriek"
    StatMatrixColumnName(SM_NURTURE) = "Nurture"
    StatMatrixColumnName(SM_NATURE) = "Nature"
    StatMatrixColumnName(SM_REGROWTH) = "Regrowth"
    StatMatrixColumnName(SM_MUSIC) = "Music"
    StatMatrixColumnName(SM_VALOR) = "Valor"
    StatMatrixColumnName(SM_CRUSH) = "Crush"
    StatMatrixColumnName(SM_SLASH) = "Slash"
    StatMatrixColumnName(SM_THRUST) = "Thrust"
    StatMatrixColumnName(SM_DUALWIELD) = "Dual Wield"
    StatMatrixColumnName(SM_CROSSBOW) = "Crossbow"
    StatMatrixColumnName(SM_POLEARM) = "Polearm"
    StatMatrixColumnName(SM_TWOHANDED) = "Two Handed"
    StatMatrixColumnName(SM_STAFF) = "Staff"
    StatMatrixColumnName(SM_FLEXIBLE) = "Flexible"
    StatMatrixColumnName(SM_LONGBOW) = "Archery"
    StatMatrixColumnName(SM_COLD) = "Cold Magic"
    StatMatrixColumnName(SM_EARTH) = "Earth Magic"
    StatMatrixColumnName(SM_FIRE) = "Fire Magic"
    StatMatrixColumnName(SM_WIND) = "Wind Magic"
    StatMatrixColumnName(SM_MATTER) = "Matter Magic"
    StatMatrixColumnName(SM_MIND) = "Mind Magic"
    StatMatrixColumnName(SM_BODY) = "Body Magic"
    StatMatrixColumnName(SM_SPIRIT) = "Spirit Magic"
    StatMatrixColumnName(SM_SOULRENDING) = "Soulrending"
    StatMatrixColumnName(SM_DEATHSERVANT) = "Death Servant"
    StatMatrixColumnName(SM_DEATHSIGHT) = "Deathsight"
    StatMatrixColumnName(SM_PAINWORKING) = "Painworking"
    StatMatrixColumnName(SM_INSTRUMENTS) = "Instruments"
    StatMatrixColumnName(SM_ENHANCEMENT) = "Enhancement"
    StatMatrixColumnName(SM_REJUVENATION) = "Rejuvenation"
    StatMatrixColumnName(SM_SMITE) = "Smite"
    StatMatrixColumnName(SM_CHANTS) = "Chants"
    StatMatrixColumnName(SM_HAMMER) = "Hammer"
    StatMatrixColumnName(SM_AXE) = "Axe"
    StatMatrixColumnName(SM_SWORD) = "Sword"
    StatMatrixColumnName(SM_SPEAR) = "Spear"
    StatMatrixColumnName(SM_HANDTOHAND) = "Hand to Hand"
    StatMatrixColumnName(SM_COMPOSITEBOW) = "Archery"
    StatMatrixColumnName(SM_THROWN) = "Thrown Weapons"
    StatMatrixColumnName(SM_LEFTAXE) = "Left Axe"
    StatMatrixColumnName(SM_BONEARMY) = "Bone Army"
    StatMatrixColumnName(SM_DARKNESS) = "Darkness"
    StatMatrixColumnName(SM_SUPPRESSION) = "Suppression"
    StatMatrixColumnName(SM_MENDING) = "Mending"
    StatMatrixColumnName(SM_AUGMENTATION) = "Augmentation"
    StatMatrixColumnName(SM_BEASTCRAFT) = "Beastcraft"
    StatMatrixColumnName(SM_RUNECARVING) = "Runecarving"
    StatMatrixColumnName(SM_CAVEMAGIC) = "Cave Magic"
    StatMatrixColumnName(SM_BATTLESONGS) = "Battlesongs"
    StatMatrixColumnName(SM_SUMMONING) = "Summoning"
    StatMatrixColumnName(SM_STORMCALLING) = "Stormcalling"
    StatMatrixColumnName(SM_ODINSWILL) = "Odin's Will"
    StatMatrixColumnName(SM_CURSING) = "Cursing"
    StatMatrixColumnName(SM_HEXING) = "Hexing"
    StatMatrixColumnName(SM_PACIFICATION) = "Pacification"
    StatMatrixColumnName(SM_WITCHCRAFT) = "Witchcraft"
    StatMatrixColumnName(SM_STEALTH) = "Stealth"
    StatMatrixColumnName(SM_CRITICALSTRIKE) = "Critical Strike"
    StatMatrixColumnName(SM_ENVENOM) = "Envenom"
    StatMatrixColumnName(SM_SHIELD) = "Shield"
    StatMatrixColumnName(SM_PARRY) = "Parry"
    StatMatrixColumnName(SM_STR_CAP) = "Strength Cap"
    StatMatrixColumnName(SM_CON_CAP) = "Constitution Cap"
    StatMatrixColumnName(SM_DEX_CAP) = "Dexterity Cap"
    StatMatrixColumnName(SM_QUI_CAP) = "Quickness Cap"
    StatMatrixColumnName(SM_INT_CAP) = "Intelligence Cap"
    StatMatrixColumnName(SM_EMP_CAP) = "Empathy Cap"
    StatMatrixColumnName(SM_PIE_CAP) = "Piety Cap"
    StatMatrixColumnName(SM_CHA_CAP) = "Charisma Cap"
    StatMatrixColumnName(SM_HIT_CAP) = "Hits Cap"
    StatMatrixColumnName(SM_POW_CAP) = "Power Cap"
    StatMatrixColumnName(SM_ACU_CAP) = "Acuity Cap"
    StatMatrixColumnName(SM_AFBONUS) = "Armor Factor"
    StatMatrixColumnName(SM_PERCPOWER) = "% Power Pool"
    StatMatrixColumnName(SM_ARCHERYDAMAGE) = "% Archery Damage"
    StatMatrixColumnName(SM_ARCHERYRANGE) = "% Archery Range"
    StatMatrixColumnName(SM_ARCHERYSPEED) = "% Archery Speed"
    StatMatrixColumnName(SM_FATIGUE) = "Fatigue"
    StatMatrixColumnName(SM_HEALINGBONUS) = "% Healing Effectiveness"
    StatMatrixColumnName(SM_MELEESPEED) = "% Melee Combat Speed"
    StatMatrixColumnName(SM_MELEEDAMAGE) = "% Melee Damage"
    StatMatrixColumnName(SM_MELEESTYLE) = "% Melee Style Damage"
    StatMatrixColumnName(SM_SPELLDAMAGE) = "% Spell Damage"
    StatMatrixColumnName(SM_SPELLDURATION) = "% Spell Duration"
    StatMatrixColumnName(SM_SPELLPIERCE) = "% Spell Pierce"
    StatMatrixColumnName(SM_SPELLSPEED) = "% Spell Haste"
    StatMatrixColumnName(SM_SPELLRANGE) = "% Spell Range"
    StatMatrixColumnName(SM_BUFFBONUS) = "% Stat Buff Effectiveness"
    StatMatrixColumnName(SM_DEBUFFBONUS) = "% Stat Debuff Effectiveness"
    StatMatrixColumnName(SM_TOAUNIQUE) = "Unique ToA Bonus"
    StatMatrixColumnName(SM_ARROWRECOVERY) = "Arcane Siphon" '% Arrow Recovery"
    StatMatrixColumnName(SM_BLADETURN) = "Bladeturn Reinforcement"
    StatMatrixColumnName(SM_BLOCKBONUS) = "% Block"
    StatMatrixColumnName(SM_CONCENTRATION) = "Concentration"
    StatMatrixColumnName(SM_DAMAGEREDUCTION) = "Damage Reduction"
    StatMatrixColumnName(SM_EXPLOSSREDUCTION) = "% Death Exp Loss Reduction"
    StatMatrixColumnName(SM_DEFENSIVEBONUS) = "Defensive"
    StatMatrixColumnName(SM_EVADEBONUS) = "% Evade"
    StatMatrixColumnName(SM_NEGEFFECTDURATION) = "% Negative Effect Duration"
    StatMatrixColumnName(SM_PARRYBONUS) = "% Parry"
    StatMatrixColumnName(SM_PIECEABLATIVE) = "Piece Ablative"
    StatMatrixColumnName(SM_REACTIONARYBONUS) = "% Reactionary Style Damage"
    StatMatrixColumnName(SM_SPELLCOSTBONUS) = "% Spell Power Cost Reduction"
    StatMatrixColumnName(SM_STYLECOSTBONUS) = "% Style Cost Reduction"
    StatMatrixColumnName(SM_TOHITBONUS) = "% To-Hit"
    StatMatrixColumnName(SM_PVEUNIQUE) = "Unique PvE Bonus"
    StatMatrixColumnName(SM_ALLFOCUS) = "Focus - All Spell Lines"
    StatMatrixColumnName(SM_ARBOREAL_FOCUS) = "Arboreal Path Focus"
    StatMatrixColumnName(SM_CREEPING_FOCUS) = "Creeping Path Focus"
    StatMatrixColumnName(SM_VERDANT_FOCUS) = "Verdant Path Focus"
    StatMatrixColumnName(SM_ETHEREAL_FOCUS) = "Ethereal Shriek Focus"
    StatMatrixColumnName(SM_PHANTASMAL_FOCUS) = "Phantasmal Wail Focus"
    StatMatrixColumnName(SM_SPECTRAL_FOCUS) = "Spectral Guard Focus"
    StatMatrixColumnName(SM_LIGHT_FOCUS) = "Light Focus"
    StatMatrixColumnName(SM_MANA_FOCUS) = "Mana Focus"
    StatMatrixColumnName(SM_VOID_FOCUS) = "Void Focus"
    StatMatrixColumnName(SM_ENCHANTMENT_FOCUS) = "Enchantment Focus"
    StatMatrixColumnName(SM_MENTALISM_FOCUS) = "Mentalism Focus"
    StatMatrixColumnName(SM_BODY_FOCUS) = "Body Magic Focus"
    StatMatrixColumnName(SM_MATTER_FOCUS) = "Matter Magic Focus"
    StatMatrixColumnName(SM_SPIRIT_FOCUS) = "Spirit Magic Focus"
    StatMatrixColumnName(SM_DEATHSIGHT_FOCUS) = "Deathsight Focus"
    StatMatrixColumnName(SM_DEATHSERVANT_FOCUS) = "Death Servant Focus"
    StatMatrixColumnName(SM_PAINWORKING_FOCUS) = "Painworking Focus"
    StatMatrixColumnName(SM_MIND_FOCUS) = "Mind Magic Focus"
    StatMatrixColumnName(SM_COLD_FOCUS) = "Cold Magic Focus"
    StatMatrixColumnName(SM_EARTH_FOCUS) = "Earth Magic Focus"
    StatMatrixColumnName(SM_WIND_FOCUS) = "Wind Magic Focus"
    StatMatrixColumnName(SM_FIRE_FOCUS) = "Fire Magic Focus"
    StatMatrixColumnName(SM_BONE_FOCUS) = "Bone Army Focus"
    StatMatrixColumnName(SM_DARKNESS_FOCUS) = "Darkness Focus"
    StatMatrixColumnName(SM_SUPPRESSION_FOCUS) = "Suppression Focus"
    StatMatrixColumnName(SM_RUNECARVING_FOCUS) = "Runecarving Focus"
    StatMatrixColumnName(SM_SUMMONING_FOCUS) = "Summoning Focus"
    StatMatrixColumnName(SM_CURSING_FOCUS) = "Cursing Focus"
    StatMatrixColumnName(SM_HEXING_FOCUS) = "Hexing Focus"
    StatMatrixColumnName(SM_WITCHCRAFT_FOCUS) = "Witchcraft Focus"
    StatMatrixColumnName(SM_ALLMELEE) = "All Melee Weapon Skills"
    StatMatrixColumnName(SM_ALLMAGIC) = "All Magic Skills"
    StatMatrixColumnName(SM_ALLDUALWIELD) = "All Dual Wield Skills"
    StatMatrixColumnName(SM_ALLARCHERY) = "All Archery Skills"
    StatMatrixColumnName(SM_LOTM_ENC) = "Encumberence Increase"
    StatMatrixColumnName(SM_LOTM_ESSENCE_RESIST) = "Essence Resist"
    StatMatrixColumnName(SM_LOTM_COIN_DROP_INCREASE) = "Coin Drop Increase"
    StatMatrixColumnName(SM_LOTM_REZ_SICK_REDUCTION) = "Rez Sick Reduction"
    StatMatrixColumnName(SM_LOTM_SAFE_FALL) = "Safe Fall"
    StatMatrixColumnName(SM_LOTM_ARROW_RECOVERY) = "Arcane Siphon" 'Arrow Recovery"
    StatMatrixColumnName(SM_LOTM_STR_CAP) = "Str Cap"
    StatMatrixColumnName(SM_LOTM_STR_CAP_PLUS) = "Str Cap (+Str)"
    StatMatrixColumnName(SM_LOTM_CON_CAP) = "Con Cap"
    StatMatrixColumnName(SM_LOTM_CON_CAP_PLUS) = "Con Cap (+Con)"
    StatMatrixColumnName(SM_LOTM_DEX_CAP) = "Dex Cap"
    StatMatrixColumnName(SM_LOTM_DEX_CAP_PLUS) = "Dex Cap (+Dex)"
    StatMatrixColumnName(SM_LOTM_QUI_CAP) = "Qui Cap"
    StatMatrixColumnName(SM_LOTM_QUI_CAP_PLUS) = "Qui Cap (+Qui)"
    StatMatrixColumnName(SM_LOTM_ACU_CAP) = "Acuity Cap)"
    StatMatrixColumnName(SM_LOTM_ACU_CAP_PLUS) = "Acuity Cap (+Acu)"
    StatMatrixColumnName(SM_LOTM_CRUSH_RESIST_CAP) = "Crush Cap"
    StatMatrixColumnName(SM_LOTM_CRUSH_RESIST_CAP_PLUS) = "Crush Cap (+Crush)"
    StatMatrixColumnName(SM_LOTM_SLASH_RESIST_CAP) = "Slash Cap"
    StatMatrixColumnName(SM_LOTM_SLASH_RESIST_CAP_PLUS) = "Slash Cap (+Slash)"
    StatMatrixColumnName(SM_LOTM_THRUST_RESIST_CAP) = "Thrust Cap"
    StatMatrixColumnName(SM_LOTM_THRUST_RESIST_CAP_PLUS) = "Thrust Cap (+Thrust)"
    StatMatrixColumnName(SM_LOTM_BODY_RESIST_CAP) = "Body Cap"
    StatMatrixColumnName(SM_LOTM_BODY_RESIST_CAP_PLUS) = "Body Cap (+Body)"
    StatMatrixColumnName(SM_LOTM_COLD_RESIST_CAP) = "Cold Cap"
    StatMatrixColumnName(SM_LOTM_COLD_RESIST_CAP_PLUS) = "Cold Cap (+Cold)"
    StatMatrixColumnName(SM_LOTM_ENERGY_RESIST_CAP) = "Energy Cap"
    StatMatrixColumnName(SM_LOTM_ENERGY_RESIST_CAP_PLUS) = "Energy Cap (+Energy)"
    StatMatrixColumnName(SM_LOTM_HEAT_RESIST_CAP) = "Heat Cap"
    StatMatrixColumnName(SM_LOTM_HEAT_RESIST_CAP_PLUS) = "Heat Cap (+Heat)"
    StatMatrixColumnName(SM_LOTM_MATTER_RESIST_CAP) = "Matter Cap"
    StatMatrixColumnName(SM_LOTM_MATTER_RESIST_CAP_PLUS) = "Matter Cap (+Matter)"
    StatMatrixColumnName(SM_LOTM_SPIRIT_RESIST_CAP) = "Spirit Cap"
    StatMatrixColumnName(SM_LOTM_SPIRIT_RESIST_CAP_PLUS) = "Spirit Cap (+Spirit)"
    StatMatrixColumnName(SM_LOTM_REALMPOINT_INCREASE) = "Realm Point Bonus"
    StatMatrixColumnName(SM_LOTM_PARRY) = "% Parry"
    StatMatrixColumnName(SM_LOTM_EVADE) = "% Evade"
    StatMatrixColumnName(SM_LOTM_BLOCK) = "% Block"
    StatMatrixColumnName(SM_LOTM_SIEGE_DAMAGE_REDUCTION) = "% Siege Damage Decrease"
    StatMatrixColumnName(SM_LOTM_SPELL_LEVEL_INCREASE) = "Spell Level Increase"
    StatMatrixColumnName(SM_LOTM_CROWDCONTROL_REDUCTION) = "% Crowd Control Decrease"
    StatMatrixColumnName(SM_LOTM_DAMAGE_INCREASE) = "% Damage Increase"
    StatMatrixColumnName(SM_LOTM_PHYSICAL_DAMAGE_DECREASE) = "% Physical Damage Decrease"
    StatMatrixColumnName(SM_LOTM_HEALTH_REGEN) = "Health Regen"
    StatMatrixColumnName(SM_LOTM_POWER_REGEN) = "Power Regen"
    StatMatrixColumnName(SM_LOTM_ENDURANCE_REGEN) = "Endurance Regen"
    StatMatrixColumnName(SM_LOTM_SIEGE_SPEED_INCREASE) = "% Siege Speed Increase"
    StatMatrixColumnName(SM_LOTM_WATERBREATHING) = "Water Breathing"
    StatMatrixColumnName(SM_LOTM_AURAMANIP) = "Aura Manipulation"
    StatMatrixColumnName(SM_LOTM_FISTWRAP) = "Fist Wraps"
    StatMatrixColumnName(SM_LOTM_MAGNETISM) = "Magnetism"
    StatMatrixColumnName(SM_LOTM_MAULERSTAFF) = "Mauler Staff"
    StatMatrixColumnName(SM_LOTM_POWERSTRIKES) = "Power Strikes"

End Sub

Private Sub ResetRealmAbilityArray(ByRef tToon As TOON_TYPE)

    Dim i As Long
    
    For i = 0 To 40
        tToon.REALM_ABILITIES(i) = WS_NO_VALUE
    Next i
    
End Sub

Public Sub InitCharacterArrays(ByRef ProgressBar As Label, ByRef Status As Label)
    
    Status.Caption = "Initializing Character Matricies: Races"
    DoEvents
    Call Init_RealmRaceArray
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Status.Caption = "Initializing Character Matricies: Classes"
    DoEvents
    Call Init_RealmClassArray
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Call Init_StatMatrixColumnName
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Call Init_StatMatrixRowName
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
    Call Init_SpecialStatsArray
    ProgressBar.Width = ProgressBar.Width + SP_LG_PROGRESS_CHANGE
    
End Sub

Public Function FindCastSpeed(DEX As Long)

    Dim dCast As Long, mDex As Long
    
    If DEX <= 250 Then
        dCast = (DEX - 50) / 10
    Else
        mDex = DEX - 250
        dCast = 20 + (mDex / 20)
    End If

    FindCastSpeed = dCast
    
End Function

Public Function CalcSpecPoints(tToon As TOON_TYPE, lAutoTrain As Long) As Long
'working as intended 10/23/06

    Dim Ctr As Long
    Dim Sum As Long
    Dim Level_LO As Single
    Dim Level_HI As Long
    
    With tToon
        Level_HI = Trunc(.LEVEL)
        
        If (.LEVEL <= 0) Or (.LEVEL > 50) Then
            Sum = 0
        ElseIf .LEVEL < 6 Then
        
            For Ctr = 0 To .LEVEL
                Sum = Sum + Ctr
            Next Ctr
            
            Sum = Sum - 1
        Else
            For Ctr = 6 To Level_HI
            
                If Ctr > 40 Then
                    Sum = Sum + Trunc(.MULTIPLIER * Trunc(Ctr - 1) / 2)
                End If
                
                Sum = Sum + Trunc(.MULTIPLIER * Ctr)
                
            Next Ctr
            
            Level_LO = .LEVEL - Level_HI
            
            If Level_LO > 0 Then
            
                Sum = Sum + Trunc(.MULTIPLIER * Level_HI / 2)
                
            End If
            
            'add lvl 1-5 points
            Sum = Sum + 14
            
            If lAutoTrain <> 0 Then
            'add autotrained points
                Sum = Sum + ((lAutoTrain * (lAutoTrain + 1)) / 2) - 1
            End If
            
        End If
    End With
    
    CalcSpecPoints = Sum

End Function

Public Function FindEvade(Skill As Long, DEX As Long, QUI As Long) As Long

'5x evade level + ((((dex+qui)/2)-50)/10) *.01  +dodger ::evade
    FindEvade = ((5 * Skill) + ((((DEX + QUI) / 2) - 50) / 10))
    
End Function

Public Sub PopulateRaceSelection(Combo As ComboBox, tToon As TOON_TYPE)

    Dim iCtr As Long
    
    Combo.Clear  'clear old items from list
    For iCtr = 0 To 6
        If RealmRaces(tToon.REALM, iCtr) <> vbNullString Then
            If LCase$(RealmRaces(tToon.REALM, iCtr)) = "minotaur" And tToon.GENDER = GENDER_MALE Then
                Combo.AddItem RealmRaces(tToon.REALM, iCtr)
            ElseIf LCase$(RealmRaces(tToon.REALM, iCtr)) <> "minotaur" Then
                Combo.AddItem RealmRaces(tToon.REALM, iCtr)
            End If
        End If
    Next iCtr
    
End Sub

Public Sub PopulateClassSelection(Combo As ComboBox, tToon As TOON_TYPE)
    
    Combo.Clear 'clear old items
        
    If ((TOON.REALM < 0) Or (TOON.REALM > 2)) Then TOON.REALM = 1
    
    Select Case tToon.REALM
        Case REALM_ALBION
            Select Case tToon.RACE
                Case TRA_AVALONIAN
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_ARMSMAN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CABALIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CLERIC)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_FRIAR)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_HERETIC)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MERCENARY)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_PALADIN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SORCERER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_THEURGIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_WIZARD)
                Case TRA_BRITON
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_ARMSMAN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CABALIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CLERIC)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_FRIAR)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_HERETIC)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_INFILTRATOR)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MAULER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MERCENARY)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MINSTREL)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_NECROMANCER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_PALADIN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_REAVER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SCOUT)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SORCERER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_THEURGIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_WIZARD)
                Case TRA_HALFOGRE
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_ARMSMAN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CABALIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MERCENARY)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SORCERER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_THEURGIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_WIZARD)
                Case TRA_HIGHLANDER
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_ARMSMAN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CLERIC)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_FRIAR)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MERCENARY)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MINSTREL)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_PALADIN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SCOUT)
                Case TRA_INCONNU
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_ARMSMAN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CABALIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_HERETIC)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_INFILTRATOR)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MAULER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MERCENARY)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_NECROMANCER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_REAVER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SCOUT)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SORCERER)
                Case TRA_SARACEN
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_ARMSMAN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_CABALIST)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_INFILTRATOR)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MERCENARY)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MINSTREL)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_NECROMANCER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_PALADIN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_REAVER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SCOUT)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_SORCERER)
                Case TRA_MINOTAUR
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_ARMSMAN)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_HERETIC)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MAULER)
                    Combo.AddItem RealmClass(REALM_ALBION, TCA_MERCENARY)
            End Select
        Case REALM_HIBERNIA
            Select Case tToon.RACE
                Case TRH_CELT
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_ANIMIST)
                    If TOON.GENDER = GENDER_FEMALE Then
                        Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BAINSHEE)
                    End If
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BARD)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BLADEMASTER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_CHAMPION)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_DRUID)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_HERO)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_MAULER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_MENTALIST)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_NIGHTSHADE)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_RANGER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_VALEWALKER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_VAMPIIR)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_WARDEN)
                Case TRH_ELF
                    If TOON.GENDER = GENDER_FEMALE Then
                        Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BAINSHEE)
                    End If
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BLADEMASTER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_CHAMPION)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_ELDRITCH)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_ENCHANTER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_MENTALIST)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_NIGHTSHADE)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_RANGER)
                Case TRH_FIRBOLG
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_ANIMIST)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BARD)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BLADEMASTER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_DRUID)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_HERO)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_VALEWALKER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_WARDEN)
                Case TRH_LURIKEEN
                    If TOON.GENDER = GENDER_FEMALE Then
                        Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BAINSHEE)
                    End If
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_CHAMPION)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_ELDRITCH)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_ENCHANTER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_HERO)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_MAULER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_MENTALIST)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_NIGHTSHADE)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_RANGER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_VAMPIIR)
                Case TRH_SHAR
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BLADEMASTER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_CHAMPION)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_HERO)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_MENTALIST)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_RANGER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_VAMPIIR)
                Case TRH_SYLVAN
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_ANIMIST)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_DRUID)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_HERO)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_VALEWALKER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_WARDEN)
                Case TRH_MINOTAUR
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_BLADEMASTER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_HERO)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_MAULER)
                    Combo.AddItem RealmClass(REALM_HIBERNIA, TCH_WARDEN)
            End Select
        Case REALM_MIDGARD
            Select Case tToon.RACE
                Case TRM_DWARF
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BERSERKER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HEALER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HUNTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_RUNEMASTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SAVAGE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHAMAN)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SKALD)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_THANE)
                    If TOON.GENDER = GENDER_FEMALE Then
                        Combo.AddItem RealmClass(REALM_MIDGARD, TCM_VALKYRIE)
                    End If
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARRIOR)
                Case TRM_FROSTALF
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_THANE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HEALER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HUNTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_RUNEMASTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHADOWBLADE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SPIRITMASTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHAMAN)
                    If TOON.GENDER = GENDER_FEMALE Then
                        Combo.AddItem RealmClass(REALM_MIDGARD, TCM_VALKYRIE)
                    End If
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARLOCK)
                Case TRM_KOBOLD
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BONEDANCER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HUNTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_MAULER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_RUNEMASTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SAVAGE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHADOWBLADE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHAMAN)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SKALD)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SPIRITMASTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARLOCK)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARRIOR)
                Case TRM_NORSEMAN
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BERSERKER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HEALER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HUNTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_MAULER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_RUNEMASTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SAVAGE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHADOWBLADE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SKALD)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SPIRITMASTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_THANE)
                    If TOON.GENDER = GENDER_FEMALE Then
                        Combo.AddItem RealmClass(REALM_MIDGARD, TCM_VALKYRIE)
                    End If
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARLOCK)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARRIOR)
                Case TRM_TROLL
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BERSERKER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BONEDANCER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SAVAGE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHAMAN)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SKALD)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_THANE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARRIOR)
                Case TRM_VALKYN
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BERSERKER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BONEDANCER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_HUNTER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SAVAGE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_SHADOWBLADE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARRIOR)
                Case TRM_MINOTAUR
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_BERSERKER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_MAULER)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_THANE)
                    Combo.AddItem RealmClass(REALM_MIDGARD, TCM_WARRIOR)
            End Select
    End Select
    
End Sub

Private Sub Init_RealmRaceArray()
    
    'albion races
    'fighter - armsman, mercenaries, paladin, reaver
    'mage - sorc, cabalist
    'acolyte - friar, cleric
    'rogue - scount, minst, inf
    'elementalist - wizard, theugist
    'disciple - necromancer
    RealmRaces(REALM_ALBION, TRA_AVALONIAN) = "Avalonian"  'armsman, paladin, mercenary, cleric, cabalist, sorcerer, theurgist, wizard, heretic
    RealmRaces(REALM_ALBION, TRA_BRITON) = "Briton"     'armsman, paladin, mercenary, reaver, cleric, friar, infiltrator, minstrel, scout, cabalist, sorcerer, theurgist, wizard, necromancer, heretic
    RealmRaces(REALM_ALBION, TRA_HALFOGRE) = "Half Ogre"  'armsman, mercenary, cabalist, sorcerer, theurgist, wizard
    RealmRaces(REALM_ALBION, TRA_HIGHLANDER) = "Highlander" 'armsman, paladin, mercenary, cleric, minstrel, scout
    RealmRaces(REALM_ALBION, TRA_INCONNU) = "Inconnu"    'armsman, mercenary, reaver, infiltrator, scout, cabalist, sorcerer, necromancer, heretic
    RealmRaces(REALM_ALBION, TRA_SARACEN) = "Saracen"    'armsman, paladin, mercenary, reaver, infiltrator, minstrel, scout, cabalist, sorcerer, necromancer
    RealmRaces(REALM_ALBION, TRA_MINOTAUR) = "Minotaur"   'mauler, ???
   
    'hibernia races
    'naturalist - bard, druid, warden
    'guardian - blademaster, hero, champion
    'magician - eldritch, enchanter, mentalist
    'stalker - nightshade, ranger
    'forester - animist, valewalker
    'vampiir - vampiir
    'bainshee - bainshee
    RealmRaces(REALM_HIBERNIA, TRH_CELT) = "Celt"       'bard, druid, warden, blademaster, hero, champion, mentalist, ranger,animist, valewalker, vampiir, bainshee(F)
    RealmRaces(REALM_HIBERNIA, TRH_ELF) = "Elf"        'blademaster, champion, eldritch, enchanter, mentalist, nightshade, bainshee(F)
    RealmRaces(REALM_HIBERNIA, TRH_FIRBOLG) = "Firbolg"    'bard, druid, warden, blademaster, hero, animist, valewalker
    RealmRaces(REALM_HIBERNIA, TRH_LURIKEEN) = "Lurikeen"   'hero, champion, eldritch, enchanter, mentalist, nightshade, vampiir, bainshee(F)
    RealmRaces(REALM_HIBERNIA, TRH_SHAR) = "Shar"       'blademaster, hero, champion, mentalist, ranger, vampiir
    RealmRaces(REALM_HIBERNIA, TRH_SYLVAN) = "Sylvan"     'druid, warden, hero, animist, valewalker
    RealmRaces(REALM_HIBERNIA, TRH_MINOTAUR) = "Minotaur"   'mauler, ???
    
    'midgard races
    'viking - berserker, skald, thane, warrior, savage
    'seer - healer, shaman
    'rogue - hunter, shadowblade
    'mystic - runemaster, spiritmaster, bonedancer
    'warlock - warlock
    'Valkyrie - Valkyrie
    RealmRaces(REALM_MIDGARD, TRM_DWARF) = "Dwarf"      'berserker, skald, thane, warrior, savage, healer, hunter, runemaster, Valkyrie
    RealmRaces(REALM_MIDGARD, TRM_FROSTALF) = "Frostalf"   'thane, healer, shaman, hunter, shadowblade, runemaster, spiritmaster, warlock, Valkyrie
    RealmRaces(REALM_MIDGARD, TRM_KOBOLD) = "Kobold"     'skald, warrior, savage, shaman, hunter, shadowblade, runemaster, spiritmaster, bonedancer, warlock
    RealmRaces(REALM_MIDGARD, TRM_NORSEMAN) = "Norseman"   'berserker, skald, thane, warrior, savage, healer, hunter, shadowblade, runemaster, spiritmaster, warlock, Valkyrie
    RealmRaces(REALM_MIDGARD, TRM_TROLL) = "Troll"      'berserker, skald, thane, warrior, savage, shaman, bonedancer
    RealmRaces(REALM_MIDGARD, TRM_VALKYN) = "Valkyn"     'berserker, warrior, savage, hunter, shadowblade, bonedancer
    RealmRaces(REALM_MIDGARD, TRM_MINOTAUR) = "Minotaur"   'mauler, ???
    
End Sub

Private Sub Init_SpecialStatsArray()

    SpecialStats(0, 6) = 1
    SpecialStats(0, 7) = 3
    SpecialStats(0, 8) = 4
    SpecialStats(0, 9) = 6
    SpecialStats(0, 10) = 8
    SpecialStats(0, 11) = 10
    SpecialStats(0, 12) = 12
    SpecialStats(0, 13) = 14
    SpecialStats(0, 14) = 16
    SpecialStats(0, 15) = 18
    SpecialStats(0, 16) = 20
    SpecialStats(0, 17) = 23
    SpecialStats(0, 18) = 25
    SpecialStats(0, 19) = 28
    SpecialStats(0, 20) = 30
    SpecialStats(0, 21) = 33
    SpecialStats(0, 22) = 36
    SpecialStats(0, 23) = 38
    SpecialStats(0, 24) = 41
    SpecialStats(0, 25) = 45
    SpecialStats(0, 26) = 47
    SpecialStats(0, 27) = 50
    SpecialStats(0, 28) = 53
    SpecialStats(0, 29) = 57
    SpecialStats(0, 30) = 60
    SpecialStats(0, 31) = 64
    SpecialStats(0, 32) = 68
    SpecialStats(0, 33) = 73
    SpecialStats(0, 34) = 77
    SpecialStats(0, 35) = 81
    SpecialStats(0, 36) = 86
    SpecialStats(0, 37) = 91
    SpecialStats(0, 38) = 96
    SpecialStats(0, 39) = 100
    SpecialStats(0, 40) = 105
    SpecialStats(0, 41) = 108
    SpecialStats(0, 42) = 111
    SpecialStats(0, 43) = 114
    SpecialStats(0, 44) = 117
    SpecialStats(0, 45) = 120
    SpecialStats(0, 46) = 123
    SpecialStats(0, 47) = 126
    SpecialStats(0, 48) = 129
    SpecialStats(0, 49) = 132
    SpecialStats(0, 50) = 135
    
    SpecialStats(1, 6) = 1
    SpecialStats(1, 7) = 2
    SpecialStats(1, 8) = 3
    SpecialStats(1, 9) = 4
    SpecialStats(1, 10) = 5
    SpecialStats(1, 11) = 6
    SpecialStats(1, 12) = 8
    SpecialStats(1, 13) = 9
    SpecialStats(1, 14) = 10
    SpecialStats(1, 15) = 12
    SpecialStats(1, 16) = 13
    SpecialStats(1, 17) = 15
    SpecialStats(1, 18) = 17
    SpecialStats(1, 19) = 18
    SpecialStats(1, 20) = 20
    SpecialStats(1, 21) = 22
    SpecialStats(1, 22) = 24
    SpecialStats(1, 23) = 25
    SpecialStats(1, 24) = 27
    SpecialStats(1, 25) = 30
    SpecialStats(1, 26) = 31
    SpecialStats(1, 27) = 33
    SpecialStats(1, 28) = 35
    SpecialStats(1, 29) = 38
    SpecialStats(1, 30) = 40
    SpecialStats(1, 31) = 43
    SpecialStats(1, 32) = 45
    SpecialStats(1, 33) = 43
    SpecialStats(1, 34) = 51
    SpecialStats(1, 35) = 54
    SpecialStats(1, 36) = 57
    SpecialStats(1, 37) = 60
    SpecialStats(1, 38) = 64
    SpecialStats(1, 39) = 67
    SpecialStats(1, 40) = 70
    SpecialStats(1, 41) = 72
    SpecialStats(1, 42) = 74
    SpecialStats(1, 43) = 76
    SpecialStats(1, 44) = 78
    SpecialStats(1, 45) = 80
    SpecialStats(1, 46) = 82
    SpecialStats(1, 47) = 84
    SpecialStats(1, 48) = 86
    SpecialStats(1, 49) = 88
    SpecialStats(1, 50) = 90
End Sub

Private Sub Init_RealmClassArray()

    'Albion
    RealmClass(REALM_ALBION, TCA_ARMSMAN) = "Armsman"
    RealmClass(REALM_ALBION, TCA_CABALIST) = "Cabalist"
    RealmClass(REALM_ALBION, TCA_CLERIC) = "Cleric"
    RealmClass(REALM_ALBION, TCA_FRIAR) = "Friar"
    RealmClass(REALM_ALBION, TCA_HERETIC) = "Heretic"
    RealmClass(REALM_ALBION, TCA_INFILTRATOR) = "Infiltrator"
    RealmClass(REALM_ALBION, TCA_MERCENARY) = "Mercenary"
    RealmClass(REALM_ALBION, TCA_MINSTREL) = "Minstrel"
    RealmClass(REALM_ALBION, TCA_NECROMANCER) = "Necromancer"
    RealmClass(REALM_ALBION, TCA_PALADIN) = "Paladin"
    RealmClass(REALM_ALBION, TCA_REAVER) = "Reaver"
    RealmClass(REALM_ALBION, TCA_SCOUT) = "Scout"
    RealmClass(REALM_ALBION, TCA_SORCERER) = "Sorcerer"
    RealmClass(REALM_ALBION, TCA_THEURGIST) = "Theurgist"
    RealmClass(REALM_ALBION, TCA_WIZARD) = "Wizard"
    RealmClass(REALM_ALBION, TCA_MAULER) = "Mauler"
    
    'Hibernia
    RealmClass(REALM_HIBERNIA, TCH_ANIMIST) = "Animist"
    RealmClass(REALM_HIBERNIA, TCH_BAINSHEE) = "Bainshee"
    RealmClass(REALM_HIBERNIA, TCH_BARD) = "Bard"
    RealmClass(REALM_HIBERNIA, TCH_BLADEMASTER) = "Blademaster"
    RealmClass(REALM_HIBERNIA, TCH_CHAMPION) = "Champion"
    RealmClass(REALM_HIBERNIA, TCH_DRUID) = "Druid"
    RealmClass(REALM_HIBERNIA, TCH_ELDRITCH) = "Eldritch"
    RealmClass(REALM_HIBERNIA, TCH_ENCHANTER) = "Enchanter"
    RealmClass(REALM_HIBERNIA, TCH_HERO) = "Hero"
    RealmClass(REALM_HIBERNIA, TCH_MENTALIST) = "Mentalist"
    RealmClass(REALM_HIBERNIA, TCH_NIGHTSHADE) = "Nightshade"
    RealmClass(REALM_HIBERNIA, TCH_RANGER) = "Ranger"
    RealmClass(REALM_HIBERNIA, TCH_VALEWALKER) = "Valewalker"
    RealmClass(REALM_HIBERNIA, TCH_VAMPIIR) = "Vampiir"
    RealmClass(REALM_HIBERNIA, TCH_WARDEN) = "Warden"
    RealmClass(REALM_HIBERNIA, TCH_MAULER) = "Mauler"
    
    'Midgard
    RealmClass(REALM_MIDGARD, TCM_BERSERKER) = "Berserker"
    RealmClass(REALM_MIDGARD, TCM_BONEDANCER) = "Bonedancer"
    RealmClass(REALM_MIDGARD, TCM_HEALER) = "Healer"
    RealmClass(REALM_MIDGARD, TCM_HUNTER) = "Hunter"
    RealmClass(REALM_MIDGARD, TCM_RUNEMASTER) = "Runemaster"
    RealmClass(REALM_MIDGARD, TCM_SAVAGE) = "Savage"
    RealmClass(REALM_MIDGARD, TCM_SHADOWBLADE) = "Shadowblade"
    RealmClass(REALM_MIDGARD, TCM_SHAMAN) = "Shaman"
    RealmClass(REALM_MIDGARD, TCM_SKALD) = "Skald"
    RealmClass(REALM_MIDGARD, TCM_SPIRITMASTER) = "Spiritmaster"
    RealmClass(REALM_MIDGARD, TCM_THANE) = "Thane"
    RealmClass(REALM_MIDGARD, TCM_VALKYRIE) = "Valkyrie"
    RealmClass(REALM_MIDGARD, TCM_WARLOCK) = "Warlock"
    RealmClass(REALM_MIDGARD, TCM_WARRIOR) = "Warrior"
    RealmClass(REALM_MIDGARD, TCM_MAULER) = "Mauler"
    
End Sub

Public Sub InitializeClass(CLASS As String, ByRef tToon As TOON_TYPE)

    Select Case tToon.REALM
        Case REALM_ALBION
            Select Case LCase$(CLASS)
                Case "armsman"      'id = 0
                    Call Init_ARMSMAN(tToon)
                Case "cabalist"     'id = 1
                    Call Init_CABALIST(tToon)
                Case "cleric"       'id = 2
                    Call Init_CLERIC(tToon)
                Case "friar"        'id = 3
                    Call Init_FRIAR(tToon)
                Case "heretic"      'id = 4
                    Call Init_HERETIC(tToon)
                Case "infiltrator"  'id = 5
                    Call Init_INFILTRATOR(tToon)
                Case "mercenary"    'id = 6
                    Call Init_MERCENARY(tToon)
                Case "minstrel"     'id = 7
                    Call Init_MINSTREL(tToon)
                Case "necromancer"  'id = 8
                    Call Init_NECROMANCER(tToon)
                Case "paladin"      'id = 9
                    Call Init_PALADIN(tToon)
                Case "reaver"       'id = 10
                    Call Init_REAVER(tToon)
                Case "scout"        'id = 11
                    Call Init_SCOUT(tToon)
                Case "sorcerer"     'id = 12
                    Call Init_SORCERER(tToon)
                Case "theurgist"    'id = 13
                    Call Init_THEURGIST(tToon)
                Case "wizard"       'id = 14
                    Call Init_WIZARD(tToon)
                Case "mauler"
                    Call Init_MAULER_ALB(tToon)
            End Select
        Case REALM_HIBERNIA
            Select Case LCase$(CLASS)
                Case "animist"      'id = 0
                    Call Init_ANIMIST(tToon)
                Case "bainshee"     'id = 1??
                    Call Init_BAINSHEE(tToon)
                Case "bard"         'id = 2
                    Call Init_BARD(tToon)
                Case "blademaster"  'id = 3
                    Call Init_BLADEMASTER(tToon)
                Case "champion"     'id = 4
                    Call Init_CHAMPION(tToon)
                Case "druid"        'id = 5
                    Call Init_DRUID(tToon)
                Case "eldritch"     'id = 6
                    Call Init_ELDRITCH(tToon)
                Case "enchanter"    'id = 7
                    Call Init_ENCHANTER(tToon)
                Case "hero"         'id = 8
                    Call Init_HERO(tToon)
                Case "mentalist"    'id = 9
                    Call Init_MENTALIST(tToon)
                Case "nightshade"   'id = 10
                    Call Init_NIGHTSHADE(tToon)
                Case "ranger"       'id = 11
                    Call Init_RANGER(tToon)
                Case "valewalker"   'id = 12
                    Call Init_VALEWALKER(tToon)
                Case "vampiir"      'id = 13
                    Call Init_VAMPIIR(tToon)
                Case "warden"       'id = 14
                    Call Init_WARDEN(tToon)
                Case "mauler"
                    Call Init_MAULER_HIB(tToon)
            End Select
        Case REALM_MIDGARD
            Select Case LCase$(CLASS)
                Case "berserker"
                    Call Init_BERSERKER(tToon)
                Case "bonedancer"
                    Call Init_BONEDANCER(tToon)
                Case "healer"
                    Call Init_HEALER(tToon)
                Case "hunter"
                    Call Init_HUNTER(tToon)
                Case "runemaster"
                    Call Init_RUNEMASTER(tToon)
                Case "savage"
                    Call Init_SAVAGE(tToon)
                Case "shadowblade"
                    Call Init_SHADOWBLADE(tToon)
                Case "shaman"
                    Call Init_SHAMAN(tToon)
                Case "skald"
                    Call Init_SKALD(tToon)
                Case "spiritmaster"
                    Call Init_SPIRITMASTER(tToon)
                Case "thane"
                    Call Init_THANE(tToon)
                Case "valkyrie"
                    Call Init_VALKYRIE(tToon)
                Case "warlock"
                    Call Init_WARLOCK(tToon)
                Case "warrior"
                    Call Init_WARRIOR(tToon)
                Case "mauler"
                    Call Init_MAULER_MID(tToon)
            End Select
    End Select
        
End Sub

Public Sub AssignRaceAttributes(ByRef tToon As TOON_TYPE)
    
    Select Case tToon.REALM
        Case REALM_ALBION
            Select Case tToon.RACE
                Case TRA_AVALONIAN
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 45    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 45    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 60    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 70    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 80    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_MATTER_RESIST, SM_LOC_BASE) = 5    'set matter
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 2    'set crush
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 3    'set slash
                Case TRA_BRITON
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 60    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 60    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 60    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 60    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_SPIRIT_RESIST, SM_LOC_BASE) = 5    'set spirit
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 2    'set crush
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 3    'set slash
                Case TRA_HALFOGRE
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 90    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 70    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 40    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 40    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_MATTER_RESIST, SM_LOC_BASE) = 5    'set matter
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 2    'set thrust
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 3    'set slash
                Case TRA_HIGHLANDER
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 70    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 70    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 50    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 50    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                
                    tToon.STAT_MATRIX(SM_COLD_RESIST, SM_LOC_BASE) = 5    'set cold
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 3    'set crush
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 2    'set slash
                Case TRA_INCONNU
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 50    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 60    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 70    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 50    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 70    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha

                    tToon.STAT_MATRIX(SM_HEAT_RESIST, SM_LOC_BASE) = 5    'set heat
                    tToon.STAT_MATRIX(SM_SPIRIT_RESIST, SM_LOC_BASE) = 5    'set spirit
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 2    'set crush
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 3    'set thrust
                Case TRA_SARACEN
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 50    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 50    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 80    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 60    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
            
                    tToon.STAT_MATRIX(SM_HEAT_RESIST, SM_LOC_BASE) = 5    'set heat
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 3    'set thrust
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 2    'set slash
                Case TRA_MINOTAUR
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 80    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 70    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 50    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 40    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
            
                    tToon.STAT_MATRIX(SM_HEAT_RESIST, SM_LOC_BASE) = 3      'set heat
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 4     'set crush
                    tToon.STAT_MATRIX(SM_COLD_RESIST, SM_LOC_BASE) = 3      'cold resist
            End Select
        Case REALM_HIBERNIA
            Select Case tToon.RACE
                Case TRH_CELT
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 60    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 60    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 60    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 60    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                
                    tToon.STAT_MATRIX(SM_SPIRIT_RESIST, SM_LOC_BASE) = 5    'set spirit
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 2    'set crush
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 3    'set slash
                Case TRH_ELF
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 40    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 40    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 75    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 75    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 70    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_SPIRIT_RESIST, SM_LOC_BASE) = 5    'set spirit
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 3    'set thrust
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 2    'set slash
                Case TRH_FIRBOLG
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 90    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 60    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 40    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 40    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 70    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_HEAT_RESIST, SM_LOC_BASE) = 5    'set heat
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 3    'set crush
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 2    'set slash
                Case TRH_LURIKEEN
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 40    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 40    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 80    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 80    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha

                    tToon.STAT_MATRIX(SM_ENERGY_RESIST, SM_LOC_BASE) = 5    'set energy
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 5    'set crush
                Case TRH_SHAR
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 60    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 80    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 50    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 50    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_ENERGY_RESIST, SM_LOC_BASE) = 5    'set energy
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 5    'set crush
                Case TRH_SYLVAN
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 70    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 60    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 55    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 45    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 70    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_ENERGY_RESIST, SM_LOC_BASE) = 5    'set energy
                    tToon.STAT_MATRIX(SM_MATTER_RESIST, SM_LOC_BASE) = 5    'set matter
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 3    'set crush
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 2    'set thrust
                Case TRH_MINOTAUR
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 80    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 70    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 50    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 40    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
            
                    tToon.STAT_MATRIX(SM_HEAT_RESIST, SM_LOC_BASE) = 3      'set heat
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 4     'set crush
                    tToon.STAT_MATRIX(SM_COLD_RESIST, SM_LOC_BASE) = 3      'cold resist
            End Select
        Case REALM_MIDGARD
            Select Case tToon.RACE
                Case TRM_DWARF
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 60    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 80    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 50    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 50    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_BODY_RESIST, SM_LOC_BASE) = 5    'set body
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 3    'set thrust
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 2    'set slash
                Case TRM_FROSTALF
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 55    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 55    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 55    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 60    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 75    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_SPIRIT_RESIST, SM_LOC_BASE) = 5    'set spirit
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 3    'set thrust
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 2    'set slash
                Case TRM_KOBOLD
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 50    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 50    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 70    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 70    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_ENERGY_RESIST, SM_LOC_BASE) = 5    'set energy
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 5    'set crush
                Case TRM_NORSEMAN
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 70    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 70    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 50    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 50    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_COLD_RESIST, SM_LOC_BASE) = 5    'set cold
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 2    'set crush
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 3    'set slash
                Case TRM_TROLL
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 100   'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 70    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 35    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 35    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_MATTER_RESIST, SM_LOC_BASE) = 5    'set matter
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 2    'set thrust
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 3    'set slash
                Case TRM_VALKYN
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 55    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 45    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 65    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 75    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
                    
                    tToon.STAT_MATRIX(SM_BODY_RESIST, SM_LOC_BASE) = 5    'set body
                    tToon.STAT_MATRIX(SM_COLD_RESIST, SM_LOC_BASE) = 5    'set cold
                    tToon.STAT_MATRIX(SM_THRUST_RESIST, SM_LOC_BASE) = 2    'set thrust
                    tToon.STAT_MATRIX(SM_SLASH_RESIST, SM_LOC_BASE) = 3    'set slash
                Case TRM_MINOTAUR
                    tToon.STAT_MATRIX(SM_STR, SM_LOC_BASE) = 80    'set str
                    tToon.STAT_MATRIX(SM_CON, SM_LOC_BASE) = 70    'set con
                    tToon.STAT_MATRIX(SM_DEX, SM_LOC_BASE) = 50    'set dex
                    tToon.STAT_MATRIX(SM_QUI, SM_LOC_BASE) = 40    'set qui
                    tToon.STAT_MATRIX(SM_INT, SM_LOC_BASE) = 60    'set int
                    tToon.STAT_MATRIX(SM_EMP, SM_LOC_BASE) = 60    'set emp
                    tToon.STAT_MATRIX(SM_PIE, SM_LOC_BASE) = 60    'set pie
                    tToon.STAT_MATRIX(SM_CHA, SM_LOC_BASE) = 60    'set cha
            
                    tToon.STAT_MATRIX(SM_HEAT_RESIST, SM_LOC_BASE) = 3      'set heat
                    tToon.STAT_MATRIX(SM_CRUSH_RESIST, SM_LOC_BASE) = 4     'set crush
                    tToon.STAT_MATRIX(SM_COLD_RESIST, SM_LOC_BASE) = 3      'cold resist
            End Select
    End Select
End Sub
        
