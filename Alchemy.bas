Attribute VB_Name = "Alchemy"
Option Explicit

Private Type ProcChargeReactive
    AlchType As String 'proc, charge, reactive, drop
    AlchDetails(3, 40) As String
End Type

Public AlchemyEffects(3) As ProcChargeReactive

Public Const AE_PROCS = 0
Public Const AE_CHARGES = 1
Public Const AE_REACTIVES = 2
Public Const AE_DROPS = 3

Public Const AD_NAME = 0
Public Const AD_COST = 1
Public Const AD_EFFECT = 2
Public Const AD_LEVEL = 3

Public Sub InitAlchemyArrays(ProgressBar As Control, StatusBar As Control)

'Index 0 will be Procs
'AlchDetails(ad_name,x) = name
'AlchDetails(AD_COST,x) = effect
'AlchDetails(AD_EFFECT,x) = price
'AlchDetails(AD_LEVEL,x) = level

StatusBar.Caption = "Initialzing Alchemy Effects"

With AlchemyEffects(AE_PROCS)
    
    StatusBar.Caption = "Initialzing Procs"
    DoEvents
    .AlchType = "Procs"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 0) = "Volatile Fire Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 0) = "14180"
    .AlchDetails(AD_EFFECT, 0) = "41 Heat"
    .AlchDetails(AD_LEVEL, 0) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 1) = "Volatile Cold Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 1) = "14180"
    .AlchDetails(AD_EFFECT, 1) = "41 Cold"
    .AlchDetails(AD_LEVEL, 1) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 2) = "Volatile Energy Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 2) = "14180"
    .AlchDetails(AD_EFFECT, 2) = "41 Energy"
    .AlchDetails(AD_LEVEL, 2) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 3) = "Volatile Spirit Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 3) = "14180"
    .AlchDetails(AD_EFFECT, 3) = "41 Spirit"
    .AlchDetails(AD_LEVEL, 3) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 4) = "Volatile Fire Fine Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 4) = "27540"
    .AlchDetails(AD_EFFECT, 4) = "50 Heat"
    .AlchDetails(AD_LEVEL, 4) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 5) = "Volatile Cold Fine Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 5) = "27540"
    .AlchDetails(AD_EFFECT, 5) = "50 Cold"
    .AlchDetails(AD_LEVEL, 5) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 6) = "Volatile Energy Fine Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 6) = "27540"
    .AlchDetails(AD_EFFECT, 6) = "50 Energy"
    .AlchDetails(AD_LEVEL, 6) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 7) = "Volatile Spirit Fine Alloy Weapon Tincture"
    .AlchDetails(AD_COST, 7) = "27540"
    .AlchDetails(AD_EFFECT, 7) = "50 Spirit"
    .AlchDetails(AD_LEVEL, 7) = "25"
        
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 8) = "Volatile Fire Mithril Weapon Tincture"
    .AlchDetails(AD_COST, 8) = "52580"
    .AlchDetails(AD_EFFECT, 8) = "59 Heat"
    .AlchDetails(AD_LEVEL, 8) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 9) = "Volatile Cold Mithril Weapon Tincture"
    .AlchDetails(AD_COST, 9) = "52580"
    .AlchDetails(AD_EFFECT, 9) = "59 Cold"
    .AlchDetails(AD_LEVEL, 9) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 10) = "Volatile Energy Mithril Weapon Tincture"
    .AlchDetails(AD_COST, 10) = "52580"
    .AlchDetails(AD_EFFECT, 10) = "59 Energy"
    .AlchDetails(AD_LEVEL, 10) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 11) = "Volatile Spirit Mithril Weapon Tincture"
    .AlchDetails(AD_COST, 11) = "52580"
    .AlchDetails(AD_EFFECT, 11) = "59 Spirit"
    .AlchDetails(AD_LEVEL, 11) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 12) = "Volatile Fire Adamantium Weapon Tincture"
    .AlchDetails(AD_COST, 12) = "97620"
    .AlchDetails(AD_EFFECT, 12) = "68 Heat"
    .AlchDetails(AD_LEVEL, 12) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 13) = "Volatile Cold Adamantium Weapon Tincture"
    .AlchDetails(AD_COST, 13) = "97620"
    .AlchDetails(AD_EFFECT, 13) = "68 Cold"
    .AlchDetails(AD_LEVEL, 13) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 14) = "Volatile Energy Adamantium Weapon Tincture"
    .AlchDetails(AD_COST, 14) = "97620"
    .AlchDetails(AD_EFFECT, 14) = "68 Energy"
    .AlchDetails(AD_LEVEL, 14) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 15) = "Volatile Spirit Adamantium Weapon Tincture"
    .AlchDetails(AD_COST, 15) = "97620"
    .AlchDetails(AD_EFFECT, 15) = "68 Spirit"
    .AlchDetails(AD_LEVEL, 15) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 16) = "Volatile Fire Asterite Weapon Tincture"
    .AlchDetails(AD_COST, 16) = "185180"
    .AlchDetails(AD_EFFECT, 16) = "77 Heat"
    .AlchDetails(AD_LEVEL, 16) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 17) = "Volatile Cold Asterite Weapon Tincture"
    .AlchDetails(AD_COST, 17) = "185180"
    .AlchDetails(AD_EFFECT, 17) = "77 Cold"
    .AlchDetails(AD_LEVEL, 17) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 18) = "Volatile Energy Asterite Weapon Tincture"
    .AlchDetails(AD_COST, 18) = "185180"
    .AlchDetails(AD_EFFECT, 18) = "77 Energy"
    .AlchDetails(AD_LEVEL, 18) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 19) = "Volatile Spirit Asterite Weapon Tincture"
    .AlchDetails(AD_COST, 19) = "185180"
    .AlchDetails(AD_EFFECT, 19) = "77 Spirit"
    .AlchDetails(AD_LEVEL, 19) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 20) = "Volatile Fire Netherium Weapon Tincture"
    .AlchDetails(AD_COST, 20) = "356520"
    .AlchDetails(AD_EFFECT, 20) = "86 Heat"
    .AlchDetails(AD_LEVEL, 20) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 21) = "Volatile Cold Netherium Weapon Tincture"
    .AlchDetails(AD_COST, 21) = "356520"
    .AlchDetails(AD_EFFECT, 21) = "86 Cold"
    .AlchDetails(AD_LEVEL, 21) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 22) = "Volatile Energy Netherium Weapon Tincture"
    .AlchDetails(AD_COST, 22) = "356520"
    .AlchDetails(AD_EFFECT, 22) = "86 Energy"
    .AlchDetails(AD_LEVEL, 22) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 23) = "Volatile Spirit Netherium Weapon Tincture"
    .AlchDetails(AD_COST, 23) = "356520"
    .AlchDetails(AD_EFFECT, 23) = "86 Spirit"
    .AlchDetails(AD_LEVEL, 23) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 24) = "Volatile Provoking Netherium Weapon Tincture"
    .AlchDetails(AD_COST, 24) = "739020"
    .AlchDetails(AD_EFFECT, 24) = "Taunt 1"
    .AlchDetails(AD_LEVEL, 24) = "45"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 25) = "Volatile Depletion Netherium Weapon Tincture"
    .AlchDetails(AD_COST, 25) = "739020"
    .AlchDetails(AD_EFFECT, 25) = "Power Drain: 35 Dmg: 50%"
    .AlchDetails(AD_LEVEL, 25) = "45"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 26) = "Volatile Ablative Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 26) = "765030"
    .AlchDetails(AD_EFFECT, 26) = "Self Melee Health Buffer: 50"
    .AlchDetails(AD_LEVEL, 26) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 27) = "Volatile Hardening Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 27) = "765030"
    .AlchDetails(AD_EFFECT, 27) = "Self AF Buff: 75"
    .AlchDetails(AD_LEVEL, 27) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 28) = "Volatile Eroding Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 28) = "751816"
    .AlchDetails(AD_EFFECT, 28) = "DoT: 64/Tick"
    .AlchDetails(AD_LEVEL, 28) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 29) = "Volatile Cleric Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 29) = "776530"
    .AlchDetails(AD_EFFECT, 29) = "Self Haste: 20%"
    .AlchDetails(AD_LEVEL, 29) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 30) = "Volatile Shard Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 30) = "768530"
    .AlchDetails(AD_EFFECT, 30) = "Self Dmg Shield: 5 DPS"
    .AlchDetails(AD_LEVEL, 30) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 31) = "Volatile Fire Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 31) = "693930"
    .AlchDetails(AD_EFFECT, 31) = "95 Heat"
    .AlchDetails(AD_LEVEL, 31) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 32) = "Volatile Cold Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 32) = "693930"
    .AlchDetails(AD_EFFECT, 32) = "95 Cold"
    .AlchDetails(AD_LEVEL, 32) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 33) = "Volatile Energy Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 33) = "693930"
    .AlchDetails(AD_EFFECT, 33) = "95 Energy"
    .AlchDetails(AD_LEVEL, 33) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 34) = "Volatile Spirit Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 34) = "693930"
    .AlchDetails(AD_EFFECT, 34) = "95 Spirit"
    .AlchDetails(AD_LEVEL, 34) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 35) = "Volatile Provoking Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 35) = "916030"
    .AlchDetails(AD_EFFECT, 35) = "Taunt 2"
    .AlchDetails(AD_LEVEL, 35) = "49"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 36) = "Volatile Depletion Arcanium Weapon Tincture"
    .AlchDetails(AD_COST, 36) = "916030"
    .AlchDetails(AD_EFFECT, 36) = "Power Drain: 55 Dmg: 50%"
    .AlchDetails(AD_LEVEL, 36) = "49"
End With    'AlchemyEffects(0)

'Index 1 is Charges
With AlchemyEffects(AE_CHARGES)
    
    StatusBar.Caption = "Initializing Charges"
    DoEvents
    .AlchType = "Charges"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 0) = "Stable Fire Alloy Tincture"
    .AlchDetails(AD_COST, 0) = "11680"
    .AlchDetails(AD_EFFECT, 0) = "41 Heat"
    .AlchDetails(AD_LEVEL, 0) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 1) = "Stable Cold Alloy Tincture"
    .AlchDetails(AD_COST, 1) = "11680"
    .AlchDetails(AD_EFFECT, 1) = "41 Cold"
    .AlchDetails(AD_LEVEL, 1) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 2) = "Stable Energy Alloy Tincture"
    .AlchDetails(AD_COST, 2) = "11680"
    .AlchDetails(AD_EFFECT, 2) = "41 Energy"
    .AlchDetails(AD_LEVEL, 2) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 3) = "Stable Spirit Alloy Tincture"
    .AlchDetails(AD_COST, 3) = "11680"
    .AlchDetails(AD_EFFECT, 3) = "41 Spirit"
    .AlchDetails(AD_LEVEL, 3) = "20"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 4) = "Stable Fire Fine Alloy Tincture"
    .AlchDetails(AD_COST, 4) = "22540"
    .AlchDetails(AD_EFFECT, 4) = "50 Heat"
    .AlchDetails(AD_LEVEL, 4) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 5) = "Stable Cold Fine Alloy Tincture"
    .AlchDetails(AD_COST, 5) = "22540"
    .AlchDetails(AD_EFFECT, 5) = "50 Cold"
    .AlchDetails(AD_LEVEL, 5) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 6) = "Stable Energy Fine Alloy Tincture"
    .AlchDetails(AD_COST, 6) = "22540"
    .AlchDetails(AD_EFFECT, 6) = "50 Energy"
    .AlchDetails(AD_LEVEL, 6) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 7) = "Stable Spirit Fine Alloy Tincture"
    .AlchDetails(AD_COST, 7) = "22540"
    .AlchDetails(AD_EFFECT, 7) = "50 Spirit"
    .AlchDetails(AD_LEVEL, 7) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 8) = "Stable Fire Mithril Tincture"
    .AlchDetails(AD_COST, 8) = "42580"
    .AlchDetails(AD_EFFECT, 8) = "59 Heat"
    .AlchDetails(AD_LEVEL, 8) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 9) = "Stable Cold Mithril Tincture"
    .AlchDetails(AD_COST, 9) = "42580"
    .AlchDetails(AD_EFFECT, 9) = "59 Cold"
    .AlchDetails(AD_LEVEL, 9) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 10) = "Stable Energy Mithril Tincture"
    .AlchDetails(AD_COST, 10) = "42580"
    .AlchDetails(AD_EFFECT, 10) = "59 Energy"
    .AlchDetails(AD_LEVEL, 10) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 11) = "Stable Spirit Mithril Tincture"
    .AlchDetails(AD_COST, 11) = "42580"
    .AlchDetails(AD_EFFECT, 11) = "59 Spirit"
    .AlchDetails(AD_LEVEL, 11) = "30"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 12) = "Stable Fire Adamantium Tincture"
    .AlchDetails(AD_COST, 12) = "77620"
    .AlchDetails(AD_EFFECT, 12) = "68 Heat"
    .AlchDetails(AD_LEVEL, 12) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 13) = "Stable Cold Adamantium Tincture"
    .AlchDetails(AD_COST, 13) = "77620"
    .AlchDetails(AD_EFFECT, 13) = "68 Cold"
    .AlchDetails(AD_LEVEL, 13) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 14) = "Stable Energy Adamantium Tincture"
    .AlchDetails(AD_COST, 14) = "77620"
    .AlchDetails(AD_EFFECT, 14) = "68 Energy"
    .AlchDetails(AD_LEVEL, 14) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 15) = "Stable Spirit Adamantium Tincture"
    .AlchDetails(AD_COST, 15) = "77620"
    .AlchDetails(AD_EFFECT, 15) = "68 Spirit"
    .AlchDetails(AD_LEVEL, 15) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 16) = "Stable Fire Asterite Tincture"
    .AlchDetails(AD_COST, 16) = "145180"
    .AlchDetails(AD_EFFECT, 16) = "77 Heat"
    .AlchDetails(AD_LEVEL, 16) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 17) = "Stable Cold Asterite Tincture"
    .AlchDetails(AD_COST, 17) = "145180"
    .AlchDetails(AD_EFFECT, 17) = "77 Cold"
    .AlchDetails(AD_LEVEL, 17) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 18) = "Stable Energy Asterite Tincture"
    .AlchDetails(AD_COST, 18) = "145180"
    .AlchDetails(AD_EFFECT, 18) = "77 Energy"
    .AlchDetails(AD_LEVEL, 18) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 19) = "Stable Spirit Asterite Tincture"
    .AlchDetails(AD_COST, 19) = "145180"
    .AlchDetails(AD_EFFECT, 19) = "77 Spirit"
    .AlchDetails(AD_LEVEL, 19) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 20) = "Stable Fire Netherium Tincture"
    .AlchDetails(AD_COST, 20) = "276520"
    .AlchDetails(AD_EFFECT, 20) = "86 Heat"
    .AlchDetails(AD_LEVEL, 20) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 21) = "Stable Cold Netherium Tincture"
    .AlchDetails(AD_COST, 21) = "276520"
    .AlchDetails(AD_EFFECT, 21) = "86 Cold"
    .AlchDetails(AD_LEVEL, 21) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 22) = "Stable Energy Netherium Tincture"
    .AlchDetails(AD_COST, 22) = "276520"
    .AlchDetails(AD_EFFECT, 22) = "86 Energy"
    .AlchDetails(AD_LEVEL, 22) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 23) = "Stable Spirit Netherium Tincture"
    .AlchDetails(AD_COST, 23) = "276520"
    .AlchDetails(AD_EFFECT, 23) = "86 Spirit"
    .AlchDetails(AD_LEVEL, 23) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 24) = "Stable Fire Arcanium Tincture"
    .AlchDetails(AD_COST, 24) = "533530"
    .AlchDetails(AD_EFFECT, 24) = "95 Heat"
    .AlchDetails(AD_LEVEL, 24) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 25) = "Stable Cold Arcanium Tincture"
    .AlchDetails(AD_COST, 25) = "533530"
    .AlchDetails(AD_EFFECT, 25) = "95 Cold"
    .AlchDetails(AD_LEVEL, 25) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 26) = "Stable Energy Arcanium Tincture"
    .AlchDetails(AD_COST, 26) = "533530"
    .AlchDetails(AD_EFFECT, 26) = "95 Energy"
    .AlchDetails(AD_LEVEL, 26) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 27) = "Stable Spirit Arcanium Tincture"
    .AlchDetails(AD_COST, 27) = "533530"
    .AlchDetails(AD_EFFECT, 27) = "95 Spirit"
    .AlchDetails(AD_LEVEL, 27) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 28) = "Stable Ablative Arcanium Tincture"
    .AlchDetails(AD_COST, 28) = "615030"
    .AlchDetails(AD_EFFECT, 28) = "50"
    .AlchDetails(AD_LEVEL, 28) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 29) = "Stable Hardening Arcanium Tincture"
    .AlchDetails(AD_COST, 29) = "615030"
    .AlchDetails(AD_EFFECT, 29) = "Armor Factor: 75"
    .AlchDetails(AD_LEVEL, 29) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 30) = "Stable Enlightening Arcanium Tincture"
    .AlchDetails(AD_COST, 30) = "607030"
    .AlchDetails(AD_EFFECT, 30) = "Acuity: 75"
    .AlchDetails(AD_LEVEL, 30) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 31) = "Stable Eroding Arcanium Tincture"
    .AlchDetails(AD_COST, 31) = "607316"
    .AlchDetails(AD_EFFECT, 31) = "64 per tick (Matter)"
    .AlchDetails(AD_LEVEL, 31) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 32) = "Stable Celeric Arcanium Tincture"
    .AlchDetails(AD_COST, 32) = "626530"
    .AlchDetails(AD_EFFECT, 32) = "17% Self Haste"
    .AlchDetails(AD_LEVEL, 32) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 33) = "Stable Shard Arcanium Tincture"
    .AlchDetails(AD_COST, 33) = "624030"
    .AlchDetails(AD_EFFECT, 33) = "Shield: 4 Dps"
    .AlchDetails(AD_LEVEL, 33) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 34) = "Stable Honing Arcanium Tincture"
    .AlchDetails(AD_COST, 34) = "818530"
    .AlchDetails(AD_EFFECT, 34) = "Damage Add: 11 Dps (Matter)"
    .AlchDetails(AD_LEVEL, 34) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 35) = "Stable Leeching Arcanium Tincture"
    .AlchDetails(AD_COST, 35) = "816030"
    .AlchDetails(AD_EFFECT, 35) = "Lifedrain: 56"
    .AlchDetails(AD_LEVEL, 35) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 36) = "Stable Withering Arcanium Tincture"
    .AlchDetails(AD_COST, 36) = "815030"
    .AlchDetails(AD_EFFECT, 36) = "Str/Con Debuff: 56"
    .AlchDetails(AD_LEVEL, 36) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 37) = "Stable Crippling Arcanium Tincture"
    .AlchDetails(AD_COST, 37) = "815030"
    .AlchDetails(AD_EFFECT, 37) = "Dex/Qui Debuff: 56"
    .AlchDetails(AD_LEVEL, 37) = "47"
        
End With    'alchemyeffects(ae_charges)

'Index 2 is Reactives
With AlchemyEffects(AE_REACTIVES)

    StatusBar.Caption = "Initializing Reactives"
    DoEvents
    .AlchType = "Reactives"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 0) = "Reactive Shard Fine Alloy Armor Tincture"
    .AlchDetails(AD_COST, 0) = "510040"
    .AlchDetails(AD_EFFECT, 0) = "Self  Shield: 2.6 DPS"
    .AlchDetails(AD_LEVEL, 0) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 1) = "Reactive Hardening Fine Alloy Armor Tincture"
    .AlchDetails(AD_COST, 1) = "510040"
    .AlchDetails(AD_EFFECT, 1) = "Self AF Buff: 37"
    .AlchDetails(AD_LEVEL, 1) = "25"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 2) = "Reactive Ablative Fine Alloy Armor Tincture"
    .AlchDetails(AD_COST, 2) = "510040"
    .AlchDetails(AD_EFFECT, 2) = "Self Melee Health Buffer: 50"
    .AlchDetails(AD_LEVEL, 2) = ""
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 3) = "Reactive Shard Adamantium Armor Tincture"
    .AlchDetails(AD_COST, 3) = "710120"
    .AlchDetails(AD_EFFECT, 3) = "Self  Shield: 3.6 DPS"
    .AlchDetails(AD_LEVEL, 3) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 4) = "Reactive Hardening Adamantium Armor Tincture"
    .AlchDetails(AD_COST, 4) = "710120"
    .AlchDetails(AD_EFFECT, 4) = "Self AF Buff: 56"
    .AlchDetails(AD_LEVEL, 4) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 5) = "Reactive Ablative Adamantium Armor Tincture"
    .AlchDetails(AD_COST, 5) = "710120"
    .AlchDetails(AD_EFFECT, 5) = "Self Melee Health Buffer: 75"
    .AlchDetails(AD_LEVEL, 5) = "35"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 6) = "Reactive Fire Asterite Armor Tincture"
    .AlchDetails(AD_COST, 6) = "707860"
    .AlchDetails(AD_EFFECT, 6) = "77 Heat"
    .AlchDetails(AD_LEVEL, 6) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 7) = "Reactive Cold Asterite Armor Tincture"
    .AlchDetails(AD_COST, 7) = "707860"
    .AlchDetails(AD_EFFECT, 7) = "77 Cold"
    .AlchDetails(AD_LEVEL, 7) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 8) = "Reactive Energy Asterite Armor Tincture"
    .AlchDetails(AD_COST, 8) = "707860"
    .AlchDetails(AD_EFFECT, 8) = "77 Energy"
    .AlchDetails(AD_LEVEL, 8) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 9) = "Reactive Spirit Asterite Armor Tincture"
    .AlchDetails(AD_COST, 9) = "707860"
    .AlchDetails(AD_EFFECT, 9) = "77 Spirit"
    .AlchDetails(AD_LEVEL, 9) = "40"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 10) = "Reactive Fire Netherium Armor Tincture"
    .AlchDetails(AD_COST, 10) = "1390540"
    .AlchDetails(AD_EFFECT, 10) = "86 Heat"
    .AlchDetails(AD_LEVEL, 10) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 11) = "Reactive Cold Netherium Armor Tincture"
    .AlchDetails(AD_COST, 11) = "1390540"
    .AlchDetails(AD_EFFECT, 11) = "86 Cold"
    .AlchDetails(AD_LEVEL, 11) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 12) = "Reactive Energy Netherium Armor Tincture"
    .AlchDetails(AD_COST, 12) = "1390540"
    .AlchDetails(AD_EFFECT, 12) = "86 Energy"
    .AlchDetails(AD_LEVEL, 12) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 13) = "Reactive Spirit Netherium Armor Tincture"
    .AlchDetails(AD_COST, 13) = "1390540"
    .AlchDetails(AD_EFFECT, 13) = "86 Spirit"
    .AlchDetails(AD_LEVEL, 13) = "43"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 14) = "Reactive Draining Netherium Armor Tincture"
    .AlchDetails(AD_COST, 14) = "2034020"
    .AlchDetails(AD_EFFECT, 14) = "Omni Drain:75:H:100%:P:60%:E:40%"
    .AlchDetails(AD_LEVEL, 14) = "45"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 15) = "Reactive Coil Netherium Armor Tincture"
    .AlchDetails(AD_COST, 15) = "2034020"
    .AlchDetails(AD_EFFECT, 15) = "Speed Decrease: 30%, 15 Sec"
    .AlchDetails(AD_LEVEL, 15) = "45"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 16) = "Reactive Depletion Netherium Armor Tincture"
    .AlchDetails(AD_COST, 16) = "2018040"
    .AlchDetails(AD_EFFECT, 16) = "Power Drain: 35 : 50%"
    .AlchDetails(AD_LEVEL, 16) = "45"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 17) = "Reactive Fire Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 17) = "2244560"
    .AlchDetails(AD_EFFECT, 17) = "95 Heat"
    .AlchDetails(AD_LEVEL, 17) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 18) = "Reactive Cold Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 18) = "2244560"
    .AlchDetails(AD_EFFECT, 18) = "95 Cold"
    .AlchDetails(AD_LEVEL, 18) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 19) = "Reactive Energy Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 19) = "2244560"
    .AlchDetails(AD_EFFECT, 19) = "95 Energy"
    .AlchDetails(AD_LEVEL, 19) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 20) = "Reactive Spirit Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 20) = "2244560"
    .AlchDetails(AD_EFFECT, 20) = "95 Spirit"
    .AlchDetails(AD_LEVEL, 20) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 21) = "Reactive Hardening Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 21) = "2541060"
    .AlchDetails(AD_EFFECT, 21) = "Self AF Buff: 75"
    .AlchDetails(AD_LEVEL, 21) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 22) = "Reactive Eroding Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 22) = "2527846"
    .AlchDetails(AD_EFFECT, 22) = "DoT: 64/Tick"
    .AlchDetails(AD_LEVEL, 22) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 23) = "Reactive Cleric Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 23) = "2544560"
    .AlchDetails(AD_EFFECT, 23) = "Self Haste: 20%"
    .AlchDetails(AD_LEVEL, 23) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 24) = "Reactive Shard Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 24) = "2544560"
    .AlchDetails(AD_EFFECT, 24) = "Self  Shield: 5%"
    .AlchDetails(AD_LEVEL, 24) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 25) = "Reactive Ablative Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 25) = "2541060"
    .AlchDetails(AD_EFFECT, 25) = "Self Melee Health Buffer: 100"
    .AlchDetails(AD_LEVEL, 25) = "47"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 26) = "Reactive Coil Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 26) = "2814060"
    .AlchDetails(AD_EFFECT, 26) = "Speed Decrease: 35%, 15 Sec"
    .AlchDetails(AD_LEVEL, 26) = "49"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 27) = "Reactive Depletion Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 27) = "2857060"
    .AlchDetails(AD_EFFECT, 27) = "Power Drain: 55% : 50%"
    .AlchDetails(AD_LEVEL, 27) = "49"
    
    ProgressBar.Width = ProgressBar.Width + SP_VS_PROGRESS_CHANGE
    .AlchDetails(AD_NAME, 28) = "Reactive Draining Arcanium Armor Tincture"
    .AlchDetails(AD_COST, 28) = "2827090"
    .AlchDetails(AD_EFFECT, 28) = "Omni Drain:100:H:100%:P:60%:E:40%"
    .AlchDetails(AD_LEVEL, 28) = "49"
End With 'alchemyeffects(ae_reactives)

'Index 3 is Dropped
With AlchemyEffects(AE_DROPS)

    StatusBar.Caption = "Initializing Drops"
    DoEvents
    
    .AlchType = "Drops"
    
End With    'alchemyeffects(ae_drops)

End Sub


