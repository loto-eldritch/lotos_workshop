VERSION 5.00
Begin VB.Form Suit_Configuration 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Template Configuration"
   ClientHeight    =   5115
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   8190
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
   ScaleHeight     =   5115
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frame_TemplateConfig 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8190
      Begin VB.TextBox txt_Template 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   4890
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   150
         Width           =   8040
      End
   End
   Begin VB.Menu mnuTemplate_SaveText 
      Caption         =   "&Save"
   End
   Begin VB.Menu mnuTemplate_Print 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuTemplate_Close 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "Suit_Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sBREAK As String = "________________________________________________________________" & vbCrLf & vbCrLf

Private Sub Form_Load()

    StayOnTop Me
     
    txt_Template.Text = ReportT(TOON)
     
End Sub

Private Function GetSuitLocation(Index As Integer, LocationName As String, IsAccessory As Long) As String
    
    Dim sText As String
    
    sText = sText & LocationName & ":" & vbCrLf & vbCrLf
    
    If IsAccessory = 0 Then
        If WS.chk_Equip_SC(Index).Value = vbChecked Then
            sText = sText & GetSCInfoT(Index)
        ElseIf WS.chk_Equip_DP(Index).Value = vbChecked Then
            sText = sText & GetDPInfoT(Index)
        End If
    Else
        If WS.chk_Equip_DP(Index).Value = vbChecked Then
            sText = sText & GetDPInfoT(Index)
        End If
    End If
    
    GetSuitLocation = sText

End Function

Private Function GetSCInfoT(Index As Integer) As String

    Dim sText As String
    Dim sTemp As String
    
    sText = sText & vbTab & "Item Quality: " & WS.cmb_ItemQuality(Index).Text & vbTab & "Imbue Points: " & WS.lbl_ImbuePTS_Total_SC(Index).Caption & WS.lbl_ImbuePTS_Avail_SC(Index).Caption & vbCrLf
    sText = sText & vbTab & "Overcharge: " & WS.lbl_Overcharge_SC(Index).Caption & vbCrLf
    sText = sText & vbTab & "Utility: " & WS.lbl_ItemUtility_Value(Index).Caption & vbCrLf & vbCrLf
    
    
    If Len(WS.cmb_GemEffectSC1(Index).Text) <> 0 Then
        sTemp = WS.cmb_GemAmountSC1(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemEffectSC1(Index).Text) > 5 Then
            sText = sText & vbTab & "Gem 1: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC1(Index).Text & vbTab & _
                            ": " & WS.lbl_GemNameSC1(Index).Caption & vbCrLf
        Else
            sText = sText & vbTab & "Gem 1: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC1(Index).Text & vbTab & vbTab & _
                            ": " & WS.lbl_GemNameSC1(Index).Caption & vbCrLf
        End If
        
    End If
    
    If Len(WS.cmb_GemEffectSC2(Index).Text) <> 0 Then
        sTemp = WS.cmb_GemAmountSC2(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemEffectSC2(Index).Text) > 5 Then
            sText = sText & vbTab & "Gem 2: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC2(Index).Text & vbTab & _
                            ": " & WS.lbl_GemNameSC2(Index).Caption & vbCrLf
        Else
            sText = sText & vbTab & "Gem 2: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC2(Index).Text & vbTab & vbTab & _
                            ": " & WS.lbl_GemNameSC2(Index).Caption & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffectSC3(Index).Text) <> 0 Then
        sTemp = WS.cmb_GemAmountSC3(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemEffectSC3(Index).Text) > 5 Then
            sText = sText & vbTab & "Gem 3: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC3(Index).Text & vbTab & _
                            ": " & WS.lbl_GemNameSC3(Index).Caption & vbCrLf
        Else
            sText = sText & vbTab & "Gem 3: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC3(Index).Text & vbTab & vbTab & _
                            ": " & WS.lbl_GemNameSC3(Index).Caption & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffectSC4(Index).Text) <> 0 Then
        sTemp = WS.cmb_GemAmountSC4(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemEffectSC4(Index).Text) > 5 Then
            sText = sText & vbTab & "Gem 4: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC4(Index).Text & vbTab & _
                            ": " & WS.lbl_GemNameSC4(Index).Caption & vbCrLf
        Else
            sText = sText & vbTab & "Gem 4: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC4(Index).Text & vbTab & vbTab & _
                            ": " & WS.lbl_GemNameSC4(Index).Caption & vbCrLf
        End If
    End If

    If Len(WS.cmb_GemEffectSC5(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmountSC5(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        sText = sText & vbTab & "Slot 5: " & vbTab & sTemp & " " & WS.cmb_GemEffectSC5(Index).Text & " " & WS.cmb_GemSelectSC5(Index).Text & vbCrLf
        
    End If
    
    If Len(WS.lbl_TinctureNameSC(Index).Caption) <> 0 Then
        sText = sText & vbTab & "Tincture: " & WS.lbl_TinctureNameSC(Index).Caption & vbCrLf
    End If
    
    sText = sText & sBREAK

    GetSCInfoT = sText
    
End Function

Private Function GetDPInfoT(Index As Integer) As String

    Dim sText As String
    Dim sTemp As String
    
    sText = sText & vbTab & WS.txt_ItemName_DP(Index).Text & vbCrLf
    sText = sText & vbTab & WS.lbl_ItemUtilityDP_Value(Index).Caption & vbCrLf & vbCrLf
    
    If Len(WS.cmb_GemEffect_DP1(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP1(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP1(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 1: " & WS.cmb_GemSelect_DP1(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP1(Index).Text & " " & WS.cmb_GemEffect_DP1(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 1: " & WS.cmb_GemSelect_DP1(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP1(Index).Text & " " & WS.cmb_GemEffect_DP1(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP2(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP2(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP2(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 2: " & WS.cmb_GemSelect_DP2(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP2(Index).Text & " " & WS.cmb_GemEffect_DP2(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 2: " & WS.cmb_GemSelect_DP2(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP2(Index).Text & " " & WS.cmb_GemEffect_DP2(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP3(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP3(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP3(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 3: " & WS.cmb_GemSelect_DP3(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP3(Index).Text & " " & WS.cmb_GemEffect_DP3(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 3: " & WS.cmb_GemSelect_DP3(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP3(Index).Text & " " & WS.cmb_GemEffect_DP3(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP4(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP4(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP4(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 4: " & WS.cmb_GemSelect_DP4(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP4(Index).Text & " " & WS.cmb_GemEffect_DP4(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 4: " & WS.cmb_GemSelect_DP4(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP4(Index).Text & " " & WS.cmb_GemEffect_DP4(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP5(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP5(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP5(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 5: " & WS.cmb_GemSelect_DP5(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP5(Index).Text & " " & WS.cmb_GemEffect_DP5(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 5: " & WS.cmb_GemSelect_DP5(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP5(Index).Text & " " & WS.cmb_GemEffect_DP5(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP6(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP6(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP6(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 6: " & WS.cmb_GemSelect_DP6(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP6(Index).Text & " " & WS.cmb_GemEffect_DP6(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 6: " & WS.cmb_GemSelect_DP6(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP6(Index).Text & " " & WS.cmb_GemEffect_DP6(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP7(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP7(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP7(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 7: " & WS.cmb_GemSelect_DP7(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP7(Index).Text & " " & WS.cmb_GemEffect_DP7(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 7: " & WS.cmb_GemSelect_DP7(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP7(Index).Text & " " & WS.cmb_GemEffect_DP7(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP8(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP8(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP8(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 8: " & WS.cmb_GemSelect_DP8(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP8(Index).Text & " " & WS.cmb_GemEffect_DP8(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 8: " & WS.cmb_GemSelect_DP8(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP8(Index).Text & " " & WS.cmb_GemEffect_DP8(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP9(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP9(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP9(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 9: " & WS.cmb_GemSelect_DP9(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP9(Index).Text & " " & WS.cmb_GemEffect_DP9(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 9: " & WS.cmb_GemSelect_DP9(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP9(Index).Text & " " & WS.cmb_GemEffect_DP9(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP10(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP10(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP10(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 10: " & WS.cmb_GemSelect_DP10(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP10(Index).Text & " " & WS.cmb_GemEffect_DP10(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 10: " & WS.cmb_GemSelect_DP10(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP10(Index).Text & " " & WS.cmb_GemEffect_DP10(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP11(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP11(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP11(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 11: " & WS.cmb_GemSelect_DP11(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP11(Index).Text & " " & WS.cmb_GemEffect_DP11(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 11: " & WS.cmb_GemSelect_DP11(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP11(Index).Text & " " & WS.cmb_GemEffect_DP11(Index).Text & vbCrLf
        End If
    End If
    
    If Len(WS.cmb_GemEffect_DP12(Index).Text) <> 0 Then
        sTemp = WS.txt_GemAmount_DP12(Index).Text
        If Len(sTemp) = 1 Then sTemp = sTemp & " "
        
        If Len(WS.cmb_GemSelect_DP12(Index).Text) < 11 Then
            sText = sText & vbTab & "Slot 12: " & WS.cmb_GemSelect_DP12(Index).Text & vbTab & vbTab & " :: " & _
                                WS.txt_GemAmount_DP12(Index).Text & " " & WS.cmb_GemEffect_DP12(Index).Text & vbCrLf
        Else
            sText = sText & vbTab & "Slot 12: " & WS.cmb_GemSelect_DP12(Index).Text & vbTab & " :: " & _
                                WS.txt_GemAmount_DP12(Index).Text & " " & WS.cmb_GemEffect_DP12(Index).Text & vbCrLf
        End If
    End If
        
    sText = sText & sBREAK
    
    GetDPInfoT = sText
    
End Function

Private Function Report_ArmorT() As String

    Dim sText As String
    
    sText = sText & "Piece Listing: Armor" & vbCrLf
    sText = sText & sBREAK
    
    sText = sText & GetSuitLocation(WS_DOLL_HEAD, "Head", 0)
    sText = sText & GetSuitLocation(WS_DOLL_CHEST, "Chest", 0)
    sText = sText & GetSuitLocation(WS_DOLL_ARMS, "Arms", 0)
    sText = sText & GetSuitLocation(WS_DOLL_LEGS, "Legs", 0)
    sText = sText & GetSuitLocation(WS_DOLL_HANDS, "Hands", 0)
    sText = sText & GetSuitLocation(WS_DOLL_FEET, "Feet", 0)
        
    Report_ArmorT = sText

End Function

Private Function Report_JewelryT() As String

    Dim sText As String
    
    sText = sText & "Piece Listing: Accessories" & vbCrLf
    sText = sText & sBREAK
    
    sText = sText & GetSuitLocation(WS_DOLL_NECK, "Necklace", 1)
    sText = sText & GetSuitLocation(WS_DOLL_CLOAK, "Cloak", 1)
    sText = sText & GetSuitLocation(WS_DOLL_GEM, "Jewel", 1)
    sText = sText & GetSuitLocation(WS_DOLL_WAIST, "Belt", 1)
    sText = sText & GetSuitLocation(WS_DOLL_LRING, "Left Ring", 1)
    sText = sText & GetSuitLocation(WS_DOLL_RRING, "Right Ring", 1)
    sText = sText & GetSuitLocation(WS_DOLL_LWRIST, "Left Bracer", 1)
    sText = sText & GetSuitLocation(WS_DOLL_RWRIST, "Right Bracer", 1)
    sText = sText & GetSuitLocation(WS_DOLL_MYTHICAL, "Mythical", 1)
    
    Report_JewelryT = sText
    
End Function

Private Function Report_WeaponsT() As String

    Dim sText As String

    sText = sText & "Piece Listing: Weapons" & vbCrLf
    sText = sText & sBREAK
    
    sText = sText & GetSuitLocation(WS_DOLL_RHAND, "Main Hand", 0)
    sText = sText & GetSuitLocation(WS_DOLL_LHAND, "Off-Hand", 0)
    sText = sText & GetSuitLocation(WS_DOLL_2HAND, "2-Handed", 0)
    sText = sText & GetSuitLocation(WS_DOLL_RANGED, "Ranged", 0)
    sText = sText & GetSuitLocation(WS_DOLL_RIGHTSPARE, "Spare Main Hand", 0)
    sText = sText & GetSuitLocation(WS_DOLL_LEFTSPARE, "Spare Off-Hand", 0)
    sText = sText & GetSuitLocation(WS_DOLL_2HANDSPARE, "Spare 2-Handed", 0)
    sText = sText & GetSuitLocation(WS_DOLL_RANGEDSPARE, "Spare Ranged", 0)
    
    Report_WeaponsT = sText
    
End Function

Private Function ReportT(tToon As TOON_TYPE) As String

    Dim Ctr As Long
    Dim sText As String
    Dim ColumnSum As Long
    
    Dim SkillTotal As Long
    
    txt_Template.Text = vbNullString
    
    sText = "Template Report for: " & WS.txt_CharacterName.Text & " the " & WS.cmb_Race.Text & " " & WS.cmb_Class.Text & " of " & WS.cmb_Realm.Text & vbCrLf
    'usable stats
    sText = sText & sBREAK
    sText = sText & "Stats:" & vbCrLf
    sText = sText & sBREAK
        
    sText = sText & "Str:" & vbTab & tToon.STR & vbTab & "Int:" & vbTab & tToon.INT & vbTab & _
                    "Hits:" & vbTab & tToon.HITS & vbCrLf
                    
    sText = sText & "Con:" & vbTab & tToon.CON & vbTab & "Pie:" & vbTab & tToon.PIE & vbTab & _
                    "Power:" & vbTab & tToon.POWER & vbCrLf
                    
    sText = sText & "Dex:" & vbTab & tToon.DEX & vbTab & "Cha:" & vbTab & tToon.CHA & vbCrLf
    sText = sText & "Qui:" & vbTab & tToon.QUI & vbTab & "Emp:" & vbTab & tToon.EMP & vbCrLf
    
    sText = sText & sBREAK
    sText = sText & "Cap Increases:" & vbCrLf
    sText = sText & sBREAK

    sText = sText & "Str :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_STR).Tag & vbTab & _
                    "Int :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_INT).Tag & vbTab & _
                    "Hits :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_HIT).Tag & vbCrLf
                    
    sText = sText & "Con:" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_CON).Tag & vbTab & _
                    "Pie :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_PIE).Tag & vbTab & _
                    "Power :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_POW).Tag & vbCrLf
                    
    sText = sText & "Dex :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_DEX).Tag & vbTab & _
                    "Cha :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_CHA).Tag & vbCrLf
                    
    sText = sText & "Qui :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_QUI).Tag & vbTab & _
                    "Emp :" & vbTab & WS.lbl_Attribute_Cap_Value(WS_ATTR_EMP).Tag & vbCrLf
    
    
    sText = sText & sBREAK
    sText = sText & "Resists:" & vbCrLf
    sText = sText & sBREAK
    
    sText = sText & "Crush:" & vbTab & tToon.CRUSH & vbTab & "Heat:" & vbTab & tToon.HEAT & vbTab & "Body:" & vbTab & tToon.BODY & vbCrLf
    sText = sText & "Slash:" & vbTab & tToon.SLASH & vbTab & "Cold:" & vbTab & tToon.COLD & vbTab & "Spirit:" & vbTab & tToon.SPIRIT & vbCrLf
    sText = sText & "Thrust:" & vbTab & tToon.THRUST & vbTab & "Matter:" & vbTab & tToon.MATTER & vbTab & "Energy:" & vbTab & tToon.ENERGY & vbCrLf
    
    sText = sText & sBREAK
    sText = sText & "Skills:" & vbCrLf
    sText = sText & sBREAK
    
    For Ctr = 0 To (WS.list_Skills.ListCount - 1)
        If SC_SETTINGS.BONUS_OPTION = BO_TOCAP Then
            SkillTotal = Val(WS.list_Skills.list(Ctr))
            SkillTotal = Int(((tToon.LEVEL * 0.2) + 1) - SkillTotal)
            sText = sText & SkillTotal & " " & Trim$(Mid$(WS.list_Skills.list(Ctr), InStr(WS.list_Skills.list(Ctr), " "))) & vbCrLf
        Else
            sText = sText & WS.list_Skills.list(Ctr) & vbCrLf
        End If
    Next Ctr
   
    sText = sText & sBREAK
    sText = sText & "Focus:" & vbCrLf
    sText = sText & sBREAK
    
    For Ctr = 0 To WS.list_OtherBonus.ListCount
        If SC_SETTINGS.BONUS_OPTION = BO_TOCAP Then
        
        Else
            If InStr(LCase$(WS.list_OtherBonus.list(Ctr)), "focus") <> 0 Then sText = sText & WS.list_OtherBonus.list(Ctr) & vbCrLf
        End If
    Next Ctr
    
    sText = sText & sBREAK
    sText = sText & "Other Bonuses:" & vbCrLf
    sText = sText & sBREAK
    

    If SC_SETTINGS.BONUS_OPTION = BO_TOCAP Then
    
        For Ctr = SM_AFBONUS To SM_PVEUNIQUE
            SkillTotal = SumColumn(Ctr, SC_SETTINGS.BONUS_OPTION)

            If SkillTotal > 0 Then
                                
                sText = sText & SkillTotal & " " & StatMatrixColumnName(Ctr) & vbCrLf
                
            End If
        Next Ctr
    
        For Ctr = SM_LOTM_ENC To SM_LOTM_WATERBREATHING
        
            SkillTotal = SumColumn(Ctr, SC_SETTINGS.BONUS_OPTION)
                                   
            If SkillTotal > 0 Then sText = sText & SkillTotal & " " & StatMatrixColumnName(Ctr) & vbCrLf
        
        Next Ctr
    Else
        For Ctr = 0 To WS.list_OtherBonus.ListCount
            If InStr(LCase$(WS.list_OtherBonus.list(Ctr)), "focus") = 0 Then sText = sText & WS.list_OtherBonus.list(Ctr) & vbCrLf
        Next Ctr
    End If
    
    sText = sText & sBREAK
    
    sText = sText & Report_ArmorT
    sText = sText & Report_WeaponsT
    sText = sText & Report_JewelryT
    
    ReportT = sText

End Function

Private Sub mnuTemplate_Close_Click()

    txt_Template.Text = vbNullString
    Unload Me
    
End Sub

Private Sub mnuTemplate_Print_Click()

    On Error Resume Next
    
    Printer.Print txt_Template.Text
    Printer.EndDoc

End Sub

Private Sub mnuTemplate_SaveText_Click()

    Dim hFile As Long
    Dim sBuffer As String
    Dim sPath As String
    Dim lRet As Long
    
    Dim lWrite As Long
    Dim cmdFlags As Long
    Dim cmdFilter As String
    Dim cmdMessage As String
    
    cmdFlags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    cmdFilter = "Text Files (*.txt)" & vbNullChar & "*.txt" & vbNullChar
    cmdMessage = "Save Configuration as Text"
      
    sPath = CMD_OpenSave(lSave, Me.hwnd, cmdFilter, 1, App.Path, cmdMessage, cmdFlags)
    
    If sPath <> vbNullString Then
        If LCase$(Mid(sPath, Len(sPath) - 3)) <> ".txt" Then sPath = sPath & ".txt"
        
        hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
        
        If hFile = -1 Then Exit Sub '*If hFile is -1 the file is not there and there has been an error
        
        sBuffer = txt_Template.Text
        
        lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
        lRet = CloseHandle(hFile)
    End If
    
End Sub
