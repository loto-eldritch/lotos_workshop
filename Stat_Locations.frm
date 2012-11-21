VERSION 5.00
Begin VB.Form Stat_Locations 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2685
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frame_StatLoc 
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
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2445
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   3
         Top             =   2070
         Width           =   2265
      End
      Begin VB.ListBox StatList 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   1680
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   2265
      End
   End
   Begin VB.CommandButton form_Border 
      Height          =   2715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2685
   End
End
Attribute VB_Name = "Stat_Locations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Populate_HitsView()

    Dim Ctr As Long
    Dim lTotal As Long
    
    StatList.Clear
    
    Stat_Locations.Caption = "Stat Detail"
    frame_StatLoc.Caption = "Hits"
        
    lTotal = 0
    For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL
        If TOON.STAT_MATRIX(SM_HIT, Ctr) > 0 Then
            StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_HIT, Ctr)
            lTotal = lTotal + TOON.STAT_MATRIX(SM_HIT, Ctr)
        End If
    Next Ctr
    
    frame_StatLoc.Caption = frame_StatLoc.Caption & " :: " & lTotal
    Stat_Locations.Show vbModal, WS

End Sub

Public Sub Populate_HitsCapView()

    Dim Ctr As Long
    Dim lTotal As Long
    
    StatList.Clear
    
    Stat_Locations.Caption = "Cap Detail"
    frame_StatLoc.Caption = "Hits"
        
    lTotal = 0
    For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL
        If TOON.STAT_MATRIX(SM_HIT_CAP, Ctr) > 0 Then
            StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_HIT_CAP, Ctr)
            lTotal = lTotal + TOON.STAT_MATRIX(SM_HIT_CAP, Ctr)
        End If
    Next Ctr
    
    frame_StatLoc.Caption = frame_StatLoc.Caption & " :: " & lTotal
    Stat_Locations.Show vbModal, WS

End Sub

Public Sub Populate_PowerView()

    Dim Ctr As Long
    Dim lTotal As Long
    
    StatList.Clear
    
    Stat_Locations.Caption = "Stat Detail"
    
    frame_StatLoc.Caption = "Power"
        
    lTotal = 0
    For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL
        If TOON.STAT_MATRIX(SM_POW, Ctr) > 0 Then
            StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_POW, Ctr)
            lTotal = lTotal + TOON.STAT_MATRIX(SM_POW, Ctr)
        End If
    Next Ctr
    
    frame_StatLoc.Caption = frame_StatLoc.Caption & " :: " & lTotal
    Stat_Locations.Show vbModal, WS

End Sub

Public Sub Populate_PowerCapView()

    Dim Ctr As Long
    Dim lTotal As Long
    
    StatList.Clear
    
    Stat_Locations.Caption = "Cap Detail"
    frame_StatLoc.Caption = "Power"
    
    lTotal = 0
    For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL
        If TOON.STAT_MATRIX(SM_POW_CAP, Ctr) > 0 Then
            StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_POW_CAP, Ctr)
            lTotal = lTotal + TOON.STAT_MATRIX(SM_POW_CAP, Ctr)
        End If
    Next Ctr
    
    frame_StatLoc.Caption = frame_StatLoc.Caption & " :: " & lTotal
    Stat_Locations.Show vbModal, WS

End Sub

Public Sub Populate_ResistView(ResistIndex As Integer)

    Dim Ctr As Long
    Dim lCode As Long
    Dim lTotal As Long
    
    StatList.Clear
    
    Stat_Locations.Caption = "Resist Detail"
    frame_StatLoc.Caption = WS.lbl_Attribute_Name(ResistIndex).Caption
    
    Select Case ResistIndex
        Case WS_ATTR_CRUSH  'crush
            lCode = SM_CRUSH_RESIST
        Case WS_ATTR_SLASH  'slash
            lCode = SM_SLASH_RESIST
        Case WS_ATTR_THRUST 'thrust
            lCode = SM_THRUST_RESIST
        Case WS_ATTR_HEAT   'heat
            lCode = SM_HEAT_RESIST
        Case WS_ATTR_COLD   'cold
            lCode = SM_COLD_RESIST
        Case WS_ATTR_MATTER 'matter
            lCode = SM_MATTER_RESIST
        Case WS_ATTR_BODY   'body
            lCode = SM_BODY_RESIST
        Case WS_ATTR_SPIRIT 'spirit
            lCode = SM_SPIRIT_RESIST
        Case WS_ATTR_ENERGY 'energy
            lCode = SM_ENERGY_RESIST
    End Select
    
    lTotal = 0
    For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
        If TOON.STAT_MATRIX(lCode, Ctr) > 0 Then
            StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(lCode, Ctr)
            lTotal = lTotal + TOON.STAT_MATRIX(lCode, Ctr)
        End If
    Next Ctr
    
    frame_StatLoc.Caption = frame_StatLoc.Caption & " :: " & lTotal
    
    Stat_Locations.Show vbModal, WS
    
End Sub

Public Sub Populate_StatView(StatIndex As Integer)

    Dim Ctr As Long
    Dim lTotal As Long
    
    StatList.Clear
    
    Stat_Locations.Caption = "Stat Detail"
    
    Select Case StatIndex
        Case WS_ATTR_STR
            frame_StatLoc.Caption = "Strength"
        Case WS_ATTR_CON
            frame_StatLoc.Caption = "Constitution"
        Case WS_ATTR_DEX
            frame_StatLoc.Caption = "Dexterity"
        Case WS_ATTR_QUI
            frame_StatLoc.Caption = "Quickness"
        Case WS_ATTR_INT
            frame_StatLoc.Caption = "Intelligence"
        Case WS_ATTR_EMP
            frame_StatLoc.Caption = "Empathy"
        Case WS_ATTR_PIE
            frame_StatLoc.Caption = "Piety"
        Case WS_ATTR_CHA
            frame_StatLoc.Caption = "Charisma"
    End Select
    
    lTotal = 0
    
    For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL
        If TOON.STAT_MATRIX(StatIndex + 1, Ctr) > 0 Then
            StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(StatIndex + 1, Ctr)
            lTotal = lTotal + TOON.STAT_MATRIX(StatIndex + 1, Ctr)
        End If
    Next Ctr
    
    If StatIndex = WS_ATTR_INT Then
        
        If (TOON.pStat = SM_INT Or TOON.sStat = SM_INT Or TOON.tStat = SM_INT) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU, Ctr)
                End If
            Next Ctr
        End If
    ElseIf StatIndex = WS_ATTR_EMP Then
        
        If (TOON.pStat = SM_EMP Or TOON.sStat = SM_EMP Or TOON.tStat = SM_EMP) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU, Ctr)
                End If
            Next Ctr
        End If
    ElseIf StatIndex = WS_ATTR_PIE Then
    
        If (TOON.pStat = SM_PIE Or TOON.sStat = SM_PIE Or TOON.tStat = SM_PIE) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU, Ctr)
                End If
            Next Ctr
        End If
    ElseIf StatIndex = WS_ATTR_CHA Then
        
        If (TOON.pStat = SM_CHA Or TOON.sStat = SM_CHA Or TOON.tStat = SM_CHA) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU, Ctr)
                End If
            Next Ctr
        End If
    End If
    
    frame_StatLoc.Caption = frame_StatLoc.Caption & " :: " & lTotal
    
    Stat_Locations.Show vbModal, WS
    
End Sub

Public Sub Populate_CapView(CapIndex As Integer)
    
    Dim Ctr As Long
    Dim CapLoc As Long
    Dim lTotal As Long
    
    StatList.Clear
    
    Stat_Locations.Caption = "Cap Detail"
    
    Select Case CapIndex
        Case WS_ATTR_STR
            frame_StatLoc.Caption = "Strength"
            CapLoc = SM_STR_CAP
        Case WS_ATTR_CON
            frame_StatLoc.Caption = "Constitution"
            CapLoc = SM_CON_CAP
        Case WS_ATTR_DEX
            frame_StatLoc.Caption = "Dexterity"
            CapLoc = SM_DEX_CAP
        Case WS_ATTR_QUI
            frame_StatLoc.Caption = "Quickness"
            CapLoc = SM_QUI_CAP
        Case WS_ATTR_INT
            frame_StatLoc.Caption = "Intelligence"
            CapLoc = SM_INT_CAP
        Case WS_ATTR_EMP
            frame_StatLoc.Caption = "Empathy"
            CapLoc = SM_EMP_CAP
        Case WS_ATTR_PIE
            frame_StatLoc.Caption = "Piety"
            CapLoc = SM_PIE_CAP
        Case WS_ATTR_CHA
            frame_StatLoc.Caption = "Charisma"
            CapLoc = SM_CHA_CAP
    End Select
    
    For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL
        If TOON.STAT_MATRIX(CapLoc, Ctr) > 0 Then
            StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(CapLoc, Ctr)
            lTotal = lTotal + TOON.STAT_MATRIX(CapLoc, Ctr)
        End If
    Next Ctr
    
    If CapIndex = WS_ATTR_INT Then
        
        If (TOON.pStat = SM_INT Or TOON.sStat = SM_INT Or TOON.tStat = SM_INT) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU_CAP, Ctr)
                End If
            Next Ctr
        End If
    ElseIf CapIndex = WS_ATTR_EMP Then
        
        If (TOON.pStat = SM_EMP Or TOON.sStat = SM_EMP Or TOON.tStat = SM_EMP) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU_CAP, Ctr)
                End If
            Next Ctr
        End If
    ElseIf CapIndex = WS_ATTR_PIE Then
    
        If (TOON.pStat = SM_PIE Or TOON.sStat = SM_PIE Or TOON.tStat = SM_PIE) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU_CAP, Ctr)
                End If
            Next Ctr
        End If
    ElseIf CapIndex = WS_ATTR_CHA Then
        
        If (TOON.pStat = SM_CHA Or TOON.sStat = SM_CHA Or TOON.tStat = SM_CHA) Then
            For Ctr = SM_LOC_HEAD To SM_LOC_MYTHICAL 'SM_LOC_RANGED
                If TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) > 0 Then
                    StatList.AddItem StatMatrixRowName(Ctr) & " " & TOON.STAT_MATRIX(SM_ACU_CAP, Ctr) & " " & "Acuity"
                    lTotal = lTotal + TOON.STAT_MATRIX(SM_ACU_CAP, Ctr)
                End If
            Next Ctr
        End If
    End If
    
    frame_StatLoc.Caption = frame_StatLoc.Caption & " :: " & lTotal
    Stat_Locations.Show vbModal, WS
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub StatList_Click()

    Dim lCode As Integer
    
    If InStr(LCase$(StatList.list(StatList.ListIndex)), "chest") <> 0 Then
        lCode = WS_DOLL_CHEST
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "arms") <> 0 Then
        lCode = WS_DOLL_ARMS
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "jewel") <> 0 Then
        lCode = WS_DOLL_GEM
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "left ring") <> 0 Then
        lCode = WS_DOLL_LRING
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "left bracer") <> 0 Then
        lCode = WS_DOLL_LWRIST
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "legs") <> 0 Then
        lCode = WS_DOLL_LEGS
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "main") <> 0 Then
        If WS.chk_Equip_SC(WS_DOLL_RHAND).Value = 1 Or WS.chk_Equip_DP(WS_DOLL_RHAND).Value = 1 Then
            lCode = WS_DOLL_RHAND
        Else
            lCode = WS_DOLL_RIGHTSPARE
        End If
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "off") <> 0 Then
        If WS.chk_Equip_SC(WS_DOLL_LHAND).Value = 1 Or WS.chk_Equip_DP(WS_DOLL_LHAND).Value = 1 Then
            lCode = WS_DOLL_LHAND
        Else
            lCode = WS_DOLL_LEFTSPARE
        End If
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "two handed") <> 0 Then
        If WS.chk_Equip_SC(WS_DOLL_2HAND).Value = 1 Or WS.chk_Equip_DP(WS_DOLL_2HAND).Value = 1 Then
            lCode = WS_DOLL_2HAND
        Else
            lCode = WS_DOLL_2HANDSPARE
        End If
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "ranged") <> 0 Then
        If WS.chk_Equip_SC(WS_DOLL_RANGED).Value = 1 Or WS.chk_Equip_DP(WS_DOLL_RANGED).Value = 1 Then
            lCode = WS_DOLL_RANGED
        Else
            lCode = WS_DOLL_RANGEDSPARE
        End If
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "feet") <> 0 Then
        lCode = WS_DOLL_FEET
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "right bracer") <> 0 Then
        lCode = WS_DOLL_RWRIST
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "right ring") <> 0 Then
        lCode = WS_DOLL_RRING
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "belt") <> 0 Then
        lCode = WS_DOLL_WAIST
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "hands") <> 0 Then
        lCode = WS_DOLL_HANDS
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "head") <> 0 Then
        lCode = WS_DOLL_HEAD
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "cloak") <> 0 Then
        lCode = WS_DOLL_CLOAK
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "necklace") <> 0 Then
        lCode = WS_DOLL_NECK
    ElseIf InStr(LCase$(StatList.list(StatList.ListIndex)), "mythical") <> 0 Then
        lCode = WS_DOLL_MYTHICAL
    Else
        lCode = 0
    End If
    
    Call WS.ClickPiece(lCode)
    
End Sub
