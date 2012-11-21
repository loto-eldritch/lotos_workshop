VERSION 5.00
Begin VB.Form Material_List 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Material Listing"
   ClientHeight    =   4080
   ClientLeft      =   150
   ClientTop       =   375
   ClientWidth     =   6615
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
   ScaleHeight     =   4080
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox list_GemNames 
      Height          =   2940
      ItemData        =   "Material_List.frx":0000
      Left            =   7020
      List            =   "Material_List.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   2400
   End
   Begin VB.Frame frame_Material 
      Height          =   4155
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   6615
      Begin VB.HScrollBar HScroll_Mult 
         Height          =   225
         Left            =   60
         Max             =   100
         TabIndex        =   2
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox txt_Material 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3630
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   450
         Width           =   6480
      End
      Begin VB.Label lbl_MultValue 
         Alignment       =   2  'Center
         Caption         =   "1"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   3330
         TabIndex        =   5
         Top             =   180
         Width           =   345
      End
      Begin VB.Label lbl_Mult 
         Caption         =   "Material Multiplier"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1755
         TabIndex        =   4
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Menu mnuSaveText 
      Caption         =   "&Save"
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "Material_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type GemParams
    mRealm As Long
    mGem As String
    mIndex As Long
    mType As String
End Type

Private pOneGem As GemParams

Private aDust(11) As Long
Private aLiquid(10) As Long
Private aContainer(9) As Long
Private aTotal As Long

Private sGemReport_HTML As String

Private WhoCalled As Long   '0 = one gem, 1 = all gems

Private Const sBREAK As String = "____________________________________________________________________" & vbCrLf

Private Sub Clear_DustLiquidContainer()

    Dim Ctr As Long
    
    For Ctr = 0 To 11
        aDust(Ctr) = 0
    Next Ctr
    
    For Ctr = 0 To 10
        aLiquid(Ctr) = 0
    Next Ctr
    
    For Ctr = 0 To 9
        aContainer(Ctr) = 0
    Next Ctr
    
    aTotal = 0
    
End Sub

Private Sub HScroll_Mult_Change()

    lbl_MultValue.Caption = HScroll_Mult.Value
    txt_Material.Text = vbNullString
    
    If WhoCalled = 0 Then
        txt_Material.Text = OneGemMats(pOneGem.mRealm, pOneGem.mGem, pOneGem.mIndex, pOneGem.mType)
    ElseIf WhoCalled = 1 Then
        txt_Material.Text = AllGemMats_TEXT(TOON.REALM)
    End If
    
End Sub

Private Sub HScroll_Mult_Scroll()

    lbl_MultValue.Caption = HScroll_Mult.Value
    txt_Material.Text = vbNullString
    
    If WhoCalled = 0 Then
        txt_Material.Text = OneGemMats(pOneGem.mRealm, pOneGem.mGem, pOneGem.mIndex, pOneGem.mType)
    ElseIf WhoCalled = 1 Then
        txt_Material.Text = AllGemMats_TEXT(TOON.REALM)
    End If
    
End Sub

Private Sub AssignGemCounts(nGem As Gem)

    'sum dust
    aDust(nGem.DUST_Index) = aDust(nGem.DUST_Index) + nGem.DUST_QUAN
    
    'sum liquids
    aLiquid(nGem.LIQUID1_Index) = aLiquid(nGem.LIQUID1_Index) + nGem.LIQUID1_QUAN
    If nGem.LIQUID2_QUAN > 0 Then
        aLiquid(nGem.LIQUID2_Index) = aLiquid(nGem.LIQUID2_Index) + nGem.LIQUID2_QUAN
        aLiquid(nGem.LIQUID3_Index) = aLiquid(nGem.LIQUID3_Index) + nGem.LIQUID3_QUAN
    End If
    
    'sum container gems
    aContainer(nGem.CONTAINER_Index) = aContainer(nGem.CONTAINER_Index) + nGem.CONTAINER_QUAN
    
    aTotal = aTotal + nGem.GEM_PRICE
    
End Sub

Private Sub SumGemLocation(Index As Integer, tRealm As Long)

    Dim oneGem As Gem
    
    If WS.lbl_GemNameSC1(Index).Caption <> vbNullString Then
        oneGem = GetGemMats(tRealm, WS.cmb_GemAmountSC1(Index).ListIndex, WS.cmb_GemSelectSC1(Index).Text, WS.lbl_GemNameSC1(Index).Caption)
        list_GemNames.AddItem WS.lbl_GemNameSC1(Index)
        oneGem.GEM_PRICE = GemInfo(2, WS.cmb_GemAmountSC1(Index).ListIndex) * Val(lbl_MultValue.Caption)
        Call AssignGemCounts(oneGem)
    End If
    
    If WS.lbl_GemNameSC2(Index).Caption <> vbNullString Then
        oneGem = GetGemMats(tRealm, WS.cmb_GemAmountSC2(Index).ListIndex, WS.cmb_GemSelectSC2(Index).Text, WS.lbl_GemNameSC2(Index).Caption)
        list_GemNames.AddItem WS.lbl_GemNameSC2(Index)
        oneGem.GEM_PRICE = GemInfo(2, WS.cmb_GemAmountSC2(Index).ListIndex) * Val(lbl_MultValue.Caption)
        Call AssignGemCounts(oneGem)
    End If
    
    If WS.lbl_GemNameSC3(Index).Caption <> vbNullString Then
        oneGem = GetGemMats(tRealm, WS.cmb_GemAmountSC3(Index).ListIndex, WS.cmb_GemSelectSC3(Index).Text, WS.lbl_GemNameSC3(Index).Caption)
        list_GemNames.AddItem WS.lbl_GemNameSC3(Index)
        oneGem.GEM_PRICE = GemInfo(2, WS.cmb_GemAmountSC3(Index).ListIndex) * Val(lbl_MultValue.Caption)
        Call AssignGemCounts(oneGem)
    End If
    
    If WS.lbl_GemNameSC4(Index).Caption <> vbNullString Then
        oneGem = GetGemMats(tRealm, WS.cmb_GemAmountSC4(Index).ListIndex, WS.cmb_GemSelectSC4(Index).Text, WS.lbl_GemNameSC4(Index).Caption)
        list_GemNames.AddItem WS.lbl_GemNameSC4(Index)
        oneGem.GEM_PRICE = GemInfo(2, WS.cmb_GemAmountSC4(Index).ListIndex) * Val(lbl_MultValue.Caption)
        Call AssignGemCounts(oneGem)
    End If
    
End Sub

Public Sub MatsDisplay_ONE(tRealm As Long, sGem As String, gIndex As Long, gType As String)

    Dim oneGem As Gem
    Dim sText As String
    Dim sBREAK As String
   
    txt_Material.Text = vbNullString
    
    'only one gem to be displayed and modified by the multiplier
    WhoCalled = 0
    
    pOneGem.mGem = sGem
    pOneGem.mRealm = tRealm
    pOneGem.mIndex = gIndex
    pOneGem.mType = gType

    txt_Material = OneGemMats(tRealm, sGem, gIndex, gType)
    Material_List.Show vbModal, WS
    
End Sub

Public Sub MatsDisplay_ALL(tRealm As Long)

    Dim oneGem As Gem
    Dim sText As String
   
    txt_Material.Text = vbNullString
    
    'all gems to be displayed and modified by the multiplier
    WhoCalled = 1
    
    txt_Material = AllGemMats_TEXT(tRealm)
    Material_List.Show vbModal, WS
    
End Sub

Private Function CountGemNames_HTML(list As ListBox, Multi As Long) As String

    Dim sBuffer As String
    Dim X As Long
    Dim Y As Long
    Dim Cnt As Long
    
    Dim sBR As String
    sBR = "<BR>"
    
    Dim sNBSP As String
    sNBSP = "&nbsp; &nbsp; &nbsp;"
    
    On Error Resume Next
    With list
        If .ListCount < 1 Then Exit Function
        If .Sorted = False Then Exit Function
        For X = 0 To (.ListCount - 1)
            If X > (.ListCount - 1) Then Exit For
            DoEvents
            Cnt = 1
            Y = (X + 1)
            If .list(Y) = .list(X) Then
                Do
                    .RemoveItem Y
                    Cnt = Cnt + 1
                Loop Until .list(Y) <> .list(X) Or (.list(Y) = vbNullString)
            End If
            
            If .list(X) <> vbNullString Or .ListCount = 1 Then
                sBuffer = sBuffer & sNBSP & sNBSP & Chr$(164) & " " & (Cnt * Multi) & " " & .list(X) & sBR
            End If
        Next X
    End With
    
    CountGemNames_HTML = sBuffer

End Function

Private Function CountGemNames_TEXT(list As ListBox, Multi As Long) As String

    Dim sBuffer As String
    Dim X As Long
    Dim Y As Long
    Dim Cnt As Long
    
    On Error Resume Next
    With list
        If .ListCount < 1 Then Exit Function
        If .Sorted = False Then Exit Function
        For X = 0 To (.ListCount - 1)
            If X > (.ListCount - 1) Then Exit For
            DoEvents
            Cnt = 1
            Y = (X + 1)
            If .list(Y) = .list(X) Then
                Do
                    .RemoveItem Y
                    Cnt = Cnt + 1
                Loop Until .list(Y) <> .list(X) Or (.list(Y) = vbNullString)
            End If
            
            If .list(X) <> vbNullString Or .ListCount = 1 Then
                sBuffer = sBuffer & vbTab & vbTab & Chr$(164) & " " & (Cnt * Multi) & " " & .list(X) & vbCrLf
            End If
        Next X
    End With
    
    CountGemNames_TEXT = sBuffer

End Function

Private Function AllGemMats_HTML(tRealm As Long)

    Dim Ctr As Long
    Dim sText As String
    Dim sBREAK As String
    
    Dim sBR As String
    sBR = "<BR>"
    
    Dim sNBSP As String
    sNBSP = "&nbsp; &nbsp; &nbsp;"
    
    sBREAK = "__________________________________________________________" & sBR
        
    Call Clear_DustLiquidContainer
    
    list_GemNames.Clear
    
    Call SumGemLocation(0, tRealm)
    Call SumGemLocation(1, tRealm)
    Call SumGemLocation(5, tRealm)
    Call SumGemLocation(6, tRealm)
    Call SumGemLocation(7, tRealm)
    Call SumGemLocation(8, tRealm)
    Call SumGemLocation(9, tRealm)
    Call SumGemLocation(10, tRealm)
    Call SumGemLocation(14, tRealm)
    Call SumGemLocation(15, tRealm)
    Call SumGemLocation(18, tRealm)
    Call SumGemLocation(19, tRealm)
    Call SumGemLocation(20, tRealm)
    Call SumGemLocation(21, tRealm)
    
    sText = "Total Cost: " & FormatCost(aTotal) & sBR
    sText = sText & sBREAK & sBR
    
    'get gem name counts
    sText = sText & "Gem Name" & sBR & sBR & _
                                CountGemNames_HTML(list_GemNames, Val(lbl_MultValue.Caption))
    
    sText = sText & sBREAK & sBR
    
    sText = sText & "Materials" & sBR & sBR
    sText = sText & sNBSP & "Gems:" & sBR
    For Ctr = 0 To 9
        If aContainer(Ctr) > 0 Then
            sText = sText & sNBSP & sNBSP & Chr$(164) & " " & Val(lbl_MultValue.Caption) * aContainer(Ctr) & " " & GemInfo(0, Ctr) & sBR
        End If
    Next Ctr
    
    sText = sText & sNBSP & "Liquid:" & sBR
    
    For Ctr = 0 To 10
        If aLiquid(Ctr) > 0 Then
            sText = sText & sNBSP & sNBSP & Chr$(164) & " " & Val(lbl_MultValue.Caption) * aLiquid(Ctr) & " " & GemMaterial_Liquid(Ctr) & sBR
        End If
    Next Ctr
    
    sText = sText & sNBSP & "Dust:" & sBR
    
    For Ctr = 0 To 11
        If aDust(Ctr) > 0 Then
            sText = sText & sNBSP & sNBSP & Chr$(164) & " " & Val(lbl_MultValue.Caption) * aDust(Ctr) & " " & GemMaterial_Dust(Ctr) & sBR
        End If
    Next Ctr
    
    sText = sText & sBREAK
    
    AllGemMats_HTML = sText

End Function

Private Function AllGemMats_TEXT(tRealm As Long)

    Dim Ctr As Long
    Dim sText As String
    Dim sBREAK As String
    
    sBREAK = "__________________________________________________________" & vbCrLf
    
    Call Clear_DustLiquidContainer
    
    list_GemNames.Clear
    
    Call SumGemLocation(0, tRealm)
    Call SumGemLocation(1, tRealm)
    Call SumGemLocation(5, tRealm)
    Call SumGemLocation(6, tRealm)
    Call SumGemLocation(7, tRealm)
    Call SumGemLocation(8, tRealm)
    Call SumGemLocation(9, tRealm)
    Call SumGemLocation(10, tRealm)
    Call SumGemLocation(14, tRealm)
    Call SumGemLocation(15, tRealm)
    Call SumGemLocation(18, tRealm)
    Call SumGemLocation(19, tRealm)
    Call SumGemLocation(20, tRealm)
    Call SumGemLocation(21, tRealm)
    
    sText = "Total Cost: " & FormatCost(aTotal) & vbCrLf
    sText = sText & sBREAK & vbCrLf
    
    'get gem name counts
    sText = sText & "Gem Name" & vbCrLf & vbCrLf & _
                                CountGemNames_TEXT(list_GemNames, Val(lbl_MultValue.Caption))
    
    sText = sText & sBREAK & vbCrLf
    
    sText = sText & "Materials" & vbCrLf & vbCrLf
    sText = sText & vbTab & "Gems:" & vbCrLf
    For Ctr = 0 To 9
        If aContainer(Ctr) > 0 Then
            sText = sText & vbTab & vbTab & Chr$(164) & " " & Val(lbl_MultValue.Caption) * aContainer(Ctr) & " " & GemInfo(0, Ctr) & vbCrLf
        End If
    Next Ctr
    
    sText = sText & vbTab & "Liquid:" & vbCrLf
    
    For Ctr = 0 To 10
        If aLiquid(Ctr) > 0 Then
            sText = sText & vbTab & vbTab & Chr$(164) & " " & Val(lbl_MultValue.Caption) * aLiquid(Ctr) & " " & GemMaterial_Liquid(Ctr) & vbCrLf
        End If
    Next Ctr
    
    sText = sText & vbTab & "Dust:" & vbCrLf
    
    For Ctr = 0 To 11
        If aDust(Ctr) > 0 Then
            sText = sText & vbTab & vbTab & Chr$(164) & " " & Val(lbl_MultValue.Caption) * aDust(Ctr) & " " & GemMaterial_Dust(Ctr) & vbCrLf
        End If
    Next Ctr
    
    sText = sText & sBREAK
    
    AllGemMats_TEXT = sText

End Function

Private Function OneGemMats(tRealm As Long, sGem As String, gIndex As Long, gType As String) As String
    
    Dim oneGem As Gem
    Dim sText As String
    Dim sBREAK As String
    Dim lPrice As Long
    
    Dim sBR As String
    sBR = "<BR>"
    
    Dim sNBSP As String
    sNBSP = "&nbsp; &nbsp; &nbsp;"
    
    sBREAK = "__________________________________________________________" & vbCrLf
    oneGem = GetGemMats(tRealm, gIndex, gType, sGem)
    
    lPrice = GemInfo(2, gIndex) * Val(lbl_MultValue.Caption)
    sText = "Total Cost: " & FormatCost(lPrice) & vbCrLf
    
    sText = sText & sBREAK & vbCrLf
    
    sText = sText & "Gem Name" & vbCrLf & vbCrLf & vbTab & vbTab & _
                        Chr$(164) & " " & lbl_MultValue.Caption & " " & sGem & vbCrLf
    
    sText = sText & sBREAK & vbCrLf
    
    sText = sText & "Materials" & vbCrLf & vbCrLf
    sText = sText & vbTab & "Gems:" & vbCrLf & vbTab & vbTab & _
                            Chr$(164) & " " & Val(lbl_MultValue.Caption) * oneGem.CONTAINER_QUAN & _
                                        " " & oneGem.CONTAINER_NAME & vbCrLf
                                        
    sText = sText & vbTab & "Liquid:" & vbCrLf & vbTab & vbTab & _
                            Chr$(164) & " " & Val(lbl_MultValue.Caption) * oneGem.LIQUID1_QUAN & _
                                        " " & oneGem.LIQUID1_NAME & vbCrLf
    If oneGem.LIQUID2_QUAN <> 0 Then
        sText = sText & vbTab & vbTab & Chr$(164) & _
                    " " & Val(lbl_MultValue.Caption) * oneGem.LIQUID2_QUAN & " " & _
                            oneGem.LIQUID2_NAME & vbCrLf
                            
        sText = sText & vbTab & vbTab & Chr$(164) & _
                    " " & Val(lbl_MultValue.Caption) * oneGem.LIQUID3_QUAN & " " & _
                            oneGem.LIQUID3_NAME & vbCrLf
    End If
    sText = sText & vbTab & "Dust:" & vbCrLf & vbTab & vbTab & Chr$(164) & _
                            " " & Val(lbl_MultValue.Caption) * oneGem.DUST_QUAN & " " & _
                                    oneGem.DUST_NAME & vbCrLf & vbCrLf
    sText = sText & sBREAK
    
    
    sGemReport_HTML = "Total Cost: " & FormatCost(lPrice) & sBR & "__________________________________________________________" & sBR & sBR
    sGemReport_HTML = sGemReport_HTML & "Gem Name" & sBR & sBR & sNBSP & sNBSP & Chr$(164) & _
                    " " & lbl_MultValue.Caption & " " & sGem & sBR
                    
    sGemReport_HTML = sGemReport_HTML & "__________________________________________________________" & sBR & sBR
    
    sGemReport_HTML = sGemReport_HTML & "Materials" & sBR & sBR & sNBSP & sNBSP & "Gems:" & sBR & sNBSP & sNBSP & _
                         sNBSP & Chr$(164) & " " & Val(lbl_MultValue.Caption) * oneGem.CONTAINER_QUAN & _
                                    " " & oneGem.CONTAINER_NAME & sBR
    sGemReport_HTML = sGemReport_HTML & sNBSP & sNBSP & "Liquid:" & sBR & sNBSP & sNBSP & sNBSP & Chr$(164) & " " & _
                    Val(lbl_MultValue.Caption) * oneGem.LIQUID1_QUAN & " " & oneGem.LIQUID1_NAME & sBR
                    
    If oneGem.LIQUID2_QUAN <> 0 Then
        sGemReport_HTML = sGemReport_HTML & sNBSP & sNBSP & sNBSP & Chr$(164) & _
                        " " & Val(lbl_MultValue.Caption) * oneGem.LIQUID2_QUAN & " " & _
                            oneGem.LIQUID2_NAME & sNBSP & sBR
        sGemReport_HTML = sGemReport_HTML & sNBSP & sNBSP & sNBSP & Chr$(164) & _
                        " " & Val(lbl_MultValue.Caption) * oneGem.LIQUID3_QUAN & " " & _
                            oneGem.LIQUID3_NAME & sNBSP & sBR
    End If
    
    sGemReport_HTML = sGemReport_HTML & sNBSP & sNBSP & "Dust:" & sBR & sNBSP & sNBSP & sNBSP & Chr$(164) & _
                        " " & Val(lbl_MultValue.Caption) * oneGem.DUST_QUAN & " " & _
                                oneGem.DUST_NAME & sBR & sBR
    
    sGemReport_HTML = sGemReport_HTML & "__________________________________________________________" & sBR
    
    OneGemMats = sText
    
End Function

Private Sub mnuClose_Click()

    Unload Me
    
End Sub

Private Sub mnuPrint_Click()

    On Error Resume Next
    
    Printer.Print txt_Material
    Printer.EndDoc
    
End Sub

Private Sub mnuSaveHTML_Click()

    Dim hFile As Long
    Dim sBuffer As String
    Dim sPath As String
    Dim lRet As Long
    
    Dim lWrite As Long
    Dim cmdFlags As Long
    Dim cmdFilter As String
    Dim cmdMessage As String
    
    
    cmdFlags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    cmdFilter = "HTML Files (*.html)" & vbNullChar & "*.html" & vbNullChar
    cmdMessage = "Save Configuration as HTML"
      
    sPath = CMD_OpenSave(lSave, Me.hwnd, cmdFilter, 1, App.Path, cmdMessage, cmdFlags)
    
    If sPath <> vbNullString Then
        
        If LCase$(Mid(sPath, Len(sPath) - 4)) <> ".html" Then sPath = sPath & ".html"
        
        hFile = CreateFile(sPath, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_WRITE Or FILE_SHARE_READ, ByVal CLng(0), CREATE_ALWAYS, FILE_ATTRIBUTE_ARCHIVE, 0)
        
        If hFile = -1 Then Exit Sub '*If hFile is -1 the file is not there and there has been an error
        
        sBuffer = "<html><br><title>Material Listing</title><br><body><br><font face=verdana size=2 color=000088><br>"
        If WhoCalled = 0 Then
            sBuffer = sBuffer & sGemReport_HTML
        ElseIf WhoCalled = 1 Then
            sBuffer = sBuffer & AllGemMats_HTML(TOON.REALM)
        End If
        
        sBuffer = sBuffer & "<br></font><br></body><br></html>"
        
        lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
        lRet = CloseHandle(hFile)
    End If
    
End Sub

Private Sub mnuSaveText_Click()

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
        
        sBuffer = txt_Material.Text
        lRet = WriteFile(hFile, ByVal sBuffer, Len(sBuffer), lWrite, ByVal CLng(0))
        lRet = CloseHandle(hFile)
    End If
    
End Sub
