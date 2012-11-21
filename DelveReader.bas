Attribute VB_Name = "DelveReader"
Option Explicit

Private Type ITEM

    NAME As String
    nSLOT(12) As String
    vSLOT(12) As Long
    
    TYPE As Long    '0 - crafted
                    '1 - rog
                    '2 - artifact
                    '3 - drop
                    
    F_POWER_POOL As Boolean
    
End Type

Private sItem As ITEM

Private Sub DebugItem()
    Dim X As Long
    
    For X = 0 To 12
        Debug.Print sItem.nSLOT(X)
    Next X
    
End Sub

Private Function FindItem(sBuffer As String, sArray, iIndex As Long) As Long
    
    Dim iCtr As Long, lResult As Long
    
    lResult = -1 'default return value, means failure
    
    For iCtr = 0 To iIndex
        If InStr(sBuffer, LCase(sArray(iCtr))) <> 0 Then
            lResult = iCtr
            Exit For
        End If
        If sArray(iCtr) = vbNullString Then Exit For
    Next iCtr
    
    FindItem = lResult
    
End Function

Public Function GetItemFromDelve(sFILE As String) As String

Dim bGotName As Boolean, bGotType As Boolean    'bools for parse stuff
Dim lRet As Long, lPos As Long, iCtr As Long    'return value for calls
                                                'character position
                                                'counter index variable
Dim sBuffer As String   'string buffer

bGotName = False    'name bool init to false
bGotType = False    'type bool init to false

    sFILE = "e:\games\mythic\catacombs\chat.log"
    
    Open sFILE For Input As #1  'Open file for input.
    
        Do While Not EOF(1)     'Check for end of file.
        
            Line Input #1, sBuffer  'Read line of data.
            
            sBuffer = LCase(sBuffer)
            lPos = InStr(sBuffer, "<begin info:")   'Check for Item Identifier
            If lPos <> 0 Then
                 
                sItem.NAME = Mid(sBuffer, 25, Len(sBuffer) - 26)    'get items name
                bGotName = True
                Do
                    '-----------------------------get item type--------------------------
                    If bGotType = False Then
                        Do
                            Line Input #1, sBuffer
                            sBuffer = LCase(sBuffer)
                            
                            If InStr(sBuffer, "crafted by:") = 0 Then
                                If InStr(sBuffer, "unique object") = 0 Then
                                    Line Input #1, sBuffer
                                    sBuffer = LCase(sBuffer)
                                Else
                                    sItem.TYPE = 1  'rog
                                    bGotType = True
                                End If
                            Else
                                sItem.TYPE = 0  'crafted
                                bGotType = True
                            End If
                            
                            If InStr(sBuffer, "artifact:") = 0 Then
                                sItem.TYPE = 3  'drop
                                bGotType = True
                            Else
                                sItem.TYPE = 2  'arti
                                bGotType = True
                            End If
                        Loop Until bGotType = True
                    End If
                    '--------------------------------------------------------------------
                    
                    '-----------------------------find magical bonuses-------------------
                    Do
                        Line Input #1, sBuffer
                        sBuffer = LCase(sBuffer)
                    Loop Until InStr(sBuffer, "magical bonuses:") <> 0
                    
                    iCtr = 0    'initialize ictr
                    
                    Do
                        Line Input #1, sBuffer
                        sBuffer = LCase(sBuffer)
                        
                        If Len(Trim(sBuffer)) = 10 Then Exit Do
                        
                        If InStr(sBuffer, "require") = 0 Then
                            If InStr(sBuffer, "hits:") <> 0 Then
                                sItem.nSLOT(iCtr) = "Hits"
                                iCtr = iCtr + 1
                            ElseIf InStr(sBuffer, "power:") <> 0 Then
                                If InStr(sBuffer, "pool") <> 0 Then
                                    sItem.F_POWER_POOL = True
                                Else
                                    sItem.F_POWER_POOL = False
                                End If
                                sItem.nSLOT(iCtr) = "Power"
                                iCtr = iCtr + 1
                            Else
                                lRet = FindItem(sBuffer, nStat(), 7)
                        
                                If lRet <> -1 Then  'try to find a stat
                                    sItem.nSLOT(iCtr) = nStat(lRet)
                                    iCtr = iCtr + 1
                                Else    'didn't find a stat
                                    lRet = FindItem(sBuffer, nResist(), 8)  'try to find a resist
                                    If lRet <> -1 Then
                                        sItem.nSLOT(iCtr) = nResist(lRet)
                                        iCtr = iCtr + 1
                                    Else    'didn't find resist so
                                        lRet = FindItem(sBuffer, nSkill, 36)    'try to find a skill
                                        If lRet <> -1 Then
                                            sItem.nSLOT(iCtr) = nSkill(lRet)
                                            iCtr = iCtr + 1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Loop
                    '--------------------------------------------------------------------
                    DebugItem
                    '-----------------------------find focus bonus-----------------------
                    Line Input #1, sBuffer
                    sBuffer = LCase(sBuffer)
                    
                    If InStr(sBuffer, "focus") <> 0 Then
                        Do
                            Line Input #1, sBuffer
                            sBuffer = LCase(sBuffer)
                            If Len(Trim(sBuffer)) = 10 Then Exit Do
                            lRet = FindItem(sBuffer, nFocus(), 12)
                            
                            If lRet <> -1 Then  'try to find focus type
                                sItem.nSLOT(iCtr) = nFocus(lRet)
                                iCtr = iCtr + 1
                            End If
                        Loop
                    DebugItem
                    End If
                            
                            
                            
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                Loop While InStr(sBuffer, "<end info>") = 0
            End If
        Loop
        
    Close #1    'Close file.
    
End Function

Public Function GetItemFromChat(sFILE As String) As String


End Function

